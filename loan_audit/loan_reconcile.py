# -*- coding: utf-8 -*-
"""贷款收款 / 余额勾稽引擎（pyexcel 原生 Book 输出）。

口径（固定，不随跨月、部分还款或结清后多收而改变）：
  * 金额一律为整数 ``分``；年利率一律为整数 ``基点``（1 基点 = 0.01%）。
  * 计息：实际天数 / 365（act/365），按日、按币种独立四舍五入（ROUND_HALF_UP）。
  * 事件顺序：同一自然日先生效利率变更，再按业务序号逐笔处理收款，
    日终若仍有剩余本金才对剩余本金计一天利息；本金为零后不再计息。
  * 收款分配：先还已欠利息（accrued unpaid interest），再还本金；
    超过“已欠利息 + 剩余本金”的部分单列为余款（overpayment），
    不作为任何后续贷项、不抵扣其它贷款；未付利息不资本化、不计复利。
  * 放款日当日即起息（日终按放款本金计一天利息）。

用法：
    python3 -I loan_reconcile.py <输入工作簿> <输出工作簿> [截止日 YYYY-MM-DD]

输入工作簿必须含三张 Sheet：当前贷款 / 利率变更 / 收款（列名按各 Sheet 实际表头）。
来源文件与任何已存在的旧结果文件都不会被修改；输出路径已存在时直接报错退出。
任何数据问题（重复身份、坏日期、悬空引用、重复利率/序号等）都会整次报错，
不产生部分结果文件。
"""

from __future__ import annotations

import datetime
import math
import os
import re
import sys
from collections import OrderedDict, defaultdict

import pyexcel

# ---------------------------------------------------------------------------
# 常量与表头别名（按 Sheet 当前实际列名取数；别名仅做去空格/大小写归一）
# ---------------------------------------------------------------------------

LOAN_SHEET_ALIASES = ("当前贷款", "贷款", "loans", "loan")
RATE_SHEET_ALIASES = ("利率变更", "利率调整", "rate_changes", "rates", "rate")
PAY_SHEET_ALIASES = ("收款", "收款明细", "payments", "payment")

LOAN_COLUMNS = OrderedDict(
    [
        ("loan_id", ("贷款编号", "贷款id", "loan_id", "id")),
        ("borrower", ("借款人", "客户", "borrower", "name")),
        ("currency", ("币种", "货币", "currency", "ccy")),
        ("start_date", ("放款日", "起息日", "起贷日", "disbursement_date",
                        "start_date")),
        ("principal", ("本金_分", "本金(分)", "本金（分）", "本金分",
                       "principal_cents", "principal")),
        ("rate_bps", ("年利率_基点", "年利率(基点)", "年利率（基点）",
                      "利率_基点", "rate_bps", "annual_rate_bps")),
    ]
)

RATE_COLUMNS = OrderedDict(
    [
        ("loan_id", ("贷款编号", "贷款id", "loan_id", "id")),
        ("eff_date", ("生效日", "变更日", "effective_date", "change_date")),
        ("rate_bps", ("新年利率_基点", "新年利率(基点)", "新年利率（基点）",
                      "新利率_基点", "new_rate_bps", "rate_bps")),
    ]
)

PAY_COLUMNS = OrderedDict(
    [
        ("loan_id", ("贷款编号", "贷款id", "loan_id", "id")),
        ("pay_date", ("收款日", "到账日", "payment_date", "date")),
        ("amount", ("金额_分", "收款金额_分", "收款金额(分)", "收款金额（分）",
                    "amount_cents", "amount")),
        ("seq", ("业务序号", "序号", "顺序", "seq", "sequence")),
    ]
)

DAYS_PER_YEAR = 365
MAX_RATE_BPS = 1_000_000  # 10000%，合理性上限，超出视为坏数据

# 输出 Sheet 名（固定，便于下游勾稽）
OUT_PARAM = "参数与口径"
OUT_ALLOC = "逐笔分配"
OUT_BALANCE = "截止日余额"
OUT_LEDGER = "逐日台账"
OUT_MONTHLY = "逐月勾稽"
OUT_CCY_MONTHLY = "币种月度汇总"
OUT_SKIPPED = "异常与跳过"

RESULT_SHEET_NAMES = (
    OUT_PARAM, OUT_ALLOC, OUT_BALANCE, OUT_LEDGER,
    OUT_MONTHLY, OUT_CCY_MONTHLY, OUT_SKIPPED,
)


class DataError(ValueError):
    """输入数据问题：整次报错，不写结果。"""


# ---------------------------------------------------------------------------
# 基础工具
# ---------------------------------------------------------------------------

def _norm(text):
    """表头归一：去首尾及内部空白、转小写，便于匹配别名。"""
    if text is None:
        return ""
    return re.sub(r"\s+", "", str(text)).strip().lower()


_ISO_RE = re.compile(r"^(\d{4})-(\d{1,2})-(\d{1,2})$")


def parse_date(value, sheet, row_no, col_name):
    """严格解析日期：date/datetime 或 ISO 字符串；坏日期整次报错。"""
    if isinstance(value, datetime.datetime):
        return value.date()
    if isinstance(value, datetime.date):
        return value
    if isinstance(value, bool):  # bool 是 int 的子类，必须先挡
        raise DataError(
            f"{sheet} 第 {row_no} 行列「{col_name}」日期不能是布尔值: {value!r}")
    if isinstance(value, (int, float)):
        raise DataError(
            f"{sheet} 第 {row_no} 行列「{col_name}」必须是日期(YYYY-MM-DD)，"
            f"收到数字 {value!r}（请在来源表中改为真实日期）")
    if value is None or (isinstance(value, str) and not value.strip()):
        raise DataError(
            f"{sheet} 第 {row_no} 行列「{col_name}」日期为空")
    text = str(value).strip()
    m = _ISO_RE.match(text)
    if not m:
        raise DataError(
            f"{sheet} 第 {row_no} 行列「{col_name}」日期无法解析: {value!r}，"
            f"要求 YYYY-MM-DD")
    try:
        return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    except ValueError as exc:
        raise DataError(
            f"{sheet} 第 {row_no} 行列「{col_name}」坏日期 {value!r}: {exc}")


def parse_int(value, sheet, row_no, col_name, *, positive=False,
              non_negative=False, max_value=None):
    """严格解析整数分/基点/序号：拒绝浮点、千分位、空白、布尔。"""
    if isinstance(value, bool):
        raise DataError(
            f"{sheet} 第 {row_no} 行列「{col_name}」不能是布尔值: {value!r}")
    if isinstance(value, int):
        result = value
    elif isinstance(value, float):
        if not math.isfinite(value) or not value.is_integer():
            raise DataError(
                f"{sheet} 第 {row_no} 行列「{col_name}」必须是整数，"
                f"收到小数 {value!r}")
        result = int(value)
    elif value is None or (isinstance(value, str) and not value.strip()):
        raise DataError(
            f"{sheet} 第 {row_no} 行列「{col_name}」数值为空")
    else:
        text = str(value).strip().replace(",", "")
        if not re.match(r"^[+-]?\d+$", text):
            raise DataError(
                f"{sheet} 第 {row_no} 行列「{col_name}」不是整数: {value!r}")
        result = int(text)
    if positive and result <= 0:
        raise DataError(
            f"{sheet} 第 {row_no} 行列「{col_name}」必须为正数，收到 {result}")
    if non_negative and result < 0:
        raise DataError(
            f"{sheet} 第 {row_no} 行列「{col_name}」不能为负，收到 {result}")
    if max_value is not None and result > max_value:
        raise DataError(
            f"{sheet} 第 {row_no} 行列「{col_name}」超出合理上限 "
            f"{max_value}: {result}")
    return result


def parse_text(value, sheet, row_no, col_name):
    if value is None or (isinstance(value, str) and not value.strip()):
        raise DataError(
            f"{sheet} 第 {row_no} 行列「{col_name}」文本为空")
    return str(value).strip()


def round_half_up(numerator, denominator):
    """整数域四舍五入（ROUND_HALF_UP），与 Decimal ROUND_HALF_UP 一致。

    用于每日利息：principal * bps / (10000 * 365) 的逐币种独立取整。
    分母为正，调用方保证 numerator >= 0。
    """
    q, r = divmod(numerator, denominator)
    return q + (1 if 2 * r >= denominator else 0)


def month_key(d):
    return (d.year, d.month)


def month_end(d):
    if d.month == 12:
        return d.replace(day=31)
    first_next = d.replace(day=1) + datetime.timedelta(days=32)
    return first_next.replace(day=1) - datetime.timedelta(days=1)


# ---------------------------------------------------------------------------
# Sheet 读取与表头映射
# ---------------------------------------------------------------------------

def resolve_sheet(book, aliases, purpose):
    names = book.sheet_names()
    for alias in aliases:
        if alias in names:
            return book.sheet_by_name(alias), alias
    norm_map = {_norm(n): n for n in names}
    for alias in aliases:
        if _norm(alias) in norm_map:
            real = norm_map[_norm(alias)]
            return book.sheet_by_name(real), real
    raise DataError(
        f"缺少{purpose} Sheet（可接受名称: {'/'.join(aliases)}），"
        f"工作簿现有 Sheet: {names}")


def map_headers(sheet, columns, sheet_name):
    """按当前实际表头建立 规范字段 -> 列下标 映射；缺列/重名整次报错。"""
    array = sheet.to_array()
    header_idx = None
    for i, row in enumerate(array):
        if any(c is not None and str(c).strip() != "" for c in row):
            header_idx = i
            break
    if header_idx is None:
        raise DataError(f"Sheet「{sheet_name}」为空，没有表头行")
    header_row = array[header_idx]
    seen_norm = defaultdict(list)
    for j, cell in enumerate(header_row):
        key = _norm(cell)
        if key:
            seen_norm[key].append(j)
    mapping = {}
    for field, aliases in columns.items():
        hits = []
        for alias in aliases:
            hits.extend(seen_norm.get(_norm(alias), []))
        hits = sorted(set(hits))
        if not hits:
            raise DataError(
                f"Sheet「{sheet_name}」缺少必需列「{aliases[0]}」，"
                f"当前表头: {[c for c in header_row if str(c).strip()]}")
        if len(hits) > 1:
            raise DataError(
                f"Sheet「{sheet_name}」列「{aliases[0]}」在表头中重复出现，"
                f"列下标 {hits}")
        mapping[field] = hits[0]
    return array, header_idx, mapping


def iter_data_rows(array, header_idx, sheet_name):
    """逐行产出 (真实行号, 行)；全空行跳过，短行按 None 补齐。"""
    width = max(len(r) for r in array)
    for i in range(header_idx + 1, len(array)):
        row = list(array[i]) + [None] * (width - len(array[i]))
        if all(c is None or (isinstance(c, str) and not c.strip())
               for c in row):
            continue
        # 电子表格行号从 1 起；header_idx 是 0 基下标
        yield header_idx + 1 + (i - header_idx), row


def cell(row, mapping, field):
    return row[mapping[field]]


# ---------------------------------------------------------------------------
# 输入装载与校验
# ---------------------------------------------------------------------------

def load_inputs(book):
    loan_sheet, loan_sheet_name = resolve_sheet(
        book, LOAN_SHEET_ALIASES, "当前贷款")
    rate_sheet, rate_sheet_name = resolve_sheet(
        book, RATE_SHEET_ALIASES, "利率变更")
    pay_sheet, pay_sheet_name = resolve_sheet(
        book, PAY_SHEET_ALIASES, "收款")

    loans = OrderedDict()
    arr, hidx, mp = map_headers(loan_sheet, LOAN_COLUMNS, loan_sheet_name)
    for row_no, row in iter_data_rows(arr, hidx, loan_sheet_name):
        lid = parse_text(cell(row, mp, "loan_id"),
                         loan_sheet_name, row_no, "贷款编号")
        if lid in loans:
            raise DataError(
                f"Sheet「{loan_sheet_name}」贷款编号重复: {lid!r}（"
                f"首次出现于第 {loans[lid]['row_no']} 行，再次出现于第 "
                f"{row_no} 行）")
        borrower = parse_text(cell(row, mp, "borrower"),
                              loan_sheet_name, row_no, "借款人")
        currency = parse_text(cell(row, mp, "currency"),
                              loan_sheet_name, row_no, "币种").upper()
        start = parse_date(cell(row, mp, "start_date"),
                           loan_sheet_name, row_no, "放款日")
        principal = parse_int(cell(row, mp, "principal"),
                              loan_sheet_name, row_no, "本金_分",
                              positive=True)
        rate = parse_int(cell(row, mp, "rate_bps"),
                         loan_sheet_name, row_no, "年利率_基点",
                         non_negative=True, max_value=MAX_RATE_BPS)
        loans[lid] = {
            "loan_id": lid, "borrower": borrower, "currency": currency,
            "start_date": start, "principal": principal, "rate_bps": rate,
            "row_no": row_no,
        }

    rate_events = defaultdict(dict)   # loan_id -> {date: bps}
    arr, hidx, mp = map_headers(rate_sheet, RATE_COLUMNS, rate_sheet_name)
    for row_no, row in iter_data_rows(arr, hidx, rate_sheet_name):
        lid = parse_text(cell(row, mp, "loan_id"),
                         rate_sheet_name, row_no, "贷款编号")
        if lid not in loans:
            raise DataError(
                f"Sheet「{rate_sheet_name}」第 {row_no} 行贷款编号 {lid!r} "
                f"在当前贷款表中不存在（悬空引用）")
        eff = parse_date(cell(row, mp, "eff_date"),
                         rate_sheet_name, row_no, "生效日")
        if eff < loans[lid]["start_date"]:
            raise DataError(
                f"Sheet「{rate_sheet_name}」第 {row_no} 行贷款 {lid} 的利率"
                f"生效日 {eff.isoformat()} 早于放款日 "
                f"{loans[lid]['start_date'].isoformat()}")
        bps = parse_int(cell(row, mp, "rate_bps"),
                        rate_sheet_name, row_no, "新年利率_基点",
                        non_negative=True, max_value=MAX_RATE_BPS)
        if eff in rate_events[lid]:
            raise DataError(
                f"Sheet「{rate_sheet_name}」贷款 {lid} 在 {eff.isoformat()} "
                f"存在重复利率变更（同日只允许一条）")
        rate_events[lid][eff] = bps

    payments = []
    seen_pay_key = set()
    arr, hidx, mp = map_headers(pay_sheet, PAY_COLUMNS, pay_sheet_name)
    for row_no, row in iter_data_rows(arr, hidx, pay_sheet_name):
        lid = parse_text(cell(row, mp, "loan_id"),
                         pay_sheet_name, row_no, "贷款编号")
        if lid not in loans:
            raise DataError(
                f"Sheet「{pay_sheet_name}」第 {row_no} 行贷款编号 {lid!r} "
                f"在当前贷款表中不存在（悬空引用）")
        pdate = parse_date(cell(row, mp, "pay_date"),
                           pay_sheet_name, row_no, "收款日")
        if pdate < loans[lid]["start_date"]:
            raise DataError(
                f"Sheet「{pay_sheet_name}」第 {row_no} 行贷款 {lid} 的收款日 "
                f"{pdate.isoformat()} 早于放款日 "
                f"{loans[lid]['start_date'].isoformat()}")
        amount = parse_int(cell(row, mp, "amount"),
                           pay_sheet_name, row_no, "金额_分",
                           positive=True)
        seq = parse_int(cell(row, mp, "seq"),
                        pay_sheet_name, row_no, "业务序号", positive=True)
        key = (lid, pdate, seq)
        if key in seen_pay_key:
            raise DataError(
                f"Sheet「{pay_sheet_name}」贷款 {lid} 在 {pdate.isoformat()} "
                f"业务序号 {seq} 重复")
        seen_pay_key.add(key)
        payments.append({"loan_id": lid, "date": pdate, "amount": amount,
                         "seq": seq, "row_no": row_no})

    return loans, rate_events, payments, (loan_sheet_name,
                                          rate_sheet_name,
                                          pay_sheet_name)


# ---------------------------------------------------------------------------
# 单贷款模拟
# ---------------------------------------------------------------------------

INTEREST_DENOM = 10000 * DAYS_PER_YEAR  # bps -> 日利率分母


def simulate_loan(loan, rate_map, payment_list, cutoff):
    """返回 (逐日台账行, 逐笔分配行, 截止日余额 dict)。

    rate_map: {生效日: 新利率基点}；payment_list: 该贷款全部收款（已排序视图）。
    截止日之后的事件由调用方剔除并登记，不进入模拟。
    """
    lid = loan["loan_id"]
    start = loan["start_date"]
    if start > cutoff:
        raise DataError(
            f"贷款 {lid} 放款日 {start.isoformat()} 晚于截止日 "
            f"{cutoff.isoformat()}，无法在该截止日核对")

    pays_by_day = defaultdict(list)
    for p in payment_list:
        pays_by_day[p["date"]].append(p)
    for d in pays_by_day:
        pays_by_day[d].sort(key=lambda x: x["seq"])

    principal = loan["principal"]
    accrued = 0                 # 已欠未付利息（分），永不资本化
    overpayment = 0             # 累计余款（分），不作后续贷项
    active_rate = loan["rate_bps"]

    ledger_rows = []
    alloc_rows = []
    day = start
    one_day = datetime.timedelta(days=1)
    while day <= cutoff:
        # 1) 同日先换利率（日初生效；放款日当天的变更也先于收款）
        if day in rate_map:
            active_rate = rate_map[day]
        # 2) 再按业务序号收款：先利息、后本金、余款单列
        day_cash = day_int = day_prin = day_over = 0
        for p in pays_by_day.get(day, ()):
            cash = p["amount"]
            to_int = min(cash, accrued)
            accrued -= to_int
            remain = cash - to_int
            to_prin = min(remain, principal)
            principal -= to_prin
            remain -= to_prin
            to_over = remain  # 结清后多收：仅挂账，不回流
            overpayment += to_over
            day_cash += cash
            day_int += to_int
            day_prin += to_prin
            day_over += to_over
            alloc_rows.append(OrderedDict([
                ("贷款编号", lid),
                ("币种", loan["currency"]),
                ("收款日", day.isoformat()),
                ("业务序号", p["seq"]),
                ("收款金额_分", cash),
                ("还利息_分", to_int),
                ("还本金_分", to_prin),
                ("余款_分", to_over),
                ("收款后剩余本金_分", principal),
                ("收款后已欠利息_分", accrued),
                ("来源行号", p["row_no"]),
            ]))
        # 3) 日终：仍有剩余本金才计息（本金为零不再产生利息）
        day_accrual = 0
        if principal > 0:
            day_accrual = round_half_up(
                principal * active_rate, INTEREST_DENOM)
            accrued += day_accrual
        ledger_rows.append(OrderedDict([
            ("贷款编号", lid),
            ("币种", loan["currency"]),
            ("日期", day.isoformat()),
            ("当日利率_基点", active_rate),
            ("当日收款_分", day_cash),
            ("当日还利息_分", day_int),
            ("当日还本金_分", day_prin),
            ("当日余款_分", day_over),
            ("当日计息_分", day_accrual),
            ("日终剩余本金_分", principal),
            ("日终已欠利息_分", accrued),
            ("日终余款累计_分", overpayment),
        ]))
        day += one_day

    balance = OrderedDict([
        ("贷款编号", lid),
        ("借款人", loan["borrower"]),
        ("币种", loan["currency"]),
        ("放款日", start.isoformat()),
        ("截止日", cutoff.isoformat()),
        ("原始本金_分", loan["principal"]),
        ("截止日剩余本金_分", principal),
        ("截止日已欠利息_分", accrued),
        ("截止日余款累计_分", overpayment),
        ("截止日应收合计_分", principal + accrued),
        ("当前年利率_基点", active_rate),
    ])
    return ledger_rows, alloc_rows, balance


# ---------------------------------------------------------------------------
# 逐月勾稽（本金 / 利息 / 现金分配全部可滚动核对）
# ---------------------------------------------------------------------------

def build_monthly(loan, ledger_rows, alloc_rows):
    """以逐日台账为唯一来源滚动生成月度行，口径与部分还款/结清无关。"""
    lid = loan["loan_id"]
    ccy = loan["currency"]
    if not ledger_rows:
        return []
    start = datetime.date.fromisoformat(ledger_rows[0]["日期"])
    end = datetime.date.fromisoformat(ledger_rows[-1]["日期"])

    # 日 -> 台账行
    by_day = {r["日期"]: r for r in ledger_rows}
    # 月度聚合
    agg = defaultdict(lambda: {"accrue": 0, "cash": 0, "int": 0,
                               "prin": 0, "over": 0})
    for r in ledger_rows:
        k = month_key(datetime.date.fromisoformat(r["日期"]))
        a = agg[k]
        a["accrue"] += r["当日计息_分"]
        a["cash"] += r["当日收款_分"]
        a["int"] += r["当日还利息_分"]
        a["prin"] += r["当日还本金_分"]
        a["over"] += r["当日余款_分"]

    # 月初余额 = 上月末余额；首月月初取放款日初始状态
    open_prin = loan["principal"]
    open_int = 0
    rows = []
    y, m = month_key(start)
    last_key = month_key(end)
    while (y, m) <= last_key:
        period_start = max(start, datetime.date(y, m, 1))
        period_end = min(end, month_end(datetime.date(y, m, 1)))
        a = agg[(y, m)]
        close_prin = by_day[period_end.isoformat()]["日终剩余本金_分"]
        close_int = by_day[period_end.isoformat()]["日终已欠利息_分"]
        close_over = by_day[period_end.isoformat()]["日终余款累计_分"]
        # 勾稽恒等式：
        # 本金: 月初本金 - 当月还本 = 月末本金
        # 利息: 月初欠息 + 当月计息 - 当月还息 = 月末欠息（未付利息不资本化）
        # 现金: 当月收款 = 当月还息 + 当月还本 + 当月余款
        prin_ok = open_prin - a["prin"] == close_prin
        int_ok = open_int + a["accrue"] - a["int"] == close_int
        cash_ok = a["cash"] == a["int"] + a["prin"] + a["over"]
        check = "核对一致" if (prin_ok and int_ok and cash_ok) else "核对不一致"
        rows.append(OrderedDict([
            ("贷款编号", lid),
            ("币种", ccy),
            ("年月", f"{y:04d}-{m:02d}"),
            ("期间起", period_start.isoformat()),
            ("期间止", period_end.isoformat()),
            ("月初剩余本金_分", open_prin),
            ("月初已欠利息_分", open_int),
            ("当月计息_分", a["accrue"]),
            ("当月收款_分", a["cash"]),
            ("当月还利息_分", a["int"]),
            ("当月还本金_分", a["prin"]),
            ("当月余款_分", a["over"]),
            ("月末剩余本金_分", close_prin),
            ("月末已欠利息_分", close_int),
            ("月末余款累计_分", close_over),
            ("本金滚动核对", "一致" if prin_ok else "不一致"),
            ("利息滚动核对", "一致" if int_ok else "不一致"),
            ("现金分配核对", "一致" if cash_ok else "不一致"),
            ("核对结果", check),
        ]))
        open_prin = close_prin
        open_int = close_int
        m += 1
        if m == 13:
            m = 1
            y += 1
    return rows


def build_ccy_monthly(monthly_rows):
    """币种级月度汇总：月度恒等式在汇总层再次成立。"""
    bucket = defaultdict(lambda: defaultdict(int))
    for r in monthly_rows:
        k = (r["币种"], r["年月"])
        b = bucket[k]
        for f in ("月初剩余本金_分", "月初已欠利息_分", "当月计息_分",
                  "当月收款_分", "当月还利息_分", "当月还本金_分",
                  "当月余款_分", "月末剩余本金_分",
                  "月末已欠利息_分", "月末余款累计_分"):
            b[f] += r[f]
    out = []
    for (ccy, ym) in sorted(bucket.keys()):
        b = bucket[(ccy, ym)]
        prin_ok = b["月初剩余本金_分"] - b["当月还本金_分"] == \
            b["月末剩余本金_分"]
        int_ok = b["月初已欠利息_分"] + b["当月计息_分"] - \
            b["当月还利息_分"] == b["月末已欠利息_分"]
        cash_ok = b["当月收款_分"] == b["当月还利息_分"] + \
            b["当月还本金_分"] + b["当月余款_分"]
        out.append(OrderedDict([
            ("币种", ccy),
            ("年月", ym),
            ("月初剩余本金_分", b["月初剩余本金_分"]),
            ("月初已欠利息_分", b["月初已欠利息_分"]),
            ("当月计息_分", b["当月计息_分"]),
            ("当月收款_分", b["当月收款_分"]),
            ("当月还利息_分", b["当月还利息_分"]),
            ("当月还本金_分", b["当月还本金_分"]),
            ("当月余款_分", b["当月余款_分"]),
            ("月末剩余本金_分", b["月末剩余本金_分"]),
            ("月末已欠利息_分", b["月末已欠利息_分"]),
            ("月末余款累计_分", b["月末余款累计_分"]),
            ("本金滚动核对", "一致" if prin_ok else "不一致"),
            ("利息滚动核对", "一致" if int_ok else "不一致"),
            ("现金分配核对", "一致" if cash_ok else "不一致"),
        ]))
    return out


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------

def _dict_to_sheet(rows, headers):
    data = [list(headers)]
    for r in rows:
        data.append([r.get(h, "") for h in headers])
    return data


def reconcile(input_path, output_path, cutoff):
    if not os.path.exists(input_path):
        raise DataError(f"输入文件不存在: {input_path}")
    ext = os.path.splitext(output_path)[1].lower()
    if not ext:
        raise DataError(
            f"输出路径缺少扩展名（pyexcel 按扩展名选择格式）: {output_path}")
    # 多 Sheet 原生 Book 只能落在单一工作簿格式；csv/tsv 会被 pyexcel
    # 拆成目录下的多个文件，无法保证原子替换，明确拒绝
    if ext in (".csv", ".tsv", ".csvz", ".tsvz"):
        raise DataError(
            f"输出格式 {ext} 不支持多 Sheet Book（pyexcel 会拆成多个文件），"
            f"请改用 .xlsx / .xls / .xlsm / .ods")
    if os.path.abspath(input_path) == os.path.abspath(output_path):
        raise DataError("输出路径不能与输入文件相同（来源文件不得改写）")
    if os.path.exists(output_path):
        raise DataError(
            f"输出文件已存在: {output_path}（旧结果不改，请换一个输出路径）")

    book = pyexcel.get_book(file_name=input_path)
    # 结果 Sheet 名不得与来源 Sheet 冲突，避免语义混淆
    conflict = [n for n in book.sheet_names() if n in RESULT_SHEET_NAMES]
    if conflict:
        raise DataError(
            f"来源工作簿含有保留结果 Sheet 名 {conflict}，请先在来源文件中改名")

    loans, rate_events, payments, sheet_names = load_inputs(book)
    if not loans:
        raise DataError("当前贷款 Sheet 没有任何数据行")

    # 截止日缺省：取全部利率/收款事件与放款日的最大日期
    if cutoff is None:
        cutoff = max(
            [ln["start_date"] for ln in loans.values()] +
            [d for m in rate_events.values() for d in m] +
            [p["date"] for p in payments]
        )

    # 截止日后事件：不参与计算，登记到“异常与跳过”
    skipped = []
    active_pay = []
    for p in payments:
        if p["date"] > cutoff:
            skipped.append(OrderedDict([
                ("类型", "收款(晚于截止日)"),
                ("贷款编号", p["loan_id"]),
                ("日期", p["date"].isoformat()),
                ("业务序号", p["seq"]),
                ("金额_分", p["amount"]),
                ("来源行号", p["row_no"]),
                ("说明", "晚于截止日，未参与分配与计息"),
            ]))
        else:
            active_pay.append(p)
    active_rate = {}
    for lid, m in rate_events.items():
        kept = {}
        for d, bps in m.items():
            if d > cutoff:
                skipped.append(OrderedDict([
                    ("类型", "利率变更(晚于截止日)"),
                    ("贷款编号", lid),
                    ("日期", d.isoformat()),
                    ("业务序号", ""),
                    ("金额_分", ""),
                    ("来源行号", ""),
                    ("说明", f"晚于截止日，未生效；新年利率 {bps} 基点"),
                ]))
            else:
                kept[d] = bps
        active_rate[lid] = kept

    pays_by_loan = defaultdict(list)
    for p in active_pay:
        pays_by_loan[p["loan_id"]].append(p)

    all_ledger, all_alloc, balances = [], [], []
    for lid, loan in loans.items():
        ledger_rows, alloc_rows, balance = simulate_loan(
            loan, active_rate.get(lid, {}),
            pays_by_loan.get(lid, []), cutoff)
        all_ledger.extend(ledger_rows)
        all_alloc.extend(alloc_rows)
        balances.append(balance)

    all_monthly = []
    for loan, balance in zip(loans.values(), balances):
        lid = loan["loan_id"]
        all_monthly.extend(build_monthly(
            loan,
            [r for r in all_ledger if r["贷款编号"] == lid],
            [r for r in all_alloc if r["贷款编号"] == lid],
        ))
    ccy_monthly = build_ccy_monthly(all_monthly)

    # 全局自检：逐月勾稽不允许出现任何“不一致”
    bad = [r for r in all_monthly if r["核对结果"] != "核对一致"]
    if bad:
        raise DataError(
            "内部勾稽失败（不应发生）: "
            + "; ".join(f"{r['贷款编号']} {r['年月']} {r['核对结果']}"
                        for r in bad[:5]))
    bad2 = [r for r in ccy_monthly
            if not (r["本金滚动核对"] == r["利息滚动核对"]
                    == r["现金分配核对"] == "一致")]
    if bad2:
        raise DataError("币种级汇总勾稽失败（不应发生）")

    # 现金总量交叉核对：Σ 收款 = Σ 还息 + Σ 还本 + Σ 余款
    total_cash = sum(p["amount"] for p in active_pay)
    total_int = sum(r["还利息_分"] for r in all_alloc)
    total_prin = sum(r["还本金_分"] for r in all_alloc)
    total_over = sum(r["余款_分"] for r in all_alloc)
    if total_cash != total_int + total_prin + total_over:
        raise DataError("现金分配总量不平衡（不应发生）")

    # ---- 组装原生 pyexcel Book：先三张来源 Sheet（原样），再结果 Sheet ----
    out_sheets = OrderedDict()
    for name in sheet_names:
        out_sheets[name] = book.sheet_by_name(name).to_array()

    param_rows = [
        ["项目", "值"],
        ["输入文件", os.path.abspath(input_path)],
        ["生成时间口径", "截止日 " + cutoff.isoformat()],
        ["截止日", cutoff.isoformat()],
        ["金额单位", "整数分"],
        ["利率单位", "整数基点（1 基点 = 0.01% 年利率）"],
        ["计息惯例", "实际天数/365（act/365），逐日按币种 ROUND_HALF_UP"],
        ["日内顺序", "日初利率变更 -> 业务序号收款 -> 日终按剩余本金计息"],
        ["收款分配", "先已欠利息，后本金；余款单列且不作后续贷项"],
        ["利息处理", "未付利息不资本化、不计复利；本金清零后不再计息"],
        ["参与核对收款笔数", len(active_pay)],
        ["跳过事件数", len(skipped)],
        ["收款总额_分", total_cash],
        ["还息总额_分", total_int],
        ["还本总额_分", total_prin],
        ["余款总额_分", total_over],
        ["现金总量核对",
         "一致" if total_cash == total_int + total_prin + total_over
         else "不一致"],
    ]
    out_sheets[OUT_PARAM] = param_rows
    out_sheets[OUT_ALLOC] = _dict_to_sheet(
        all_alloc,
        ["贷款编号", "币种", "收款日", "业务序号", "收款金额_分",
         "还利息_分", "还本金_分", "余款_分", "收款后剩余本金_分",
         "收款后已欠利息_分", "来源行号"])
    out_sheets[OUT_BALANCE] = _dict_to_sheet(
        balances,
        ["贷款编号", "借款人", "币种", "放款日", "截止日", "原始本金_分",
         "截止日剩余本金_分", "截止日已欠利息_分", "截止日余款累计_分",
         "截止日应收合计_分", "当前年利率_基点"])
    out_sheets[OUT_LEDGER] = _dict_to_sheet(
        all_ledger,
        ["贷款编号", "币种", "日期", "当日利率_基点", "当日收款_分",
         "当日还利息_分", "当日还本金_分", "当日余款_分", "当日计息_分",
         "日终剩余本金_分", "日终已欠利息_分", "日终余款累计_分"])
    out_sheets[OUT_MONTHLY] = _dict_to_sheet(
        all_monthly,
        ["贷款编号", "币种", "年月", "期间起", "期间止",
         "月初剩余本金_分", "月初已欠利息_分", "当月计息_分",
         "当月收款_分", "当月还利息_分", "当月还本金_分", "当月余款_分",
         "月末剩余本金_分", "月末已欠利息_分", "月末余款累计_分",
         "本金滚动核对", "利息滚动核对", "现金分配核对", "核对结果"])
    out_sheets[OUT_CCY_MONTHLY] = _dict_to_sheet(
        ccy_monthly,
        ["币种", "年月", "月初剩余本金_分", "月初已欠利息_分",
         "当月计息_分", "当月收款_分", "当月还利息_分", "当月还本金_分",
         "当月余款_分", "月末剩余本金_分", "月末已欠利息_分",
         "月末余款累计_分", "本金滚动核对", "利息滚动核对",
         "现金分配核对"])
    if skipped:
        out_sheets[OUT_SKIPPED] = _dict_to_sheet(
            skipped,
            ["类型", "贷款编号", "日期", "业务序号", "金额_分",
             "来源行号", "说明"])
    else:
        out_sheets[OUT_SKIPPED] = [["类型", "贷款编号", "日期", "业务序号",
                                    "金额_分", "来源行号", "说明"],
                                   ["（无跳过事件）", "", "", "", "", "", ""]]

    out_book = pyexcel.Book(out_sheets)
    # 先写临时文件再原子替换，避免失败时留下半截“结果”
    tmp_path = output_path + ".part" + ext
    if os.path.exists(tmp_path):
        os.remove(tmp_path)
    out_book.save_as(tmp_path)
    os.replace(tmp_path, output_path)
    return {
        "output": output_path,
        "cutoff": cutoff.isoformat(),
        "loans": len(loans),
        "payments": len(active_pay),
        "skipped": len(skipped),
        "totals": (total_cash, total_int, total_prin, total_over),
    }


def main(argv):
    if len(argv) not in (3, 4):
        print(__doc__)
        print("用法: python3 -I loan_reconcile.py "
              "<输入工作簿> <输出工作簿> [截止日 YYYY-MM-DD]", file=sys.stderr)
        return 2
    input_path, output_path = argv[1], argv[2]
    cutoff = None
    if len(argv) == 4:
        try:
            cutoff = datetime.date.fromisoformat(argv[3])
        except ValueError:
            print(f"截止日格式错误: {argv[3]}，要求 YYYY-MM-DD",
                  file=sys.stderr)
            return 2
    try:
        info = reconcile(input_path, output_path, cutoff)
    except DataError as exc:
        print(f"报错中止: {exc}", file=sys.stderr)
        return 1
    print("完成:", info)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
