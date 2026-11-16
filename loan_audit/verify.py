# -*- coding: utf-8 -*-
"""独立核验脚本：不调用引擎计算代码，用 Decimal 重算并逐项比对输出 Book。

核验内容：
  1. 来源各 Sheet 在输出 Book 中逐单元格保持不变；
  2. 逐日台账与独立 Decimal 重算（act/365、ROUND_HALF_UP、日内顺序）完全一致；
  3. 逐笔分配的“先息后本、余款单列”及收款后余额与台账一致；
  4. 逐月勾稽的三条恒等式仅用输出表自身数据再次成立并滚动衔接；
  5. 币种汇总 = 明细月度加总；现金总量恒等式；截止日余额 = 台账末日值。
用法：python3 -I verify.py <输入工作簿> <结果工作簿> <截止日>
"""
import datetime
import sys
from collections import defaultdict
from decimal import ROUND_HALF_UP, Decimal

import pyexcel

sys.path.insert(0, "/workspace/loan_audit")
from loan_reconcile import (  # noqa: E402  仅复用表头别名常量，计算逻辑独立
    LOAN_COLUMNS, RATE_COLUMNS, PAY_COLUMNS, _norm)

DEN = Decimal(10000 * 365)
# 各 Sheet 自己的 别名 -> 规范字段（逐表映射，避免跨表同名冲突）
_ALIAS_MAPS = []
for table in (LOAN_COLUMNS, RATE_COLUMNS, PAY_COLUMNS):
    m = {}
    for field, aliases in table.items():
        for a in aliases:
            m[_norm(a)] = field
    _ALIAS_MAPS.append(m)
LOAN_ALIAS, RATE_ALIAS, PAY_ALIAS = _ALIAS_MAPS
# 结果 Sheet 的中文表头直接使用，不做别名转换
_RESULT_FIELDS = None


def records(sheet, alias_map=None):
    arr = sheet.to_array()
    if alias_map is None:
        header = [str(c).strip() for c in arr[0]]
    else:
        header = [alias_map.get(_norm(c), str(c).strip()) for c in arr[0]]
    out = []
    for row in arr[1:]:
        if all(c is None or (isinstance(c, str) and not c.strip()) for c in row):
            continue
        out.append(dict(zip(header, row)))
    return out


def sheet_with(src, field):
    wanted = set()
    for table in (LOAN_COLUMNS, RATE_COLUMNS, PAY_COLUMNS):
        if field in table:
            wanted |= {_norm(a) for a in table[field]}
    for n in src.sheet_names():
        header = {_norm(c) for c in src.sheet_by_name(n).to_array()[0]}
        if wanted & header:
            return src.sheet_by_name(n)
    raise KeyError(field)


def as_date(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    return v if isinstance(v, datetime.date) else datetime.date.fromisoformat(str(v))


def main(inp, outp, cutoff_s):
    cutoff = datetime.date.fromisoformat(cutoff_s)
    src = pyexcel.get_book(file_name=inp)
    res = pyexcel.get_book(file_name=outp)
    errors = []

    # 1) 来源 Sheet 原样保留
    for name in src.sheet_names():
        if src.sheet_by_name(name).to_array() != \
                res.sheet_by_name(name).to_array():
            errors.append(f"来源 Sheet「{name}」在结果中被改动")

    loans = {r["loan_id"]: r
             for r in records(sheet_with(src, "principal"), LOAN_ALIAS)}
    rate_rows = records(sheet_with(src, "eff_date"), RATE_ALIAS)
    pay_rows = records(sheet_with(src, "pay_date"), PAY_ALIAS)

    # ---- 独立重算（Decimal，逻辑刻意与引擎分开写） ----
    expected_ledger = {}
    expected_alloc = []
    for lid, ln in loans.items():
        start = as_date(ln["start_date"])
        P = int(ln["principal"])
        rate = int(ln["rate_bps"])
        rmap = {as_date(r["eff_date"]): int(r["rate_bps"])
                for r in rate_rows if r["loan_id"] == lid
                and as_date(r["eff_date"]) <= cutoff}
        pays = sorted(
            [r for r in pay_rows if r["loan_id"] == lid
             and as_date(r["pay_date"]) <= cutoff],
            key=lambda r: (as_date(r["pay_date"]), int(r["seq"])))
        pmap = defaultdict(list)
        for r in pays:
            pmap[as_date(r["pay_date"])].append(r)
        I, over = 0, 0
        day = start
        while day <= cutoff:
            if day in rmap:
                rate = rmap[day]
            dcash = dint = dprin = dover = 0
            for p in pmap.get(day, ()):
                cash = int(p["amount"])
                ti = min(cash, I)
                I -= ti
                rest = cash - ti
                tp = min(rest, P)
                P -= tp
                rest -= tp
                to = rest
                over += to
                dcash += cash
                dint += ti
                dprin += tp
                dover += to
                expected_alloc.append({
                    "loan_id": lid, "date": day, "seq": int(p["seq"]),
                    "amount": cash, "to_int": ti, "to_prin": tp,
                    "to_over": to, "P_after": P, "I_after": I})
            acc = 0
            if P > 0:
                acc = int((Decimal(P) * Decimal(rate) / DEN)
                          .quantize(Decimal("1"), rounding=ROUND_HALF_UP))
                I += acc
            expected_ledger[(lid, day)] = (
                rate, dcash, dint, dprin, dover, acc, P, I, over)
            day += datetime.timedelta(days=1)

    # 2) 比对逐日台账
    got_ledger = {}
    for r in records(res.sheet_by_name("逐日台账")):
        got_ledger[(r["贷款编号"], as_date(r["日期"]))] = (
            int(r["当日利率_基点"]), int(r["当日收款_分"]),
            int(r["当日还利息_分"]), int(r["当日还本金_分"]),
            int(r["当日余款_分"]), int(r["当日计息_分"]),
            int(r["日终剩余本金_分"]), int(r["日终已欠利息_分"]),
            int(r["日终余款累计_分"]))
    if got_ledger != expected_ledger:
        diff = [k for k in expected_ledger.keys() & got_ledger.keys()
                if expected_ledger[k] != got_ledger[k]]
        errors.append(
            f"逐日台账不一致 缺={list(set(expected_ledger)-set(got_ledger))[:3]} "
            f"多={list(set(got_ledger)-set(expected_ledger))[:3]} 差异={diff[:5]}")

    # 3) 比对逐笔分配
    got_alloc = records(res.sheet_by_name("逐笔分配"))
    if len(got_alloc) != len(expected_alloc):
        errors.append(f"逐笔分配笔数 {len(got_alloc)} != 期望 {len(expected_alloc)}")
    gkey = sorted(got_alloc,
                  key=lambda r: (r["贷款编号"], as_date(r["收款日"]),
                                 int(r["业务序号"])))
    ekey = sorted(expected_alloc, key=lambda r: (r["loan_id"], r["date"], r["seq"]))
    field_map = [("收款金额_分", "amount"), ("还利息_分", "to_int"),
                 ("还本金_分", "to_prin"), ("余款_分", "to_over"),
                 ("收款后剩余本金_分", "P_after"),
                 ("收款后已欠利息_分", "I_after")]
    for g, e in zip(gkey, ekey):
        for gf, ef in field_map:
            if int(g[gf]) != e[ef]:
                errors.append(
                    f"逐笔分配 {e['loan_id']} {e['date']} 序号{e['seq']} "
                    f"列{gf}: {g[gf]} != {e[ef]}")
                break

    # 4) 月度恒等式（仅用输出表）+ 月末衔接次月月初
    mon = records(res.sheet_by_name("逐月勾稽"))
    for r in mon:
        if int(r["月初剩余本金_分"]) - int(r["当月还本金_分"]) != \
                int(r["月末剩余本金_分"]):
            errors.append(f"本金恒等式失败: {r['贷款编号']} {r['年月']}")
        if int(r["月初已欠利息_分"]) + int(r["当月计息_分"]) - \
                int(r["当月还利息_分"]) != int(r["月末已欠利息_分"]):
            errors.append(f"利息恒等式失败: {r['贷款编号']} {r['年月']}")
        if int(r["当月收款_分"]) != int(r["当月还利息_分"]) + \
                int(r["当月还本金_分"]) + int(r["当月余款_分"]):
            errors.append(f"现金恒等式失败: {r['贷款编号']} {r['年月']}")
        if r["核对结果"] != "核对一致":
            errors.append(f"核对结果列异常: {r['贷款编号']} {r['年月']}")
    by_loan = defaultdict(list)
    for r in mon:
        by_loan[r["贷款编号"]].append(r)
    for lid, rows in by_loan.items():
        rows.sort(key=lambda r: r["年月"])
        for a, b in zip(rows, rows[1:]):
            if int(a["月末剩余本金_分"]) != int(b["月初剩余本金_分"]) or \
               int(a["月末已欠利息_分"]) != int(b["月初已欠利息_分"]):
                errors.append(f"月度滚动衔接失败: {lid} {a['年月']}->{b['年月']}")
        first = rows[0]
        if int(first["月初剩余本金_分"]) != int(loans[lid]["principal"]) or \
                int(first["月初已欠利息_分"]) != 0:
            errors.append(f"首月月初余额错误: {lid}")

    # 5) 币种汇总 = 月度明细加总
    ccy = records(res.sheet_by_name("币种月度汇总"))
    bucket = defaultdict(lambda: defaultdict(int))
    num_fields = ("月初剩余本金_分", "月初已欠利息_分", "当月计息_分",
                  "当月收款_分", "当月还利息_分", "当月还本金_分",
                  "当月余款_分", "月末剩余本金_分", "月末已欠利息_分",
                  "月末余款累计_分")
    for r in mon:
        for f in num_fields:
            bucket[(r["币种"], r["年月"])][f] += int(r[f])
    if len(ccy) != len(bucket):
        errors.append("币种汇总行数不符")
    for r in ccy:
        for f in num_fields:
            if int(r[f]) != bucket[(r["币种"], r["年月"])][f]:
                errors.append(f"币种汇总不符 {r['币种']} {r['年月']} {f}")
                break

    # 6) 截止日余额 = 台账末日；应收合计 = 本金 + 欠息
    for r in records(res.sheet_by_name("截止日余额")):
        last = expected_ledger[(r["贷款编号"], cutoff)]
        if int(r["截止日剩余本金_分"]) != last[6] or \
                int(r["截止日已欠利息_分"]) != last[7] or \
                int(r["截止日余款累计_分"]) != last[8]:
            errors.append(f"截止日余额与台账不符: {r['贷款编号']}")
        if int(r["截止日应收合计_分"]) != last[6] + last[7]:
            errors.append(f"应收合计错误: {r['贷款编号']}")

    # 7) 跳过事件登记完整
    skip = records(res.sheet_by_name("异常与跳过"))
    n_skip = len([r for r in pay_rows if as_date(r["pay_date"]) > cutoff]) + \
        len([r for r in rate_rows if as_date(r["eff_date"]) > cutoff])
    if len([r for r in skip if "晚于截止日" in str(r.get("说明", ""))]) != n_skip:
        errors.append("跳过事件登记数量不符")

    if errors:
        print("核验失败:")
        for e in errors[:20]:
            print(" -", e)
        return 1
    print(f"全部核验通过：台账 {len(got_ledger)} 行 / 分配 {len(got_alloc)} 笔 / "
          f"月度 {len(mon)} 行 / 币种汇总 {len(ccy)} 行 / "
          f"余额 {len(loans)} 笔 / 跳过 {n_skip} 事件")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1], sys.argv[2], sys.argv[3]))
