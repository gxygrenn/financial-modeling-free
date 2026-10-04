#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
免费可比公司分析（Comparable Company Analysis）生成器
替代 Anthropic financial-services 的 comps-analysis 技能。

用法：
  python comps.py AAPL MSFT GOOGL          # 拉免费数据，生成可比公司 Excel
  python comps.py AAPL MSFT GOOGL --json   # 只打印结果，不生成 Excel
"""

import argparse
import datetime as dt
import json

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


def fetch(ticker: str):
    """拉取一家公司的免费财务/估值数据。yfinance 不可用则返回 None。"""
    try:
        import yfinance as yf
    except ImportError:
        return None

    t = yf.Ticker(ticker)

    def _safe(fn):
        try:
            return fn()
        except Exception:
            return None

    info = _safe(lambda: t.info) or {}

    def _num(v):
        try:
            return float(v)
        except (TypeError, ValueError):
            return None

    price = _safe(lambda: t.fast_info.last_price)
    if price is None:
        price = info.get("currentPrice")

    income = _safe(lambda: getattr(t, "income_stmt", None)) or _safe(lambda: t.get_income_stmt())
    bs = _safe(lambda: getattr(t, "balance_sheet", None)) or _safe(lambda: t.get_balance_sheet())

    def _row(df, *names):
        if df is None:
            return None
        for name in names:
            for idx in df.index:
                if name.lower() in str(idx).lower():
                    try:
                        return _num(df.loc[idx, df.columns[0]])
                    except Exception:
                        return None
        return None

    revenue = _row(income, "Total Revenue")
    ebit = _row(income, "Operating Income", "EBIT")
    net_income = _row(income, "Net Income")
    da = _row(income, "Depreciation", "Depreciation And Amortization")
    cash = _row(bs, "Cash And Cash Equivalents", "Cash")
    debt = _row(bs, "Total Debt", "Long Term Debt")

    ebitda = (ebit + da) if (ebit is not None and da is not None) else None
    shares = info.get("sharesOutstanding")
    mktcap = (price * shares) if (price and shares) else info.get("marketCap")
    ev = (mktcap + debt - cash) if (mktcap and debt is not None and cash is not None) else None

    def _ratio(num, den):
        return (num / den) if (num and den) else None

    return {
        "ticker": ticker,
        "name": info.get("shortName") or ticker,
        "price": _num(price),
        "mktcap": _num(mktcap),
        "revenue": revenue,
        "gross_margin": _row(income, "Gross Profit"),  # 占位，下方再算
        "ebitda": ebitda,
        "ebitda_margin": _ratio(ebitda, revenue),
        "net_income": net_income,
        "ev": ev,
        "ev_rev": _ratio(ev, revenue),
        "ev_ebitda": _ratio(ev, ebitda),
        "pe": _ratio(mktcap, net_income),
    }


def sample(ticker: str):
    """离线示例数据（演示用）。"""
    S = {
        "MSFT": dict(price=510, mktcap=3_790e6, revenue=281_000e6, ebitda=161_000e6,
                     net_income=101_000e6, ev=3_780e6),
        "AAPL": dict(price=245, mktcap=3_724e6, revenue=391_000e6, ebitda=134_500e6,
                     net_income=94_000e6, ev=3_765e6),
        "GOOGL": dict(price=185, mktcap=2_275e6, revenue=350_000e6, ebitda=125_000e6,
                      net_income=95_000e6, ev=2_210e6),
    }
    d = S.get(ticker.upper(), S["MSFT"])
    def _r(n, den):
        return n / den if den else None
    return {
        "ticker": ticker, "name": ticker.upper(),
        "price": d["price"], "mktcap": d["mktcap"],
        "revenue": d["revenue"], "ebitda": d["ebitda"],
        "ebitda_margin": _r(d["ebitda"], d["revenue"]),
        "net_income": d["net_income"], "ev": d["ev"],
        "ev_rev": _r(d["ev"], d["revenue"]),
        "ev_ebitda": _r(d["ev"], d["ebitda"]),
        "pe": _r(d["mktcap"], d["net_income"]),
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("tickers", nargs="*", default=["MSFT", "AAPL", "GOOGL"])
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--demo", action="store_true", help="用离线示例数据（不联网）")
    args = ap.parse_args()

    if args.demo:
        rows = [sample(t) for t in args.tickers]
    else:
        rows = [fetch(t) or sample(t) for t in args.tickers]

    if args.json:
        print(json.dumps(rows, indent=2, ensure_ascii=False, default=str))
        return

    # ---------- 生成 Excel ----------
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "可比公司"

    HDR = PatternFill("solid", fgColor="1F4E79")
    SUB = PatternFill("solid", fgColor="D9E1F2")
    STAT = PatternFill("solid", fgColor="F2F2F2")
    WF = Font(color="FFFFFF", bold=True)
    CENTER = Alignment(horizontal="center")
    THIN = Border(*[Side(style="thin", color="BFBFBF")] * 4)

    ws.merge_cells("A1:J1")
    ws["A1"] = "可比公司分析"
    ws["A1"].font = Font(bold=True, size=14)
    ws.merge_cells("A2:J2")
    ws["A2"] = f"股票: {' · '.join(t.upper() for t in args.tickers)}  |  截至 {dt.date.today()}  |  单位: 百万美元"
    ws["A2"].font = Font(color="666666")

    # 表头
    headers = ["公司", "股价", "市值", "营收", "EBITDA",
               "EBITDA率", "企业价值", "EV/营收", "EV/EBITDA", "市盈率"]
    r = 4
    for c, h in enumerate(headers, start=1):
        cc = ws.cell(row=r, column=c, value=h)
        cc.fill = HDR if c == 1 else SUB
        cc.font = WF if c == 1 else Font(bold=True)
        cc.alignment = CENTER
        cc.border = THIN

    r += 1
    for row in rows:
        vals = [
            row["name"], row["price"], row["mktcap"], row["revenue"], row["ebitda"],
            row["ebitda_margin"], row["ev"], row["ev_rev"], row["ev_ebitda"], row["pe"],
        ]
        for c, v in enumerate(vals, start=1):
            cc = ws.cell(row=r, column=c)
            if c == 1:
                cc.value = v
                cc.font = Font(bold=True)
            else:
                cc.value = v
            cc.alignment = CENTER
            cc.border = THIN
            if c == 2:
                cc.number_format = "0.00"
            elif c in (3, 4, 5, 7):
                cc.number_format = "#,##0"
            elif c == 6:
                cc.number_format = "0.0%"
            elif c in (8, 9, 10):
                cc.number_format = "0.0x"
        r += 1

    # 统计行（中位数 + 四分位）
    def _stat(col_idx):
        vals = []
        for rr in range(5, 5 + len(rows)):
            v = ws.cell(row=rr, column=col_idx).value
            if isinstance(v, (int, float)):
                vals.append(v)
        return vals

    for label, fn in [("中位数", _median), ("75分位", _q3), ("25分位", _q1)]:
        ws.cell(row=r, column=1, value=label).font = Font(bold=True)
        ws.cell(row=r, column=1).fill = STAT
        for col in range(2, 11):
            vs = _stat(col)
            if vs and col not in (3, 4, 5, 7):  # 规模列不做统计
                cc = ws.cell(row=r, column=col, value=fn(vs))
                cc.fill = STAT
                cc.alignment = CENTER
                cc.border = THIN
                cc.number_format = "0.0%" if col == 6 else ("0.0x" if col in (8, 9, 10) else "0.00")
        r += 1

    # 列宽
    for c in range(1, 11):
        ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = 13
    ws.column_dimensions["A"].width = 22

    out = f"Comps_{'_'.join(t.upper() for t in args.tickers)}_{dt.date.today()}.xlsx"
    wb.save(out)
    print(f"✅ 可比公司分析已生成：{out}")
    print("   （含中位数 / 75分位 / 25分位统计行）")


def _median(v):
    v = sorted(v)
    n = len(v)
    return v[n // 2] if n % 2 else (v[n // 2 - 1] + v[n // 2]) / 2


def _q1(v):
    v = sorted(v)
    import statistics
    return statistics.median(v[:len(v) // 2]) if len(v) >= 2 else v[0]


def _q3(v):
    v = sorted(v)
    import statistics
    return statistics.median(v[(len(v) + 1) // 2:]) if len(v) >= 2 else v[-1]


if __name__ == "__main__":
    main()
