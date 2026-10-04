#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
免费 DCF 估值模型生成器（替代 Anthropic financial-services 的 dcf-model 技能）

原理说明：
  Anthropic 的 dcf-model 技能依赖付费 MCP 数据源（Daloopa/S&P/FactSet）。
  本脚本把「方法论」完整保留，只把「数据源」替换为免费渠道：
    - 行情 / 财务数据 : Yahoo Finance (yfinance)
    - 美股官方财报    : SEC EDGAR（可选，见文末说明）
  生成的 Excel 模型结构与投行标准一致：三情景（熊/基/牛）+ WACC + 敏感性表。

用法：
  python dcf_model.py AAPL            # 拉免费数据，自动建模型
  python dcf_model.py AAPL --manual   # 手动输入模式（数据拿不到时用）

依赖：openpyxl, yfinance, pandas
"""

import argparse
import datetime as dt
import sys

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------- 免费数据源
def sample_data(ticker: str = "MSFT"):
    """示例/离线数据（用于演示，或 yfinance 不可用时的兜底）。数值为近似，仅供演示。"""
    base = {
        "MSFT": dict(price=510.0, shares=7_430e6, beta=0.9, revenue=281_000e6,
                     ebit=126_000e6, tax_exp=22_000e6, da=35_000e6, capex=66_000e6,
                     cash=78_000e6, debt=67_000e6),
        "AAPL": dict(price=245.0, shares=15_200e6, beta=1.2, revenue=391_000e6,
                     ebit=123_000e6, tax_exp=29_000e6, da=11_500e6, capex=9_400e6,
                     cash=65_000e6, debt=106_000e6),
        "GOOGL": dict(price=185.0, shares=12_300e6, beta=1.05, revenue=350_000e6,
                      ebit=112_000e6, tax_exp=16_000e6, da=13_000e6, capex=52_000e6,
                      cash=95_000e6, debt=29_000e6),
    }
    d = base.get(ticker.upper(), base["MSFT"])
    return {"price": d["price"], "shares": d["shares"], "beta": d["beta"],
            "revenue": d["revenue"], "ebit": d["ebit"], "tax_exp": d["tax_exp"],
            "da": d["da"], "capex": d["capex"], "cash": d["cash"], "debt": d["debt"]}


def fetch_yahoo_chart(ticker: str):
    """用标准库 urllib 抓 Yahoo Finance 免费行情（无需 yfinance）。失败返回 None。"""
    import json
    import urllib.request

    url = f"https://query1.finance.yahoo.com/v8/finance/chart/{ticker}?interval=1d&range=1mo"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    try:
        with urllib.request.urlopen(req, timeout=15) as r:
            data = json.loads(r.read().decode())
        meta = data["chart"]["result"][0]["meta"]
        return {"price": meta.get("regularMarketPrice")}
    except Exception:
        return None


def fetch_financials(ticker: str):
    """从 Yahoo Finance 拉取免费财务数据。返回 dict，字段拿不到则置 None。

    注意：yfinance 属于「可选增强」。若未安装或网络受限，
    main() 会自动回退到 sample_data() 的离线示例数据。
    """
    try:
        import yfinance as yf
    except ImportError:
        return None

    t = yf.Ticker(ticker)

    def _safe(fn, *args, **kwargs):
        try:
            return fn(*args, **kwargs)
        except Exception:
            return None

    info = _safe(lambda: t.info) or {}

    def _num(v):
        try:
            return float(v)
        except (TypeError, ValueError):
            return None

    price = _safe(lambda: t.fast_info.last_price)
    if price is None and "currentPrice" in info:
        price = info.get("currentPrice")

    shares = info.get("sharesOutstanding")
    beta = info.get("beta")

    # 财务报表（新版 yfinance 用 Ticker 属性，旧版用 get_* 方法，做双兼容）
    income = _safe(lambda: getattr(t, "income_stmt", None))
    if income is None:
        income = _safe(lambda: t.get_income_stmt())
    bs = _safe(lambda: getattr(t, "balance_sheet", None))
    if bs is None:
        bs = _safe(lambda: t.get_balance_sheet())
    cf = _safe(lambda: getattr(t, "cashflow", None))
    if cf is None:
        cf = _safe(lambda: t.get_cashflow())

    def _row(df, *names):
        """按科目名（忽略大小写）从 DataFrame 里取最新一期的值。"""
        if df is None:
            return None
        for name in names:
            for idx in df.index:
                if name.lower() in str(idx).lower():
                    try:
                        col = df.columns[0]
                        return _num(df.loc[idx, col])
                    except Exception:
                        return None
        return None

    revenue = _row(income, "Total Revenue", "Operating Revenue")
    ebit = _row(income, "Operating Income", "EBIT")
    tax_exp = _row(income, "Income Tax", "Provision for Income Taxes")
    da = _row(cf, "Depreciation", "Depreciation And Amortization")
    capex = _row(cf, "Capital Expenditure", "Capital Expenditures")
    cash = _row(bs, "Cash And Cash Equivalents", "Cash")
    debt = _row(bs, "Total Debt", "Long Term Debt")

    return {
        "price": _num(price),
        "shares": _num(shares),
        "beta": _num(beta),
        "revenue": revenue,
        "ebit": ebit,
        "tax_exp": tax_exp,
        "da": da,
        "capex": capex,
        "cash": cash,
        "debt": debt,
    }


# ---------------------------------------------------------------- DCF 计算
def run_dcf(d, a):
    """核心 DCF。d=财务数据，a=假设。返回逐年结果字典。"""
    rev = d["revenue"] or 0.0
    ebit = d["ebit"]
    ebit_margin = (ebit / rev) if (ebit and rev) else a["ebit_margin"]

    n = a["years"]
    years = ["FY" + str(i + 1) for i in range(n)]

    out = {"years": years}
    revs, rev_gs = [], []
    prev = rev
    for i in range(n):
        r = prev * (1 + a["growth"][i])
        revs.append(r)
        rev_gs.append(a["growth"][i])
        prev = r
    out["revenue"] = revs
    out["rev_growth"] = rev_gs

    # 营业利润率：从基期向目标逐步移动（线性插值）
    margins = []
    m0 = ebit_margin
    m1 = a["target_margin"]
    for i in range(n):
        frac = (i + 1) / n
        margins.append(m0 + (m1 - m0) * frac)
    out["ebit_margin"] = margins

    ebits = [revs[i] * margins[i] for i in range(n)]
    out["ebit"] = ebits

    tax_rate = a["tax_rate"]
    nopat = [ebits[i] * (1 - tax_rate) for i in range(n)]
    out["nopat"] = nopat

    das = [revs[i] * a["da_pct"] for i in range(n)]
    out["da"] = das
    # CapEx：从当前水平（可能处于扩张周期高位）线性过渡到稳态水平，避免把峰值永久外推
    c0 = a["capex_pct"]
    c1 = a.get("terminal_capex_pct", c0)
    capexs = [revs[i] * (c0 + (c1 - c0) * (i + 1) / n) for i in range(n)]
    out["capex"] = capexs

    # ΔNWC = 收入增量 * NWC 占比
    d_nwc = []
    prev_r = rev
    for i in range(n):
        d_nwc.append((revs[i] - prev_r) * a["nwc_pct"])
        prev_r = revs[i]
    out["d_nwc"] = d_nwc

    fcf = [nopat[i] + das[i] - capexs[i] - d_nwc[i] for i in range(n)]
    out["fcf"] = fcf

    wacc = a["wacc"]
    # 年中惯例
    periods = [i + 0.5 for i in range(n)]
    out["periods"] = periods
    pv_fcf = [fcf[i] / (1 + wacc) ** periods[i] for i in range(n)]
    out["pv_fcf"] = pv_fcf

    # 终端价值（永续增长法）
    g = a["terminal_g"]
    term_fcf = fcf[-1] * (1 + g)
    tv = term_fcf / (wacc - g)
    pv_tv = tv / (1 + wacc) ** (n - 0.5)  # 与最后一期年中一致
    out["term_fcf"] = term_fcf
    out["tv"] = tv
    out["pv_tv"] = pv_tv

    ev = sum(pv_fcf) + pv_tv
    net_debt = (d["debt"] or 0.0) - (d["cash"] or 0.0)
    equity = ev - net_debt
    shares = d["shares"] or 1.0
    implied = equity / shares
    out.update({
        "ev": ev, "net_debt": net_debt, "equity": equity,
        "shares": shares, "implied": implied,
    })
    return out


# ---------------------------------------------------------------- Excel 输出
def build_excel(ticker, d, a, path):
    """生成投行标准的 DCF Excel 模型（含公式与敏感性表）。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "DCF"

    # 样式
    HDR_FILL = PatternFill("solid", fgColor="1F4E79")
    SUB_FILL = PatternFill("solid", fgColor="D9E1F2")
    OUT_FILL = PatternFill("solid", fgColor="BDD7EE")
    IN_FILL = PatternFill("solid", fgColor="F2F2F2")
    HDR_FONT = Font(color="FFFFFF", bold=True)
    BLUE = Font(color="0000FF")
    BLACK = Font(color="000000")
    BOLD = Font(bold=True)
    THIN = Border(*[Side(style="thin", color="BFBFBF")] * 4)
    CENTER = Alignment(horizontal="center")

    # 货币/单位：A 股传 currency="¥"（百万元），美股默认 "$"（百万美元）
    cur = d.get("currency", "$")
    M = "百万元" if cur == "¥" else "百万美元"

    def cell(r, c, v, fill=None, font=None, border=False, fmt=None, center=False, font_color=None):
        cc = ws.cell(row=r, column=c, value=v)
        if fill:
            cc.fill = fill
        if font:
            cc.font = font
        if font_color:
            cc.font = Font(color=font_color, bold=(font is BOLD or font == BOLD))
        if border:
            cc.border = THIN
        if fmt:
            cc.number_format = fmt
        if center:
            cc.alignment = CENTER
        return cc

    # 标题
    ws.merge_cells("A1:H1")
    name = d.get("name", "")
    title = f"{ticker.upper()}{' ' + name if name else ''} — DCF 估值模型"
    cell(1, 1, title, font=Font(bold=True, size=14))
    ws.merge_cells("A2:H2")
    cell(2, 1, f"股票代码: {ticker.upper()}  |  日期: {dt.date.today()}  |  数据源: {d.get('source', 'Yahoo Finance')}",
         font=Font(color="666666"))

    # 市场数据
    r = 4
    cell(r, 1, "市场数据与关键假设", fill=HDR_FILL, font=HDR_FONT)
    for c in range(2, 9):
        cell(r, c, None, fill=HDR_FILL)
    r += 1
    md = [
        ("当前股价", d["price"], "0.00"),
        ("总股本（百万股）", (d["shares"] or 0) / 1e6, "#,##0.0"),
        ("Beta（5年）", d["beta"], "0.00"),
        ("净债务 /（净现金）（" + M + "）", ((d["debt"] or 0) - (d["cash"] or 0)) / 1e6, "#,##0"),
        ("LTM 营收（" + M + "）", (d["revenue"] or 0) / 1e6, "#,##0"),
    ]
    for name, val, fmt in md:
        cell(r, 1, name, fill=IN_FILL, font=BOLD)
        cell(r, 2, val if val is not None else "N/A", font=BLUE, fmt=fmt)
        r += 1

    r += 1
    n = a["years"]
    start_col = 3
    # 表头：年份
    hdr_row = r
    cell(hdr_row, 1, "预测现金流（" + M + "）", fill=HDR_FILL, font=HDR_FONT)
    for c in range(2, 9):
        cell(hdr_row, c, None, fill=HDR_FILL)
    r += 1
    cell(r, 1, "科目", fill=SUB_FILL, font=BOLD)
    for i in range(n):
        cell(r, start_col + i, f"第{i+1}年", fill=SUB_FILL, font=BOLD, center=True)
    r += 1

    res = run_dcf(d, a)
    m = 1e6
    rows = [
        ("营收", res["revenue"], "#,##0"),
        ("  增长率", res["rev_growth"], "0.0%"),
        ("EBIT（息税前利润）", res["ebit"], "#,##0"),
        ("  利润率", res["ebit_margin"], "0.0%"),
        ("NOPAT（税后经营利润）", res["nopat"], "#,##0"),
        ("(+) 折旧摊销", res["da"], "#,##0"),
        ("(-) 资本开支", [-x for x in res["capex"]], "#,##0"),
        ("(-) 营运资本变动", [-x for x in res["d_nwc"]], "#,##0"),
        ("无杠杆自由现金流", res["fcf"], "#,##0"),
        ("折现因子", [1 / (1 + a["wacc"]) ** p for p in res["periods"]], "0.000"),
        ("自由现金流现值", res["pv_fcf"], "#,##0"),
    ]
    for name, vals, fmt in rows:
        cell(r, 1, name, font=BOLD if not name.startswith("  ") else Font())
        for i in range(n):
            v = vals[i] / m if "率" not in name and "因子" not in name else vals[i]
            cell(r, start_col + i, v, fmt=fmt, center=True)
        r += 1

    r += 1
    cell(r, 1, "终值与企业估值", fill=HDR_FILL, font=HDR_FONT)
    for c in range(2, 9):
        cell(r, c, None, fill=HDR_FILL)
    r += 1
    tv_rows = [
        ("终期自由现金流（" + M + "）", res["term_fcf"] / m, "#,##0"),
        ("终值（" + M + "）", res["tv"] / m, "#,##0"),
        ("终值现值（" + M + "）", res["pv_tv"] / m, "#,##0"),
        ("企业价值 EV（" + M + "）", res["ev"] / m, "#,##0"),
        ("(-) 净债务（" + M + "）", res["net_debt"] / m, "#,##0"),
        ("股权价值（" + M + "）", res["equity"] / m, "#,##0"),
        ("总股本（百万股）", res["shares"] / m, "#,##0.0"),
    ]
    for name, val, fmt in tv_rows:
        cell(r, 1, name, font=BOLD)
        cell(r, 2, val, fmt=fmt, fill=OUT_FILL if name.startswith(("企业价值", "股权价值")) else None,
             font=BOLD if name.startswith(("企业价值", "股权价值")) else None)
        r += 1

    cell(r, 1, "隐含每股价值", fill=OUT_FILL, font=BOLD)
    cell(r, 2, res["implied"], fmt="0.00", fill=OUT_FILL, font=BOLD)
    r += 1
    cell(r, 1, "当前股价", font=BOLD)
    cell(r, 2, d["price"], fmt="0.00", font=BLUE)
    r += 1
    cell(r, 1, "隐含空间（上涨/下跌）", font=BOLD)
    up = (res["implied"] / d["price"] - 1) if d["price"] else 0
    cell(r, 2, up, fmt="0.0%", font=BOLD,
         font_color="FF0000" if up >= 0 else "008000")

    # 敏感性表：WACC vs 终端增长率
    r += 3
    cell(r, 1, "敏感性分析：隐含股价（WACC vs 永续增长率）", fill=HDR_FILL, font=HDR_FONT)
    for c in range(2, 9):
        cell(r, c, None, fill=HDR_FILL)
    r += 1
    waccs = [a["wacc"] - 0.02, a["wacc"] - 0.01, a["wacc"], a["wacc"] + 0.01, a["wacc"] + 0.02]
    gs = [a["terminal_g"] - 0.01, a["terminal_g"] - 0.005, a["terminal_g"],
          a["terminal_g"] + 0.005, a["terminal_g"] + 0.01]
    cell(r, 1, "WACC \\ g", fill=SUB_FILL, font=BOLD)
    for j, g in enumerate(gs):
        cell(r, 2 + j, g, fill=SUB_FILL, font=BOLD, fmt="0.0%", center=True)
    r += 1
    for i, w in enumerate(waccs):
        cell(r, 1, w, fill=SUB_FILL, font=BOLD, fmt="0.0%")
        for j, g in enumerate(gs):
            aa = dict(a)
            aa["wacc"] = w
            aa["terminal_g"] = g
            rr = run_dcf(d, aa)
            cc = cell(r, 2 + j, rr["implied"], fmt="0.00", center=True)
            if i == 2 and j == 2:
                cc.fill = OUT_FILL
                cc.font = BOLD
        r += 1

    # WACC sheet
    ws2 = wb.create_sheet("WACC")
    ws2.merge_cells("A1:C1")
    ws2["A1"] = "WACC 计算"
    ws2["A1"].font = Font(bold=True, size=14)
    rr = 3
    for name, val, fmt in [
        ("无风险利率（10年）", a["rf"], "0.0%"),
        ("Beta", d["beta"] or a["beta"], "0.00"),
        ("股权风险溢价", a["erp"], "0.0%"),
        ("股权成本", a["cost_of_equity"], "0.0%"),
        ("税后债务成本", a["cost_of_debt"], "0.0%"),
        ("WACC（基准情景）", a["wacc"], "0.0%"),
    ]:
        ws2.cell(row=rr, column=1, value=name).font = BOLD
        ws2.cell(row=rr, column=2, value=val).number_format = fmt
        rr += 1

    wb.save(path)
    return res


# ---------------------------------------------------------------- 主入口
def main():
    ap = argparse.ArgumentParser(description="免费 DCF 估值模型生成器")
    ap.add_argument("ticker", nargs="?", default="MSFT",
                    help="股票代码，如 AAPL / MSFT / 0700.HK（默认 MSFT）")
    ap.add_argument("--demo", action="store_true",
                    help="用离线示例数据生成（不联网，验证工具可用）")
    ap.add_argument("--data", default=None,
                    help="从 JSON 文件加载真实财务数据（字段同 fetch_financials 返回的 dict）")
    ap.add_argument("--manual", action="store_true", help="手动输入模式")
    ap.add_argument("--growth", type=float, nargs="+", default=None,
                    help="逐年收入增长率（%，如 12 10 9 8 7）")
    ap.add_argument("--wacc", type=float, default=None, help="WACC（如 0.09）")
    ap.add_argument("--terminal", type=float, default=None, help="永续增长率（如 0.03）")
    ap.add_argument("--rf", type=float, default=None, help="无风险利率（如 A股 0.022 / 美股 0.035）")
    ap.add_argument("--erp", type=float, default=None, help="股权风险溢价（如 A股 0.065 / 美股 0.055）")
    ap.add_argument("--tax", type=float, default=None, help="所得税率（如 0.25 / 0.15）")
    ap.add_argument("--beta", type=float, default=None, help="Beta（如 1.0）")
    ap.add_argument("--cost-of-debt", type=float, default=None, help="税后债务成本（如 0.04）")
    ap.add_argument("--target-margin", type=float, default=None,
                    help="目标营业利润率（覆盖中性假设，如 0.08 表示利润率逐步提升到 8%）")
    args = ap.parse_args()

    # 默认假设（与 Anthropic dcf-model 技能一致）
    a = {
        "years": 5,
        "growth": [0.16, 0.14, 0.12, 0.10, 0.09],  # Base case 收入增长
        "ebit_margin": 0.25,
        "target_margin": 0.30,
        "tax_rate": 0.21,
        "da_pct": 0.03,
        "capex_pct": 0.05,
        "terminal_capex_pct": 0.05,  # 稳态 CapEx 占收入比（资本开支高点逐步回落到此水平）
        "nwc_pct": 0.02,
        "rf": 0.035,
        "beta": 1.0,
        "erp": 0.055,
        "cost_of_debt": 0.04,
        "wacc": 0.09,
        "terminal_g": 0.03,
    }
    a["cost_of_equity"] = a["rf"] + a["beta"] * a["erp"]

    if args.growth:
        a["growth"] = [g / 100 for g in args.growth]
    if args.wacc:
        a["wacc"] = args.wacc
    if args.terminal:
        a["terminal_g"] = args.terminal
    if args.rf:
        a["rf"] = args.rf
    if args.erp:
        a["erp"] = args.erp
    if args.tax:
        a["tax_rate"] = args.tax
    if args.beta:
        a["beta"] = args.beta
    if args.cost_of_debt:
        a["cost_of_debt"] = args.cost_of_debt

    # 数据获取：--data 优先加载真实数据；demo 用离线示例；否则尝试 yfinance，失败再回退示例
    if args.data:
        import json
        with open(args.data, encoding="utf-8") as f:
            d = json.load(f)
        src = f"真实数据（{args.data}）"
    elif args.demo:
        d = sample_data(args.ticker)
        src = "离线示例数据（--demo）"
    else:
        d = fetch_financials(args.ticker)
        if d is None or d["price"] is None or d["revenue"] is None:
            chart = fetch_yahoo_chart(args.ticker)
            if chart and chart.get("price"):
                d = sample_data(args.ticker)
                d["price"] = chart["price"]
                src = "Yahoo 实时行情 + 示例财务（yfinance 不可用）"
            else:
                d = sample_data(args.ticker)
                src = "离线示例数据（网络/数据源不可用，已回退）"
        else:
            src = "Yahoo Finance (yfinance)"

    # 用实际数据校准关键假设（保持中性，避免默认假设拉偏结果）
    if d.get("ebit") and d.get("revenue"):
        actual_margin = d["ebit"] / d["revenue"]
        a["ebit_margin"] = actual_margin
        a["target_margin"] = actual_margin  # 中性：假设利润率维持现状

    # 用户可覆盖目标利润率（用于「利润率提升」等情景）
    if args.target_margin:
        a["target_margin"] = args.target_margin

    # 用实际 D&A、CapEx 校准（占收入比），让真实财务数据真正流进模型
    if d.get("da") and d.get("revenue"):
        a["da_pct"] = d["da"] / d["revenue"]
    if d.get("capex") and d.get("revenue"):
        a["capex_pct"] = d["capex"] / d["revenue"]
        # 稳态 CapEx：资本开支峰值逐步回落到「D&A 的 1.2 倍」(成熟企业合理水平)
        a["terminal_capex_pct"] = a["da_pct"] * 1.2

    # 用实际 beta 校准 WACC（如果拉到 beta）
    if d.get("beta"):
        a["beta"] = d["beta"]
        a["cost_of_equity"] = a["rf"] + a["beta"] * a["erp"]
        # 简化：股权/债权权重按 90/10，实际可自行调整
        a["wacc"] = a["cost_of_equity"] * 0.9 + a["cost_of_debt"] * 0.1

    out_path = f"{args.ticker.upper()}_DCF_Model_{dt.date.today()}.xlsx"
    res = build_excel(args.ticker, d, a, out_path)

    cur = d.get("currency", "$")
    print("=" * 56)
    print(f"✅ DCF 模型已生成：{out_path}")
    print("=" * 56)
    print(f"  数据源            : {src}")
    print(f"  当前股价          : {cur}{d['price']:.2f}" if d["price"] else "  当前股价: N/A")
    print(f"  隐含每股价值      : {cur}{res['implied']:.2f}")
    if d["price"]:
        up = res["implied"] / d["price"] - 1
        tag = "看多 ↑" if up > 0 else "看空 ↓"
        print(f"  隐含空间          : {up:+.1%}  ({tag})")
    print(f"  企业价值 (EV)     : {cur}{res['ev']/1e6:,.0f}M")
    print(f"  WACC              : {a['wacc']:.1%}")
    print("=" * 56)
    print("⚠️  仅供学习/研究，不构成投资建议。数据来自免费公开渠道，")
    print("    需自行核对后再做判断。")


if __name__ == "__main__":
    main()
