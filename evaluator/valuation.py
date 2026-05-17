"""Valuation calculations (Python side, for the CLI report).
The xlsx writer re-encodes the same logic as live Excel formulas.
"""
from typing import Dict


def dcf(inp: Dict) -> Dict:
    rev = [float(inp["rev_y1"])]
    growths = [inp["rev_growth_y2"], inp["rev_growth_y3"],
               inp["rev_growth_y4"], inp["rev_growth_y5"]]
    for g in growths:
        rev.append(rev[-1] * (1 + float(g)))
    gm = float(inp["gross_margin"])
    opex_pct = float(inp["opex_pct_rev"])
    tax = float(inp["tax_rate"])
    capex_pct = float(inp["capex_pct_rev"])
    nwc_pct = float(inp["nwc_pct_rev"])
    wacc = float(inp["discount_rate"])
    tg = float(inp["terminal_growth"])

    fcfs = []
    prev_nwc = 0.0
    for r in rev:
        ebit = r * gm - r * opex_pct
        nopat = ebit * (1 - tax)
        capex = r * capex_pct
        nwc = r * nwc_pct
        d_nwc = nwc - prev_nwc
        prev_nwc = nwc
        fcf = nopat - capex - d_nwc
        fcfs.append(fcf)

    pv = sum(f / ((1 + wacc) ** (i + 1)) for i, f in enumerate(fcfs))
    tv = fcfs[-1] * (1 + tg) / max(1e-6, (wacc - tg))
    pv_tv = tv / ((1 + wacc) ** len(fcfs))
    ev = pv + pv_tv
    equity = ev - float(inp.get("net_debt_usd", 0))
    return {"revenues": rev, "fcfs": fcfs, "ev": ev, "equity": equity, "terminal_value": tv}


def comparables(inp: Dict) -> Dict:
    by_rev = float(inp["comp_rev_multiple"]) * float(inp["comp_ntm_revenue"])
    by_ebitda = float(inp["comp_ebitda_multiple"]) * float(inp["comp_ntm_ebitda"])
    return {"by_revenue": by_rev, "by_ebitda": by_ebitda,
            "midpoint": (by_rev + by_ebitda) / 2}


def berkus(inp: Dict) -> float:
    return sum(float(inp[k]) for k in
               ["berkus_idea", "berkus_prototype", "berkus_team",
                "berkus_relationships", "berkus_sales"])


def scorecard(inp: Dict, rubric_5: float) -> float:
    # Bill Payne style: multiplier ~ rubric/3 (3=avg)
    mult = rubric_5 / 3.0
    return float(inp["scorecard_avg_premoney"]) * mult


def vc_method(inp: Dict) -> float:
    exit_v = float(inp["vc_exit_value"])
    roi = float(inp["vc_target_roi"])
    return exit_v / max(1e-6, roi)


def first_chicago(dcf_equity: float, comp_mid: float, berkus_v: float) -> Dict:
    # Weighted: success 25%, base 50%, downside 25%
    success = max(dcf_equity, comp_mid) * 1.5
    base = (dcf_equity + comp_mid) / 2
    downside = berkus_v
    weighted = 0.25 * success + 0.5 * base + 0.25 * downside
    return {"success": success, "base": base, "downside": downside, "weighted": weighted}


def patent_valuation(inp: Dict) -> Dict:
    cost = float(inp["patent_dev_cost_usd"])
    market = float(inp["patent_comparable_sale_usd"])
    royalty = float(inp["patent_royalty_rate"])
    rev_attr = float(inp["patent_revenue_attributable"])
    life = int(inp["patent_useful_life_years"])
    wacc = float(inp["discount_rate"])
    # Relief-from-royalty: PV of (royalty_rate * revenue) over life
    rfr = sum((royalty * rev_attr) / ((1 + wacc) ** t) for t in range(1, life + 1))
    # Income approach (simple): PV of attributable revenue * 20% margin
    income = sum((0.20 * rev_attr) / ((1 + wacc) ** t) for t in range(1, life + 1))
    return {"cost_approach": cost, "market_approach": market,
            "relief_from_royalty": rfr, "income_approach": income,
            "midpoint": (cost + market + rfr + income) / 4}


def all_valuations(inp: Dict, rubric_5: float) -> Dict:
    d = dcf(inp)
    c = comparables(inp)
    b = berkus(inp)
    s = scorecard(inp, rubric_5)
    v = vc_method(inp)
    fc = first_chicago(d["equity"], c["midpoint"], b)
    p = patent_valuation(inp)
    return {"dcf": d, "comparables": c, "berkus": b, "scorecard": s,
            "vc_method": v, "first_chicago": fc, "patent": p}
