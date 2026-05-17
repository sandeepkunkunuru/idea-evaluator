"""Comprehensive rubric. Weighted average of category scores (1..5), then a
0..100 normalized composite. Each category aggregates its sub-scores.

The same weights are written into the xlsx so the Scorecard sheet recomputes
live when inputs change.
"""
from typing import Dict, List, Tuple

# (category, weight 0..1, [sub-score keys])
CATEGORIES: List[Tuple[str, float, List[str]]] = [
    ("Problem",         0.12, ["score_problem_severity", "score_problem_frequency", "score_willingness_to_pay"]),
    ("Market & Timing", 0.12, ["score_market_timing"]),  # market size handled separately below
    ("Moat",            0.14, ["score_moat_tech", "score_moat_data", "score_moat_brand", "score_ip_strength"]),
    ("Team",            0.15, ["score_team_domain", "score_team_execution", "score_team_completeness"]),
    ("GTM",             0.08, ["score_gtm_clarity"]),
    ("Unit Economics",  0.10, ["score_unit_economics", "score_capital_efficiency"]),
    ("Scalability",     0.07, ["score_scalability"]),
    ("Regulatory",      0.05, ["score_regulatory_risk"]),
    ("AI Resilience",   0.17, ["score_ai_resilience"]),
]


def verdict(score_100: float) -> str:
    if score_100 >= 75:
        return "STRONG GO - pursue aggressively"
    if score_100 >= 60:
        return "GO - pursue with focused de-risking"
    if score_100 >= 45:
        return "CONDITIONAL - fix top 2 weaknesses before committing"
    if score_100 >= 30:
        return "WEAK - pivot or redefine wedge"
    return "NO-GO - fundamental issues"


def compute(inputs: Dict) -> Dict:
    cats = []
    weighted_sum = 0.0
    for name, weight, keys in CATEGORIES:
        vals = [float(inputs.get(k, 3)) for k in keys]
        avg = sum(vals) / len(vals) if vals else 0.0
        cats.append({"name": name, "weight": weight, "avg": avg, "keys": keys})
        weighted_sum += weight * avg
    # market-size bonus (log-scaled vs $1B TAM)
    import math
    tam = max(1.0, float(inputs.get("tam_usd", 1)))
    market_bonus = max(-1.0, min(1.0, math.log10(tam / 1e9)))  # -1..+1
    composite_5 = weighted_sum + 0.15 * market_bonus
    composite_5 = max(1.0, min(5.0, composite_5))
    score_100 = (composite_5 - 1) / 4 * 100
    # Top strengths / weaknesses (by sub-score)
    sub_keyed = []
    for cat in cats:
        for k in cat["keys"]:
            sub_keyed.append((k, float(inputs.get(k, 3)), cat["name"]))
    sub_keyed.sort(key=lambda x: x[1])
    weaknesses = sub_keyed[:3]
    strengths = list(reversed(sub_keyed[-3:]))
    return {
        "categories": cats,
        "composite_5": composite_5,
        "score_100": score_100,
        "verdict": verdict(score_100),
        "market_bonus": market_bonus,
        "strengths": strengths,
        "weaknesses": weaknesses,
    }
