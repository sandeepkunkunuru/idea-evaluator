"""Portfolio-screening methods for ranking many ideas cheaply.

Methods:
  - Modified RICE  (Reach x Impact x Confidence / Effort, weighted variant)
  - WSJF          (Weighted Shortest Job First, SAFe)
  - CD3           (Cost of Delay / Duration, Reinertsen)
  - EMV           (Expected Monetary Value)

Each method ranks the same ideas; divergence in rank flags scoring-sensitive
ideas that deserve a deep dive.

Input idea schema (all USD, all years; "_yrs" suffix for time):
  name                          str
  monetary_impact               annual revenue/savings if successful
  adoption                      0..1 fraction reached
  business_critical             0..1 strategic criticality
  ownership                     0/1 internal vs external dev
  risk_pos                      0..1 probability of success
  time_of_replacement_yrs       years before a competitor replaces you
  time_to_commercialize_yrs     years from now to revenue
  cost_to_develop               total dev + go-to-market cost
  # Optional WSJF/CD3 overrides — if absent we derive them:
  time_criticality              0..1 (default = 1/time_of_replacement_yrs, clipped)
  risk_reduction                0..1 (default = business_critical)
"""
from dataclasses import dataclass
from typing import Dict, List


@dataclass
class IdeaScore:
    name: str
    modified_rice: float
    wsjf: float
    cd3: float
    emv: float
    inputs: dict


def _safe(d, k, default=0.0):
    v = d.get(k, default)
    try:
        return float(v) if v is not None else default
    except (TypeError, ValueError):
        return default


def modified_rice(i: Dict) -> float:
    """Samyama variant: (MonImpact * Adoption * BizCritical * Ownership * PoS * TimeReplacement)
       / (TimeToCommercialize * Cost)."""
    num = (_safe(i, "monetary_impact") * _safe(i, "adoption")
           * _safe(i, "business_critical") * _safe(i, "ownership")
           * _safe(i, "risk_pos") * _safe(i, "time_of_replacement_yrs"))
    denom = _safe(i, "time_to_commercialize_yrs", 1) * _safe(i, "cost_to_develop", 1)
    return num / denom if denom > 0 else 0.0


def wsjf(i: Dict) -> float:
    """WSJF = (BusinessValue + TimeCriticality + RiskReduction) / JobDuration.
    We unit-normalize BV to a 0..100 scale so the three CoD components are
    comparable, per SAFe guidance."""
    annual_value = _safe(i, "monetary_impact") * _safe(i, "adoption") * _safe(i, "risk_pos")
    # Normalize to a 0..100 score against a $10M reference deal
    bv = min(100.0, annual_value / 1e5)
    tor = _safe(i, "time_of_replacement_yrs", 5)
    tc_default = max(0.0, min(1.0, 1.0 / max(tor, 0.5)))
    time_crit = _safe(i, "time_criticality", tc_default) * 100
    risk_red = _safe(i, "risk_reduction", _safe(i, "business_critical")) * 100
    duration = _safe(i, "time_to_commercialize_yrs", 1)
    return (bv + time_crit + risk_red) / duration if duration > 0 else 0.0


def cd3(i: Dict) -> float:
    """CD3 = Cost of Delay ($/yr) / Duration (yrs). Cost of Delay here =
    annual expected value lost while you wait."""
    cod = _safe(i, "monetary_impact") * _safe(i, "adoption") * _safe(i, "risk_pos")
    duration = _safe(i, "time_to_commercialize_yrs", 1)
    return cod / duration if duration > 0 else 0.0


def emv(i: Dict) -> float:
    """EMV = PoS * (annual value * useful life) - cost. A direct $-NPV-ish
    comparison; uses time_of_replacement as a proxy for useful life."""
    annual = _safe(i, "monetary_impact") * _safe(i, "adoption")
    life = _safe(i, "time_of_replacement_yrs", 3)
    return _safe(i, "risk_pos") * annual * life - _safe(i, "cost_to_develop")


def score_all(ideas: List[Dict]) -> List[IdeaScore]:
    out = []
    for idea in ideas:
        out.append(IdeaScore(
            name=str(idea.get("name", "?")),
            modified_rice=modified_rice(idea),
            wsjf=wsjf(idea),
            cd3=cd3(idea),
            emv=emv(idea),
            inputs=idea,
        ))
    return out


def rank_table(scored: List[IdeaScore]) -> List[Dict]:
    """Return rows with per-method ranks and a rank-divergence signal."""
    def ranks(values):
        order = sorted(range(len(values)), key=lambda i: values[i], reverse=True)
        rk = [0] * len(values)
        for pos, idx in enumerate(order):
            rk[idx] = pos + 1
        return rk

    r_rice = ranks([s.modified_rice for s in scored])
    r_wsjf = ranks([s.wsjf for s in scored])
    r_cd3  = ranks([s.cd3 for s in scored])
    r_emv  = ranks([s.emv for s in scored])
    rows = []
    for i, s in enumerate(scored):
        rs = [r_rice[i], r_wsjf[i], r_cd3[i], r_emv[i]]
        rows.append({
            "name": s.name,
            "modified_rice": s.modified_rice,
            "wsjf": s.wsjf,
            "cd3": s.cd3,
            "emv": s.emv,
            "rank_rice": r_rice[i], "rank_wsjf": r_wsjf[i],
            "rank_cd3": r_cd3[i], "rank_emv": r_emv[i],
            "rank_spread": max(rs) - min(rs),
            "avg_rank": sum(rs) / 4,
        })
    rows.sort(key=lambda r: r["avg_rank"])
    return rows
