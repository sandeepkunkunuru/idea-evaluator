"""Founder equity split using a Slicing-Pie-style contribution-units model,
net of an ESOP option pool.

Contribution units per founder =
    time_pct * salary_forgone_usd        (sweat equity at fair-market value)
  + capital_usd                          (cash in)
  + ip_score * IP_UNIT_VALUE             (IP / prior work brought in)
  + role_criticality * ROLE_UNIT_VALUE   (irreplaceability premium)

Each founder's share of the founder pool = their units / sum of units.
Final equity = founder_pool * share_of_pool, where founder_pool = 1 - esop_pool.

This is intentionally transparent and editable in the xlsx so co-founders
can argue about it directly with numbers rather than feelings.
"""
from typing import Dict, List

IP_UNIT_VALUE = 100_000        # USD-equivalent per IP score point (1-5)
ROLE_UNIT_VALUE = 50_000       # USD-equivalent per role-criticality point (1-5)
N_FOUNDERS = 4                 # slots; unused ones get zeroed by user


def founder_field_keys(i: int) -> Dict[str, str]:
    """Returns the schema keys for founder i (1-indexed)."""
    return {
        "name":            f"founder{i}_name",
        "time_pct":        f"founder{i}_time_pct",
        "capital_usd":     f"founder{i}_capital_usd",
        "ip_score":        f"founder{i}_ip_score",
        "salary_forgone":  f"founder{i}_salary_forgone_usd",
        "role_criticality":f"founder{i}_role_criticality",
    }


def units(founder: Dict) -> float:
    return (float(founder.get("time_pct", 0)) * float(founder.get("salary_forgone", 0))
            + float(founder.get("capital_usd", 0))
            + float(founder.get("ip_score", 0)) * IP_UNIT_VALUE
            + float(founder.get("role_criticality", 0)) * ROLE_UNIT_VALUE)


def compute(inputs: Dict) -> List[Dict]:
    esop = float(inputs.get("esop_pool_pct", 0.15))
    founder_pool = max(0.0, 1.0 - esop)
    founders = []
    for i in range(1, N_FOUNDERS + 1):
        k = founder_field_keys(i)
        f = {kk: inputs.get(v) for kk, v in k.items()}
        f["units"] = units(f)
        founders.append(f)
    total_units = sum(f["units"] for f in founders) or 1.0
    for f in founders:
        f["share_of_founder_pool"] = f["units"] / total_units
        f["final_equity"] = founder_pool * f["share_of_founder_pool"]
    return founders


def print_table(founders: List[Dict], esop_pct: float) -> None:
    print(f"\nEquity split (ESOP pool: {esop_pct:.0%}):")
    print(f"  {'Name':<20} {'Units':>12} {'Founder pool %':>15} {'Final equity %':>15}")
    for f in founders:
        if (f["units"] or 0) <= 0:
            continue
        print(f"  {str(f.get('name','?'))[:20]:<20} "
              f"{f['units']:>12,.0f} "
              f"{f['share_of_founder_pool']*100:>14.2f}% "
              f"{f['final_equity']*100:>14.2f}%")
    print(f"  {'ESOP pool':<20} {'':>12} {'':>15} {esop_pct*100:>14.2f}%")
