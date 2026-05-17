"""Input schema: every field has a key, prompt, type, default, and help text.

This single source of truth drives:
  - interactive CLI prompting
  - JSON input validation
  - the Inputs sheet in the xlsx (named ranges)
  - rubric scoring, org sizing, and valuation formulas
"""
from dataclasses import dataclass
from typing import Any, Callable, Optional

NUM = "number"
INT = "int"
STR = "string"
TEXT = "text"        # multi-line free text
PCT = "percent"      # stored as 0..1
SCORE = "score"      # 1..5 self-rating
CHOICE = "choice"


@dataclass
class Field:
    key: str
    prompt: str
    kind: str
    default: Any = None
    help: str = ""
    choices: Optional[list] = None
    section: str = "general"


# -------- Idea / qualitative -----------------------------------------------
IDEA_FIELDS = [
    Field("idea_name", "Idea / company name", STR, "Untitled", section="idea"),
    Field("one_liner", "One-line description", STR, "", section="idea"),
    Field("problem", "Problem being solved (2-4 sentences)", TEXT, "", section="idea"),
    Field("solution", "Proposed solution", TEXT, "", section="idea"),
    Field("target_customer", "Primary target customer / ICP", STR, "", section="idea"),
    Field("geography", "Initial geography", STR, "India", section="idea"),
    Field("stage", "Stage", CHOICE, "idea",
          choices=["idea", "prototype", "mvp", "early-revenue", "scaling"], section="idea"),
]

# -------- Market ------------------------------------------------------------
MARKET_FIELDS = [
    Field("tam_usd", "TAM in USD (total addressable market)", NUM, 1_000_000_000, section="market"),
    Field("sam_usd", "SAM in USD (serviceable addressable)", NUM, 100_000_000, section="market"),
    Field("som_usd", "SOM in USD (serviceable obtainable, yr-5)", NUM, 10_000_000, section="market"),
    Field("market_growth_pct", "Market CAGR (e.g. 0.15 for 15%)", PCT, 0.12, section="market"),
    Field("competitors_count", "Number of meaningful competitors", INT, 5, section="market"),
    Field("competitor_strength", "Strongest competitor strength (1=weak,5=dominant)", SCORE, 3, section="market"),
]

# -------- Self-rated rubric scores 1..5 ------------------------------------
RUBRIC_FIELDS = [
    Field("score_problem_severity", "Problem severity / pain (1-5)", SCORE, 3, section="rubric",
          help="How acute is the pain? 5 = hair-on-fire."),
    Field("score_problem_frequency", "Problem frequency (1-5)", SCORE, 3, section="rubric"),
    Field("score_willingness_to_pay", "Customer willingness to pay (1-5)", SCORE, 3, section="rubric"),
    Field("score_market_timing", "Market timing / tailwinds (1-5)", SCORE, 3, section="rubric"),
    Field("score_moat_tech", "Technology moat / defensibility (1-5)", SCORE, 3, section="rubric"),
    Field("score_moat_data", "Data moat / network effects (1-5)", SCORE, 3, section="rubric"),
    Field("score_moat_brand", "Brand / distribution moat (1-5)", SCORE, 2, section="rubric"),
    Field("score_team_domain", "Team domain expertise (1-5)", SCORE, 3, section="rubric"),
    Field("score_team_execution", "Team execution track record (1-5)", SCORE, 3, section="rubric"),
    Field("score_team_completeness", "Team completeness (1-5)", SCORE, 3, section="rubric"),
    Field("score_gtm_clarity", "GTM clarity & channel fit (1-5)", SCORE, 3, section="rubric"),
    Field("score_unit_economics", "Unit economics health (1-5)", SCORE, 3, section="rubric"),
    Field("score_capital_efficiency", "Capital efficiency (1-5)", SCORE, 3, section="rubric"),
    Field("score_regulatory_risk", "Regulatory safety (5=safe,1=heavy risk)", SCORE, 4, section="rubric"),
    Field("score_ai_resilience", "AI-disruption resilience (5=AI helps you,1=AI kills you)", SCORE, 3, section="rubric",
          help="Will commodity LLMs eat your wedge in 24 months? Higher = safer/benefiting."),
    Field("score_scalability", "Scalability of delivery (1-5)", SCORE, 3, section="rubric"),
    Field("score_ip_strength", "IP / patent strength (1-5)", SCORE, 2, section="rubric"),
]

# -------- Financials for DCF / comparables ----------------------------------
FIN_FIELDS = [
    Field("rev_y1", "Projected revenue Year 1 (USD)", NUM, 100_000, section="financials"),
    Field("rev_growth_y2", "Revenue growth Y2", PCT, 2.0, section="financials"),
    Field("rev_growth_y3", "Revenue growth Y3", PCT, 1.5, section="financials"),
    Field("rev_growth_y4", "Revenue growth Y4", PCT, 1.0, section="financials"),
    Field("rev_growth_y5", "Revenue growth Y5", PCT, 0.6, section="financials"),
    Field("gross_margin", "Gross margin", PCT, 0.7, section="financials"),
    Field("opex_pct_rev", "Opex as % of revenue (steady-state)", PCT, 0.5, section="financials"),
    Field("tax_rate", "Tax rate", PCT, 0.25, section="financials"),
    Field("capex_pct_rev", "CapEx % revenue", PCT, 0.05, section="financials"),
    Field("nwc_pct_rev", "Net working capital % revenue", PCT, 0.05, section="financials"),
    Field("discount_rate", "Discount rate / WACC", PCT, 0.25, section="financials"),
    Field("terminal_growth", "Terminal growth", PCT, 0.03, section="financials"),
    Field("net_debt_usd", "Net debt (USD)", NUM, 0, section="financials"),
]

# -------- Comparables -------------------------------------------------------
COMP_FIELDS = [
    Field("comp_rev_multiple", "Sector revenue multiple (NTM)", NUM, 6.0, section="comparables"),
    Field("comp_ebitda_multiple", "Sector EBITDA multiple", NUM, 18.0, section="comparables"),
    Field("comp_ntm_revenue", "NTM revenue for multiple (USD)", NUM, 500_000, section="comparables"),
    Field("comp_ntm_ebitda", "NTM EBITDA for multiple (USD)", NUM, 50_000, section="comparables"),
]

# -------- Early-stage (Berkus / Scorecard / VC) -----------------------------
EARLY_FIELDS = [
    Field("berkus_idea", "Berkus: sound idea value (USD, max ~500k)", NUM, 300_000, section="early"),
    Field("berkus_prototype", "Berkus: prototype value (USD, max ~500k)", NUM, 300_000, section="early"),
    Field("berkus_team", "Berkus: quality team value", NUM, 400_000, section="early"),
    Field("berkus_relationships", "Berkus: strategic relationships", NUM, 200_000, section="early"),
    Field("berkus_sales", "Berkus: product rollout / sales", NUM, 200_000, section="early"),
    Field("scorecard_avg_premoney", "Scorecard: regional avg pre-money (USD)", NUM, 2_500_000, section="early"),
    Field("vc_exit_value", "VC method: expected exit value (USD)", NUM, 50_000_000, section="early"),
    Field("vc_target_roi", "VC method: target ROI multiple", NUM, 10.0, section="early"),
    Field("vc_years_to_exit", "VC method: years to exit", NUM, 6, section="early"),
]

# -------- Patent / IP -------------------------------------------------------
PATENT_FIELDS = [
    Field("patent_count", "Number of patents / filings", INT, 0, section="patent"),
    Field("patent_dev_cost_usd", "Cost approach: total R&D + filing cost", NUM, 0, section="patent"),
    Field("patent_comparable_sale_usd", "Market approach: comparable patent sale", NUM, 0, section="patent"),
    Field("patent_royalty_rate", "Relief-from-royalty: royalty rate", PCT, 0.05, section="patent"),
    Field("patent_revenue_attributable", "Revenue attributable to patent (USD/yr)", NUM, 0, section="patent"),
    Field("patent_useful_life_years", "Useful life (yrs)", INT, 10, section="patent"),
]

# -------- Org structure inputs ---------------------------------------------
ORG_FIELDS = [
    Field("target_headcount_y1", "Target headcount end of Y1", INT, 8, section="org"),
    Field("target_headcount_y3", "Target headcount end of Y3", INT, 30, section="org"),
    Field("avg_fully_loaded_cost", "Avg fully-loaded cost / FTE (USD)", NUM, 40_000, section="org"),
    Field("ai_automation_target", "Target % of routine work automated by AI", PCT, 0.5, section="org"),
]


# -------- Equity split --------------------------------------------------
EQUITY_FIELDS = [
    Field("esop_pool_pct", "ESOP / option pool size", PCT, 0.15, section="equity"),
]
for _i, _default_name in enumerate(
        ["Founder A", "Founder B", "Founder C", "Founder D"], start=1):
    EQUITY_FIELDS.extend([
        Field(f"founder{_i}_name", f"Founder {_i} name", STR, _default_name, section="equity"),
        Field(f"founder{_i}_time_pct",
              f"Founder {_i} time commitment (0-1)", PCT, 1.0 if _i <= 2 else 0.0,
              section="equity"),
        Field(f"founder{_i}_capital_usd",
              f"Founder {_i} cash contributed (USD)", NUM, 0, section="equity"),
        Field(f"founder{_i}_ip_score",
              f"Founder {_i} IP / prior work (1-5, 0 if N/A)", SCORE,
              3 if _i <= 2 else 0, section="equity"),
        Field(f"founder{_i}_role_criticality",
              f"Founder {_i} role criticality (1-5, 0 if N/A)", SCORE,
              4 if _i <= 2 else 0, section="equity"),
        Field(f"founder{_i}_salary_forgone_usd",
              f"Founder {_i} annual salary forgone (USD)", NUM,
              100_000 if _i <= 2 else 0, section="equity"),
    ])

ALL_FIELDS: list[Field] = (
    IDEA_FIELDS + MARKET_FIELDS + RUBRIC_FIELDS + FIN_FIELDS
    + COMP_FIELDS + EARLY_FIELDS + PATENT_FIELDS + ORG_FIELDS
    + EQUITY_FIELDS
)

FIELDS_BY_KEY = {f.key: f for f in ALL_FIELDS}


def coerce(field: Field, raw: Any) -> Any:
    if raw is None or raw == "":
        return field.default
    if field.kind == INT:
        return int(float(raw))
    if field.kind in (NUM, PCT):
        return float(raw)
    if field.kind == SCORE:
        v = int(float(raw))
        return max(0, min(5, v))
    if field.kind == CHOICE:
        s = str(raw).strip().lower()
        if field.choices and s not in field.choices:
            return field.default
        return s
    return str(raw)


def defaults() -> dict:
    return {f.key: f.default for f in ALL_FIELDS}
