"""Glossary used in the Definitions sheet of every workbook.

Structure: list of (Term, Plain definition, How it's used in this tool).
Also exposes a rubric-to-screening cross-map so users can see how the
deep-evaluation dimensions correspond to RICE/WSJF/CD3/EMV concepts.
"""

GLOSSARY = [
    # --- Screening terms (from PatentValuationTool.xlsx + extensions) ---
    ("Reach (RICE)", "How many customers a project will impact.",
     "Captured as Monetary impact × Adoption fraction."),
    ("Impact (RICE)", "How much a project shifts a strategic goal.",
     "Captured as Business-critical × Ownership."),
    ("Confidence (RICE)", "How sure we are in the estimates.",
     "Captured as Risk PoS × Time of replacement."),
    ("Effort (RICE)", "Total time and money required.",
     "Captured as Time to commercialize × Cost to develop."),
    ("Modified RICE", "Samyama variant: (Reach × Impact × Confidence) / Effort.",
     "Quick portfolio-level screen — higher is better."),
    ("WSJF", "Weighted Shortest Job First (SAFe): "
             "(Business Value + Time Criticality + Risk Reduction) / Duration.",
     "Use when you care about delivery sequencing and time-to-market."),
    ("CD3", "Cost of Delay ÷ Duration (Reinertsen).",
     "Use when you can estimate $/year lost by delaying the project."),
    ("EMV", "Expected Monetary Value: PoS × (annual value × life) − cost.",
     "Use when comparing projects in dollars rather than scores."),

    # --- Deep-evaluation terms ---
    ("TAM/SAM/SOM", "Total / Serviceable / Obtainable market.",
     "Drives the market-size bonus on the composite score."),
    ("Moat", "Durable defensibility (tech, data, brand, IP).",
     "Four sub-scores feed the Moat category (14% weight)."),
    ("AI Resilience", "Will commodity LLMs help or eat this wedge in 24 months?",
     "Heaviest single weight (17%) in the rubric — by design."),
    ("Unit economics", "Per-customer revenue minus per-customer cost.",
     "Combined with capital efficiency into a 10% category."),
    ("Capital efficiency", "Revenue generated per dollar of capital consumed.",
     "Higher = less dilution / more optionality on financing."),

    # --- Valuation terms ---
    ("DCF", "Discounted Cash Flow: PV of projected FCF + terminal value.",
     "Best for businesses with credible 5-yr forecasts."),
    ("Comparables", "Apply sector revenue / EBITDA multiples to NTM numbers.",
     "Best when peer multiples are observable."),
    ("Berkus", "Pre-revenue startup valuation, max ~$2.5M across 5 axes.",
     "Use at idea / prototype stage."),
    ("Scorecard (Payne)", "Regional avg pre-money × (rubric score / 3).",
     "Use at seed; anchors to local market norms."),
    ("VC method", "Exit value ÷ target ROI multiple.",
     "Use when you have a credible exit comp."),
    ("First Chicago", "Weighted success/base/downside scenarios.",
     "Use when outcomes are bimodal."),
    ("Cost approach (IP)", "Total R&D + filing cost as patent value floor.",
     "Conservative — ignores commercial upside."),
    ("Market approach (IP)", "Comparable patent sale price.",
     "Only as good as the comp."),
    ("Relief-from-royalty", "PV of royalty payments avoided by owning the patent.",
     "Royalty rate × attributable revenue, discounted over useful life."),
    ("Income approach (IP)", "PV of incremental income enabled by the patent.",
     "Uses 20% margin proxy unless overridden."),
]

# How rubric categories map to screening concepts. Use this when you screen
# many ideas first, then deep-evaluate top N: the same evidence should
# inform both lenses.
RUBRIC_TO_SCREENING = [
    ("Rubric category",     "RICE/WSJF mapping",
     "Cross-check"),
    ("Problem",             "Reach (severity × frequency × WTP)",
     "Low problem score + high RICE Reach => suspect inflated adoption."),
    ("Market & Timing",     "WSJF Time Criticality",
     "Strong timing => higher WSJF urgency, shorter Time of Replacement."),
    ("Moat",                "Time of Replacement (years)",
     "Weak moat (<3 yrs) caps WSJF and CD3 even if value is high."),
    ("Team",                "Confidence / Risk PoS",
     "Team gaps reduce risk_pos in EMV and Confidence in RICE."),
    ("Unit Economics",      "Cost of Delay numerator",
     "Drives the $-of-delay term in CD3."),
    ("AI Resilience",       "Time of Replacement (AI-driven)",
     "If AI eats this in 24 months, set time_of_replacement ≤ 2."),
    ("Regulatory",          "Risk PoS",
     "Heavy regulation reduces PoS; multiplies through EMV and WSJF."),
]
