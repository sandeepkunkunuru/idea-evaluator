"""AI-proof org structure generator.

Produces a role list with:
  - function, role, level
  - headcount Y1 and Y3 (scaled to inputs.target_headcount_*)
  - AI augmentation pattern (what AI does for the role)
  - automation potential (0..1)
  - reporting line

The total Y1/Y3 headcount is normalized to the user's targets so the model
auto-scales to their ambition.
"""
from typing import Dict, List


# Base template: weight = relative staffing share. The structure favors small
# pods, AI-augmented ICs, and thin middle management — i.e. AI-proof shape.
TEMPLATE = [
    # (function, role, level, y1_weight, y3_weight, automation, reports_to, ai_use)
    # Every title is intentionally AI-native: roles framed around what humans
    # do that AI can't (judgment, taste, escalation, novel synthesis) rather
    # than the legacy function name.
    ("Leadership", "CEO / Founder",                    "C", 1.0, 1.0, 0.20, "Board",
     "Strategy synthesis, investor narrative drafting, competitive intel via AI research agents."),
    ("Leadership", "CTO / Chief AI Architect",         "C", 1.0, 1.0, 0.30, "CEO",
     "Architecture review with AI pair, code review automation, AI-assisted hiring loops."),
    ("Leadership", "Head of Autonomous Operations",    "C", 0.0, 1.0, 0.40, "CEO",
     "AI-run dashboards, automated reporting, exception-only escalation; designs the human-in-the-loop seams."),

    ("Product",    "Head of Product (AI-Native)",      "L", 0.5, 1.0, 0.30, "CEO",
     "AI user-research synthesis, PRD drafting copilots, AB-test analysis automation."),
    ("Product",    "Product Engineer",                 "S", 1.0, 2.0, 0.40, "Head of Product",
     "Ships, not just specs: AI-assisted prototypes, transcript coding, competitive scans."),
    ("Product",    "Product Designer (AI-Augmented)",  "S", 1.0, 2.0, 0.50, "Head of Product",
     "Generative design exploration, copy variants, accessibility audits via AI."),

    ("Engineering","Systems Architect",                "L", 1.0, 1.5, 0.25, "CTO",
     "AI architecture critic, automated design-doc review, owns the eval+observability spine."),
    ("Engineering","Forward-Deployed AI Engineer",     "S", 2.0, 4.0, 0.55, "Systems Architect",
     "Claude/Copilot pair-programming, AI test-gen, AI code review on PRs, deploys with customers."),
    ("Engineering","Applied AI Engineer",              "S", 1.0, 2.0, 0.40, "Systems Architect",
     "Owns the model layer; eval harness, dataset curation, prompt + retrieval tuning."),
    ("Engineering","Platform / AI-Ops Engineer",       "S", 0.5, 1.5, 0.60, "Systems Architect",
     "IaC generation, runbook automation, AI-driven incident triage, agent infra."),
    ("Engineering","Eval & Quality Engineer",          "S", 0.0, 1.0, 0.70, "Systems Architect",
     "Owns LLM evals + property tests; humans curate edge cases that agents miss."),

    ("Data",       "Decision Intelligence Lead",       "L", 0.0, 1.0, 0.50, "CTO",
     "Natural-language analytics, AI-built dashboards, anomaly detection, causal experimentation."),

    ("GTM",        "Head of Revenue (AI-Native)",      "L", 0.5, 1.0, 0.30, "CEO",
     "AI account research, call-summary intelligence, pipeline forecasting."),
    ("GTM",        "Forward-Deployed Account Lead",    "S", 1.0, 3.0, 0.40, "Head of Revenue",
     "AI prospecting, personalized outbound at scale, configures the customer's agent stack on-site."),
    ("GTM",        "Growth Engineer",                  "S", 1.0, 2.0, 0.60, "Head of Revenue",
     "AI content engine, SEO clusters, paid creative variants, lifecycle email automation."),
    ("GTM",        "Customer Outcomes Engineer",       "S", 1.0, 2.5, 0.55, "Head of Revenue",
     "AI-drafted QBRs, churn prediction, T1 deflection by AI agents, escalation on intent."),

    ("Ops",        "Finance Engineer",                 "S", 0.0, 1.0, 0.65, "Head of Autonomous Ops",
     "AI-built models, automated month-end, expense classification, scenario engines."),
    ("Ops",        "Talent & Team-Design Lead",        "S", 0.5, 1.0, 0.50, "Head of Autonomous Ops",
     "AI-screened pipelines, structured-interview kits, onboarding bots, org-shape continuous design."),
    ("Ops",        "Counsel (Fractional, AI-Assisted)","S", 0.2, 0.5, 0.55, "CEO",
     "Fractional; AI-drafted contracts and policy redlines, human review on novel terms."),
]


def generate(inputs: Dict) -> List[Dict]:
    y1_target = max(1, int(inputs.get("target_headcount_y1", 8)))
    y3_target = max(y1_target, int(inputs.get("target_headcount_y3", 30)))
    automation_target = float(inputs.get("ai_automation_target", 0.5))

    y1_w = sum(t[3] for t in TEMPLATE) or 1.0
    y3_w = sum(t[4] for t in TEMPLATE) or 1.0

    roles = []
    cum_y1 = 0.0
    cum_y3 = 0.0
    for fn, role, lvl, w1, w3, autom, reports, ai in TEMPLATE:
        hc1 = w1 / y1_w * y1_target
        hc3 = w3 / y3_w * y3_target
        cum_y1 += hc1
        cum_y3 += hc3
        # Nudge automation toward user's target by blending 30%
        eff_autom = autom * 0.7 + automation_target * 0.3
        roles.append({
            "function": fn, "role": role, "level": lvl,
            "headcount_y1": round(hc1, 1),
            "headcount_y3": round(hc3, 1),
            "automation": round(eff_autom, 2),
            "reports_to": reports,
            "ai_use": ai,
        })
    return roles
