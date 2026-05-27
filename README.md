# idea-evaluator

![demo](docs/demo.gif)

CLI + interactive xlsx for evaluating business ideas end-to-end:

- **Screen** a portfolio of ideas via Modified RICE, WSJF, CD3, EMV
- **Rubric-score** a single idea on 10 weighted dimensions (incl. AI-resilience)
- **Design an AI-native org structure** with role-by-role automation and cost projections
- **Value** the idea seven ways: DCF, Comparables, Berkus, Scorecard (Payne), VC method, First Chicago, Patent (cost / market / RFR / income)
- **Split founder equity** via a Slicing-Pie contribution-units model, net of ESOP
- Optionally call **Claude Code CLI** (`claude -p`) for qualitative VC-style commentary; falls back gracefully if absent

The generated xlsx is fully interactive — edit any yellow cell on the `Inputs` sheet and every downstream sheet recomputes via live Excel formulas.

## Install

```bash
source ~/projects/venv/bin/activate   # or your own venv
pip install -r requirements.txt
```

## Usage

```bash
# Interactive deep evaluation
python cli.py

# JSON-driven
python cli.py --input examples/meetingroi.json --output meetingroi.xlsx

# Portfolio screening across many ideas
python cli.py --screen examples/meeting_variants_screen.json \
              --screen-output portfolio.xlsx

# Dump a blank input template
python cli.py --dump-template my_idea.json

# Skip the Claude CLI call (rubric-only)
python cli.py --input my.json --no-ai

# Absorb an existing investor deck (.pptx) into JSON inputs
python cli.py --from-pptx deck.pptx --to-json deck_inputs.json
#   Deterministic regex pulls money/percent anchors (TAM, revenue, IRR,
#   margins, ask). If `claude` is on PATH, an LLM pass fills the rest
#   (problem, solution, ICP, rubric scores). Skip the LLM with --no-ai.

# Produce an investor pitch deck (.pptx) alongside the xlsx
python cli.py --input my.json --to-pptx pitch.pptx
#   10-slide deck: Cover · Opportunity · Solution · Traction · 5-Yr
#   Projections · Valuation (7 methods) · AI-Native Org · Equity ·
#   Investment Ask · Closing.

# Roundtrip: absorb a competitor deck and reproduce as your own
python cli.py --from-pptx their_deck.pptx --to-pptx my_deck.pptx
```

> .ppt (legacy binary, pre-2007) is not supported directly — convert first:
> `soffice --headless --convert-to pptx your_deck.ppt` (requires LibreOffice).

## Workbook sheets

`Summary` · `Inputs` · `Scorecard` · `Org` · `Equity` · `Valuation_DCF` · `Valuation_Comps` · `Valuation_EarlyStage` · `Valuation_Patent` · `Screening` · `Definitions`

## Method choice notes

- **Modified RICE** comes from the Samyama screening template. **WSJF (SAFe)**, **CD3 (Reinertsen)**, and **EMV** are run alongside it because the four methods diverge on time-sensitive vs. dollar-payoff bets — divergence in rank is itself a signal.
- The rubric weights **AI Resilience at 17%** (heaviest). If commodity LLMs eat your wedge in 24 months, the score reflects it.
- Every role title in the org template is intentionally AI-native (e.g. `Forward-Deployed AI Engineer`, `Eval & Quality Engineer`, `Decision Intelligence Lead`).

## License

MIT — see `LICENSE`.
