#!/usr/bin/env python3
"""Idea evaluator CLI.

Usage:
  python cli.py                              # fully interactive
  python cli.py --input idea.json            # JSON-driven
  python cli.py --input idea.json --output out.xlsx
  python cli.py --no-ai                      # skip Claude CLI call
  python cli.py --dump-template tmpl.json    # write a blank input template
"""
import argparse
import json
import sys
from pathlib import Path

from evaluator import schema, rubric, valuation, org, claude_client, screening, equity
from evaluator.xlsx_writer import write_workbook, write_screening_workbook
from evaluator import pptx_reader, pptx_writer


def prompt_one(field: schema.Field, current):
    suffix = ""
    if field.kind == schema.CHOICE:
        suffix = f" [{'/'.join(field.choices)}]"
    default = current if current is not None else field.default
    raw = input(f"  {field.prompt}{suffix} [{default}]: ").strip()
    return schema.coerce(field, raw if raw else default)


def interactive(seed: dict) -> dict:
    print("\n=== Idea Evaluator — Interactive Mode ===")
    print("Press Enter to accept the default in [brackets].\n")
    out = dict(seed)
    current_section = None
    for f in schema.ALL_FIELDS:
        if f.section != current_section:
            current_section = f.section
            print(f"\n--- {current_section.upper()} ---")
        out[f.key] = prompt_one(f, out.get(f.key))
    return out


def from_json(path: Path) -> dict:
    raw = json.loads(path.read_text())
    out = schema.defaults()
    for k, v in raw.items():
        if k in schema.FIELDS_BY_KEY:
            out[k] = schema.coerce(schema.FIELDS_BY_KEY[k], v)
        else:
            out[k] = v  # keep extras (e.g. notes)
    return out


def dump_template(path: Path) -> None:
    tmpl = {f.key: f.default for f in schema.ALL_FIELDS}
    path.write_text(json.dumps(tmpl, indent=2))
    print(f"Wrote template to {path}")


def print_report(inputs, rb_res, vals):
    print("\n" + "=" * 60)
    print(f"IDEA: {inputs.get('idea_name')}")
    print(f"  {inputs.get('one_liner','')}")
    print("=" * 60)
    print(f"\nComposite score: {rb_res['composite_5']:.2f}/5  "
          f"({rb_res['score_100']:.1f}/100)")
    print(f"Verdict: {rb_res['verdict']}\n")
    print("Category breakdown:")
    for c in rb_res["categories"]:
        print(f"  - {c['name']:18s} w={c['weight']:.0%}  avg={c['avg']:.2f}")
    print("\nTop strengths:")
    for k, v, cat in rb_res["strengths"]:
        print(f"  + {k} ({cat}): {v:.1f}")
    print("Top weaknesses:")
    for k, v, cat in rb_res["weaknesses"]:
        print(f"  - {k} ({cat}): {v:.1f}")

    print("\nValuations (USD):")
    print(f"  DCF equity value      : {vals['dcf']['equity']:>16,.0f}")
    print(f"  Comparables midpoint  : {vals['comparables']['midpoint']:>16,.0f}")
    print(f"  Berkus                : {vals['berkus']:>16,.0f}")
    print(f"  Scorecard (Payne)     : {vals['scorecard']:>16,.0f}")
    print(f"  VC method             : {vals['vc_method']:>16,.0f}")
    print(f"  First Chicago weighted: {vals['first_chicago']['weighted']:>16,.0f}")
    print(f"  Patent midpoint       : {vals['patent']['midpoint']:>16,.0f}")


def run_screen(in_path: Path, out_path: Path) -> None:
    data = json.loads(in_path.read_text())
    ideas = data["ideas"] if isinstance(data, dict) and "ideas" in data else data
    if not isinstance(ideas, list):
        print("--screen input must be a JSON list (or {ideas: [...]}).", file=sys.stderr)
        sys.exit(2)
    scored = screening.score_all(ideas)
    rows = screening.rank_table(scored)

    print(f"\n=== Portfolio Screening — {len(rows)} ideas ===")
    hdr = f"{'#':>2} {'Idea':<26} {'RICE':>10} {'WSJF':>8} {'CD3':>14} {'EMV':>14}  ranks(R/W/C/E)  spread"
    print(hdr)
    print("-" * len(hdr))
    for i, r in enumerate(rows, 1):
        print(f"{i:>2} {r['name'][:26]:<26} "
              f"{r['modified_rice']:>10.4f} {r['wsjf']:>8.1f} "
              f"{r['cd3']:>14,.0f} {r['emv']:>14,.0f}  "
              f"{r['rank_rice']}/{r['rank_wsjf']}/{r['rank_cd3']}/{r['rank_emv']}"
              f"   {r['rank_spread']}")
    high_spread = [r for r in rows if r["rank_spread"] >= 3]
    if high_spread:
        print("\nScoring-sensitive (rank spread ≥ 3 across methods) — deep-dive these:")
        for r in high_spread:
            print(f"  · {r['name']} (spread={r['rank_spread']})")
    write_screening_workbook(rows, str(out_path))
    print(f"\nScreening workbook written to {out_path}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", type=Path, help="JSON input file")
    ap.add_argument("--output", type=Path, default=Path("idea_evaluation.xlsx"),
                    help="Output xlsx path")
    ap.add_argument("--no-ai", action="store_true",
                    help="Skip Claude Code CLI qualitative analysis")
    ap.add_argument("--dump-template", type=Path,
                    help="Write a blank JSON input template and exit")
    ap.add_argument("--screen", type=Path,
                    help="Portfolio-screening mode: JSON list of ideas to rank")
    ap.add_argument("--screen-output", type=Path, default=Path("screening.xlsx"),
                    help="Output xlsx for --screen mode")
    ap.add_argument("--from-pptx", type=Path,
                    help="Absorb a .pptx pitch deck into a JSON inputs file "
                         "(use with --to-json, or pipe into --input).")
    ap.add_argument("--to-json", type=Path,
                    help="With --from-pptx, where to write the extracted JSON.")
    ap.add_argument("--to-pptx", type=Path,
                    help="Also produce an investor pitch deck (.pptx) "
                         "alongside the xlsx.")
    args = ap.parse_args()

    if args.dump_template:
        dump_template(args.dump_template)
        return

    if args.screen:
        run_screen(args.screen, args.screen_output)
        return

    if args.from_pptx:
        print(f"Absorbing deck: {args.from_pptx}")
        absorbed, diag = pptx_reader.absorb(
            str(args.from_pptx), use_llm=not args.no_ai)
        print(f"  slides={diag['slide_count']} "
              f"anchors={diag['anchors_found']} llm={diag['llm_used']}")
        out_json = args.to_json or args.from_pptx.with_suffix(".json")
        out_json.write_text(json.dumps(absorbed, indent=2, default=str))
        print(f"Wrote {out_json}")
        if not args.input:
            args.input = out_json

    if args.input:
        inputs = from_json(args.input)
        print(f"Loaded inputs from {args.input}")
    else:
        inputs = interactive(schema.defaults())

    rb_res = rubric.compute(inputs)
    vals = valuation.all_valuations(inputs, rb_res["composite_5"])

    qualitative = None
    if not args.no_ai:
        if claude_client.available():
            print("\nCalling Claude Code CLI for qualitative analysis...")
            qualitative = claude_client.analyze(inputs, rb_res)
            if qualitative:
                print("  ...done.")
            else:
                print("  ...skipped (no output / error).")
        else:
            print("\n(Claude Code CLI not on PATH — rubric-only mode.)")

    print_report(inputs, rb_res, vals)
    founders = equity.compute(inputs)
    equity.print_table(founders, float(inputs.get("esop_pool_pct", 0.15)))
    if qualitative:
        print("\n--- AI qualitative analysis ---\n")
        print(qualitative)

    org_plan = org.generate(inputs)
    write_workbook(inputs, {"rubric": rb_res, "valuations": vals,
                            "qualitative": qualitative},
                   str(args.output))
    print(f"\nWorkbook written to {args.output}")

    if args.to_pptx:
        pptx_writer.write_deck(
            inputs,
            {
                "rubric": rb_res,
                "valuations": vals,
                "org": org_plan,
                "equity": {
                    "founders": founders,
                    "esop_pct": float(inputs.get("esop_pool_pct", 0.15)),
                },
            },
            str(args.to_pptx),
        )
        print(f"Pitch deck written to {args.to_pptx}")


if __name__ == "__main__":
    main()
