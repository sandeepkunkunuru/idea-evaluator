"""Best-effort qualitative analysis via Claude Code CLI.

If the `claude` binary isn't on PATH (or times out / errors), we return None
and the rest of the pipeline still produces a complete report.
"""
import json
import shutil
import subprocess
from typing import Optional, Dict


PROMPT_TEMPLATE = """You are an experienced VC/operator. Given the structured
inputs and rubric results below, write a tight qualitative assessment in <= 350
words covering: (1) what's most compelling, (2) the 2 biggest risks, (3) what
must be true for this to be a $100M+ outcome, (4) AI-disruption read (will
commodity LLMs help or kill this in 24 months?), (5) one concrete next experiment
to run. No fluff, no headers, plain prose paragraphs.

INPUTS:
{inputs}

RUBRIC RESULT:
- Score (0-100): {score}
- Verdict: {verdict}
- Top strengths: {strengths}
- Top weaknesses: {weaknesses}
"""


def available() -> bool:
    return shutil.which("claude") is not None


def analyze(inputs: Dict, rubric_result: Dict, timeout: int = 120) -> Optional[str]:
    if not available():
        return None
    prompt = PROMPT_TEMPLATE.format(
        inputs=json.dumps({k: v for k, v in inputs.items()
                           if not k.startswith("score_")}, indent=2, default=str),
        score=f"{rubric_result['score_100']:.1f}",
        verdict=rubric_result["verdict"],
        strengths=", ".join(f"{k}={v:.1f}" for k, v, _ in rubric_result["strengths"]),
        weaknesses=", ".join(f"{k}={v:.1f}" for k, v, _ in rubric_result["weaknesses"]),
    )
    try:
        proc = subprocess.run(
            ["claude", "-p", prompt],
            capture_output=True, text=True, timeout=timeout,
        )
        if proc.returncode != 0:
            return None
        out = proc.stdout.strip()
        return out or None
    except (subprocess.TimeoutExpired, FileNotFoundError, OSError):
        return None
