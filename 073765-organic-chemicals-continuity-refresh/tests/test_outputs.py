"""LLM judge verifier for long-horizon tasks."""

import sys
from pathlib import Path

from rewardkit import criterion

sys.path.insert(0, str(Path(__file__).parent))

from test_utils import (
    build_agent_judge_prompt,
    call_agent_judge,
    criterion_name_to_field_name,
    get_eval_dir,
    get_llm_judge_config,
    get_task_description,
    load_rubric,
    normalize_judge_scores,
)
from checkpoint_noop import is_exact_checkpoint_output


RUBRIC = load_rubric()
TASK_DESCRIPTION = get_task_description()


@criterion(description="LLM judge score for submitted /app/output deliverables")
def agent_judge_score(workspace: Path) -> float:
    eval_dir = get_eval_dir()
    if is_exact_checkpoint_output(eval_dir):
        print("Exact unchanged checkpoint output detected; returning no-op reward 0.0")
        return 0.0

    config = get_llm_judge_config()
    if config is None:
        raise RuntimeError("LLM_JUDGE_API_KEY is required")

    prompt = build_agent_judge_prompt(TASK_DESCRIPTION, RUBRIC)
    raw_scores = call_agent_judge(prompt, config, eval_dir)
    scores = normalize_judge_scores(raw_scores, RUBRIC)

    print(f"\n{'=' * 60}")
    print("LLM JUDGE EVALUATION")
    print("=" * 60)

    for rubric_criterion in RUBRIC["criteria"]:
        name = rubric_criterion["name"]
        field_name = criterion_name_to_field_name(name)
        points = int(rubric_criterion.get("points", rubric_criterion.get("score", 0)) or 0)
        value = scores.get(field_name, 0)
        print(f"{name}: {value}/{points}")

        evidence = (scores.get("evidence") or {}).get(field_name, {})
        if evidence:
            parts = [
                evidence.get("file", ""),
                evidence.get("location", ""),
                f"expected={evidence.get('expected', '')}" if evidence.get("expected") else "",
                evidence.get("observed", ""),
                evidence.get("reason", ""),
            ]
            text = " | ".join(part for part in parts if part)
            print(f"  evidence: {text or '(none provided)'}")

    total = int(scores.get("total", 0) or 0)
    max_total = int(scores.get("max_total", 0) or 0)
    if max_total <= 0:
        raise RuntimeError("LLM judge returned zero max_total")

    reward = total / max_total
    print(f"TOTAL: {total}/{max_total}")
    print(f"\nFeedback: {scores.get('feedback', '')}")
    print("=" * 60)
    return reward
