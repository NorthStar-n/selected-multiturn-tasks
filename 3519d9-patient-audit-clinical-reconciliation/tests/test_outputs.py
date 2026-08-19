"""Task Tests - LLM Judge Evaluation.

Environment Variables:
    LLM_JUDGE_API_KEY: API key for an OpenAI-compatible chat completions provider
    LLM_JUDGE_MODEL or LLM_JUDGE_MODELS: model id(s) to use
    LLM_JUDGE_API_URL or LLM_JUDGE_BASE_URL: optional provider endpoint override
    LLM_JUDGE_AGGREGATION: min, median, mean, or max for multi-model judging
    LLM_JUDGE_MODE: optional; use jobbench_static/static for JobBench static mode,
        or agent/agent_tools to force the legacy interactive judge.
"""

import hashlib
import json
from pathlib import Path

from test_utils import (
    PASSING_THRESHOLD,
    aggregate_judge_scores,
    build_agent_judge_prompt,
    call_agent_judge,
    call_jobbench_static_judge,
    criterion_name_to_field_name,
    create_dynamic_judge_response,
    ensure_output_files_rubric_criterion,
    extract_rubric_from_json,
    get_eval_dir,
    get_criterion_points,
    get_llm_judge_config,
    get_task_description_from_instruction,
    iter_llm_judge_model_configs,
    normalize_judge_scores,
    should_use_jobbench_static_judge,
)


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def verify_conversation_rubric_lock() -> None:
    """Verify the approved conversation-scoped rubric and source snapshot."""
    tests_dir = Path(__file__).parent
    lock_path = tests_dir / ".afo-conversation-rubric-lock.json"
    source_path = tests_dir / ".afo-source-rubrics.json"
    active_path = tests_dir / "rubrics.json"

    lock = json.loads(lock_path.read_text(encoding="utf-8"))
    source_rubric = json.loads(source_path.read_text(encoding="utf-8"))
    active_rubric = json.loads(active_path.read_text(encoding="utf-8"))

    assert lock["frozen_before_official_trials"] is True
    assert lock["official_trials_under_this_lock"] == 0
    amendment = lock["manual_history_amendment"]
    assert amendment["approved_by_team"] is True
    assert amendment["modified_atif_steps"] == [40, 41, 42]
    assert amendment["workspace_content_changed"] is False
    assert amendment["checkpoint_content_reused"] is True
    assert _sha256(source_path) == lock["source_task"]["rubric_sha256"]
    assert _sha256(active_path) == lock["active_rubric"]["sha256"]

    criteria = active_rubric.get("rubrics", [])
    scoring = [item for item in criteria if not item.get("score_excluded")]
    diagnostics = [item for item in criteria if item.get("score_excluded")]
    assert len(scoring) == lock["active_rubric"]["scoring_criteria"]
    assert sum(int(item["weight"]) for item in scoring) == lock["active_rubric"]["scoring_points"]
    assert active_rubric["total_points"] == lock["active_rubric"]["scoring_points"]
    assert {item["name"] for item in diagnostics} == set(
        lock["active_rubric"]["non_scoring_source_diagnostics"]
    )

    for item in scoring:
        citations = item.get("requirement_citations", [])
        assert citations, f"scored criterion lacks requirement citations: {item['name']}"
        assert all(
            citation.get("source", "").startswith(("conversation/", "workspace:"))
            for citation in citations
        ), f"scored criterion has a non-disclosed citation: {item['name']}"

    source_names = {item["name"] for item in source_rubric.get("rubrics", [])}
    assert {item["name"] for item in diagnostics} <= source_names


def discover_expected_files_from_solution() -> list[str]:
    """Derive expected deliverables from the hidden golden solution."""
    solution_dir = Path("/solution")
    if not solution_dir.exists():
        return []

    expected: list[str] = []
    for path in solution_dir.rglob("*"):
        if not path.is_file():
            continue
        rel = path.relative_to(solution_dir)
        if rel == Path("solve.sh"):
            continue
        if any(part.startswith(".") or part == "__pycache__" for part in rel.parts):
            continue
        expected.append(str(rel))
    return sorted(expected)


# Deterministic adapter: builders write solution/**, and the verifier
# derives the required output file list from that golden package.
EXPECTED_FILES = discover_expected_files_from_solution()
OUTPUT_FILES_RUBRIC_POINTS = 10

# Load rubric and task description at module level
verify_conversation_rubric_lock()
RUBRIC = ensure_output_files_rubric_criterion(
    extract_rubric_from_json(),
    EXPECTED_FILES,
    points=OUTPUT_FILES_RUBRIC_POINTS,
)
TASK_DESCRIPTION = get_task_description_from_instruction()
TaskJudgeResponse = create_dynamic_judge_response(RUBRIC)


class TestLLMJudge:
    """LLM judge evaluation using the task rubric."""

    def test_llm_judge_evaluation(self):
        """Agent judge evaluates agent output against the rubric."""
        config = get_llm_judge_config()
        assert config is not None, (
            "LLM judge not configured - LLM_JUDGE_API_KEY environment variable is required"
        )

        eval_dir = get_eval_dir()

        use_jobbench_static = should_use_jobbench_static_judge(RUBRIC)
        agent_prompt = None if use_jobbench_static else build_agent_judge_prompt(TASK_DESCRIPTION, RUBRIC)
        model_configs = iter_llm_judge_model_configs(config)
        print(
            f"\n[Judge Ensemble] Running {len(model_configs)} model(s) "
            f"with aggregation={config.get('aggregation', 'min')} "
            f"mode={'jobbench_static' if use_jobbench_static else 'agent_tools'}"
        )

        # Retry judge if response fails validation (field name mismatches)
        max_judge_retries = 3
        judge_scores = []
        for model_config in model_configs:
            model_name = model_config["model"]
            print(f"\n[Judge Model] {model_name}")
            model_validated = None
            last_error = None
            for attempt in range(max_judge_retries):
                if attempt > 0:
                    print(
                        f"\n[Judge Retry] {model_name} attempt {attempt + 1}/{max_judge_retries} "
                        f"after validation error: {last_error}"
                    )

                if use_jobbench_static:
                    scores = call_jobbench_static_judge(TASK_DESCRIPTION, RUBRIC, model_config, eval_dir)
                else:
                    scores = call_agent_judge(agent_prompt, model_config, eval_dir)
                scores = normalize_judge_scores(scores, RUBRIC, PASSING_THRESHOLD)

                try:
                    model_validated = TaskJudgeResponse(**scores)
                    break
                except Exception as e:
                    last_error = str(e)[:200]
                    print(f"[Judge Validation Error] {model_name}: {last_error}")
                    continue

            assert model_validated is not None, (
                f"LLM judge failed validation for {model_name} after {max_judge_retries} attempts. "
                f"Last error: {last_error}"
            )
            model_scores = model_validated.dict()
            model_scores["_judge_model"] = model_name
            judge_scores.append(model_scores)
            print(f"JUDGE_MODEL {model_name}: {model_validated.total}/{model_validated.max_total}")

        aggregated_scores = aggregate_judge_scores(
            judge_scores,
            RUBRIC,
            PASSING_THRESHOLD,
            strategy=config.get("aggregation", "min"),
        )
        validated = TaskJudgeResponse(**aggregated_scores)

        print(f"\n{'='*60}")
        print("LLM JUDGE EVALUATION")
        print("=" * 60)
        if len(judge_scores) > 1:
            print(f"JUDGE_AGGREGATION: {config.get('aggregation', 'min')}")

        for criterion in RUBRIC.get("criteria", []):
            name = criterion["name"]
            points = get_criterion_points(criterion)
            field_name = criterion_name_to_field_name(name)
            score = getattr(validated, field_name, 0)
            print(f"{name}: {score}/{points}")
            evidence = validated.evidence.get(field_name, {})
            if evidence:
                evidence_parts = [
                    evidence.get("file", ""),
                    evidence.get("location", ""),
                    f"expected={evidence.get('expected', '')}" if evidence.get("expected") else "",
                    evidence.get("observed", ""),
                    evidence.get("reason", ""),
                ]
                evidence_text = " | ".join(part for part in evidence_parts if part)
                print(f"  evidence: {evidence_text or '(none provided)'}")

        # The "TOTAL: <int>/<int>" line below is parsed by test.sh to compute
        # a FRACTIONAL reward (judge_total / judge_max_total) rather than a
        # binary pass/fail at a fixed threshold. This preserves graduated
        # scoring from the rubric levels (Complete / Strong / Partial /
        # Minimal / None) all the way through to the verifier's reward.txt.
        print(f"TOTAL: {validated.total}/{validated.max_total}")
        judge_score = validated.total / validated.max_total if validated.max_total else 0.0
        print(f"LLM_JUDGE_SCORE: {judge_score:.4f}")
        print(f"\nFeedback: {validated.feedback}")
        print("=" * 60)

        # No hard assertion on the judge total — the score itself IS the
        # signal. test.sh reads the TOTAL line above and writes the
        # fractional reward; the difficulty gate downstream is fed the
        # graduated judge_score percentage by the pipeline.
        #
        # We still require that the judge ran and produced a non-degenerate
        # result so an empty / crashed judge doesn't silently slip through.
        assert validated.max_total > 0, (
            f"LLM judge returned zero max_total — rubric or judge call is broken. "
            f"Feedback: {validated.feedback}"
        )
