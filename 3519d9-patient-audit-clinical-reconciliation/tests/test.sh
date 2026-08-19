#!/bin/bash
# Verifier entry point. Runs pytest and writes a FRACTIONAL reward to
# /logs/verifier/reward.txt — the LLM judge contributes partial credit
# based on its TOTAL: <points>/<max> output, rather than a binary
# pass/fail at a fixed threshold.
#
# Reward semantics:
#   - All deterministic tests must pass (any pytest failure → reward 0).
#   - When the LLM judge ran, reward = judge_total / judge_max_total
#     (fractional 0.0–1.0, preserving graduated scoring).
#   - When the LLM judge did NOT run (no API key / skipped), reward = 1
#     (all-deterministic-pass case).

pytest /tests/test_outputs.py -s -v -rA --tb=short 2>&1 | tee /logs/verifier/pytest_output.txt
PYTEST_EXIT=${PIPESTATUS[0]}

if [ "$PYTEST_EXIT" -ne 0 ]; then
  # Any pytest test failure → reward 0. Deterministic failures are
  # structural problems the agent must fix; the judge can't redeem them.
  echo 0 > /logs/verifier/reward.txt
  exit 0
fi

# All pytest tests passed. Compute graduated reward from the LLM judge
# total (printed by TestLLMJudge.test_llm_judge_evaluation as
# "TOTAL: <int>/<int>"). If no such line is present (judge skipped),
# default to reward 1.
JUDGE_LINE=$(grep -oE 'TOTAL: [0-9]+/[0-9]+' /logs/verifier/pytest_output.txt | tail -1)
if [ -n "$JUDGE_LINE" ]; then
  TOTAL=$(echo "$JUDGE_LINE" | sed -E 's/TOTAL: ([0-9]+)\/([0-9]+)/\1/')
  MAX=$(echo "$JUDGE_LINE" | sed -E 's/TOTAL: ([0-9]+)\/([0-9]+)/\2/')
  if [ "$MAX" -gt 0 ]; then
    REWARD=$(awk "BEGIN {printf \"%.4f\", $TOTAL / $MAX}")
  else
    REWARD="1.0"
  fi
  echo "$REWARD" > /logs/verifier/reward.txt
else
  echo 1 > /logs/verifier/reward.txt
fi
