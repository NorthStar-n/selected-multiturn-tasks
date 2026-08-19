#!/bin/bash
set -euo pipefail

mkdir -p /logs/verifier
export PYTHONUNBUFFERED=1

rewardkit /tests --workspace /app --output /logs/verifier/reward.json
cat /logs/verifier/reward.json
