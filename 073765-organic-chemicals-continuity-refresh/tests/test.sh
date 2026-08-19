#!/bin/bash
set -euo pipefail

mkdir -p /logs/verifier

rewardkit /tests --workspace /app --output /logs/verifier/reward.json
cat /logs/verifier/reward.json
