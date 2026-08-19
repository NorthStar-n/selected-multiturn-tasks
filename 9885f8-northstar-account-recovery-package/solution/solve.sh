#!/usr/bin/env bash
set -euo pipefail

rm -rf /app/output
mkdir -p /app/output

if [ -d /solution ]; then
  cp -a /solution/. /app/output/
  rm -f /app/output/solve.sh
fi
