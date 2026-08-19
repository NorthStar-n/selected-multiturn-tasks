#!/usr/bin/env bash
set -euo pipefail

rm -rf /app/output
mkdir -p /app/output

SOURCE_DIR="$(cd "$(dirname "$0")" && pwd)"
find "$SOURCE_DIR" -maxdepth 1 -type f ! -name solve.sh -exec cp -a {} /app/output/ \;
