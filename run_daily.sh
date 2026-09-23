#!/bin/bash
# Daily RFC sync — called by launchd
set -euo pipefail
ROOT="$(cd "$(dirname "$0")" && pwd)"
cd "$ROOT"
export PATH="/usr/bin:/bin:/usr/sbin:/sbin:/usr/local/bin:/opt/homebrew/bin:$HOME/Library/Python/3.9/bin:$PATH"
LOG_DIR="$ROOT/state"
mkdir -p "$LOG_DIR"
{
  echo "==== $(date '+%Y-%m-%d %H:%M:%S') ===="
  /usr/bin/python3 "$ROOT/pdf_to_excel.py" --sync
  echo "OK"
} >>"$LOG_DIR/daily.log" 2>&1
