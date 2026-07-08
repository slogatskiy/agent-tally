#!/bin/bash
# Monthly orchestrator run by the LaunchDaemon on the Mac Mini.
#   scrape Compass -> update xlsx -> rebuild HTML/JSON -> commit & push.
# GitHub Pages then serves the refreshed widget. Idempotent: if nothing
# changed (same counts), it commits nothing.
set -euo pipefail
cd "$(dirname "$0")"

# LaunchDaemon PATH is stripped bare; add what we need (venv python, git, gh).
export PATH="$(pwd)/.venv/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin"
export PYTHON="${PYTHON:-python3}"

mkdir -p logs
echo "=== run $(date '+%Y-%m-%d %H:%M:%S %z') ==="

"$PYTHON" tally.py                 # updates COMP_Agent)Tally.xlsx (+ dated backup)
"$PYTHON" build_site.py            # regenerates docs/index.html + data/history.json

git add -A
if git diff --cached --quiet; then
  echo "no changes — nothing to publish"
  exit 0
fi

git commit -m "data: monthly agent tally $(date '+%Y-%m-%d')"
git push origin main
echo "published to GitHub Pages"
