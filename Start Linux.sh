#!/usr/bin/env bash
set -e
cd "$(dirname "$0")"

if ! command -v node >/dev/null 2>&1; then
  echo "Node.js 20 or newer is required to run Rubric Builder."
  echo "Install the LTS version from https://nodejs.org/, then run this starter again."
  echo
  read -r -p "Press Enter to close."
  exit 1
fi

node scripts/start-local.mjs
echo
read -r -p "Rubric Builder stopped. Press Enter to close."
