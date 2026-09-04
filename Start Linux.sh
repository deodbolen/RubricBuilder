#!/usr/bin/env bash
set -e
cd "$(dirname "$0")"
node scripts/start-local.mjs
echo
read -r -p "Rubric Builder stopped. Press Enter to close."
