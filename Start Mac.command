#!/bin/zsh
cd "$(dirname "$0")"
node scripts/start-local.mjs
echo
echo "Rubric Builder stopped. You can close this window."
read -k 1 "?Press any key to close."
