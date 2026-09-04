#!/bin/zsh
cd "$(dirname "$0")"

NODE_BIN=""
for candidate in node /opt/homebrew/bin/node /usr/local/bin/node; do
  if command -v "$candidate" >/dev/null 2>&1; then
    NODE_BIN="$(command -v "$candidate")"
    break
  fi
done

if [ -z "$NODE_BIN" ]; then
  echo "Node.js 20 or newer is required to run Rubric Builder."
  echo "Install the LTS version from https://nodejs.org/, then run this starter again."
  echo
  read -k 1 "?Press any key to close."
  exit 1
fi

"$NODE_BIN" scripts/start-local.mjs
echo
echo "Rubric Builder stopped. You can close this window."
read -k 1 "?Press any key to close."
