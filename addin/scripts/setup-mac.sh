#!/usr/bin/env bash
# One-step developer setup on macOS: dependencies, Python packages, HTTPS certificate,
# build and sideload. Afterwards run `npm run serve` and restart Excel.
set -euo pipefail
cd "$(dirname "$0")/.."

[[ "$(uname -s)" == "Darwin" ]] || { echo "This script is for macOS."; exit 1; }
command -v node >/dev/null || { echo "Node.js 22 or later is required (brew install node)."; exit 1; }
major="$(node -p 'process.versions.node.split(".")[0]')"
(( major >= 22 )) || { echo "Node.js 22 or later is required; found $(node -v)."; exit 1; }
[[ -d "/Applications/Microsoft Excel.app" ]] || echo "warning: Microsoft Excel.app not found in /Applications"

echo "==> installing dependencies"
npm ci
echo "==> fetching numpy, pandas, matplotlib for Pyodide (checksums verified)"
npm run fetch:pyodide
echo "==> creating and trusting a localhost HTTPS certificate (macOS may ask for your password)"
npx office-addin-dev-certs install
echo "==> building"
npm run build
echo "==> sideloading the manifest into Excel"
bash scripts/sideload-mac.sh

cat <<'DONE'

Done. Next:
  1. npm run serve          (keep it running)
  2. Quit and reopen Excel, open a workbook
  3. Home > Add-ins > xlfly, or the "Python" button on the Home tab
DONE
