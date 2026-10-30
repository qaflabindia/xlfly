#!/usr/bin/env bash
# Sideload the development manifest into Excel for Mac (Microsoft's documented method:
# place the manifest in Excel's "wef" folder, then restart Excel and use Home > Add-ins).
set -euo pipefail
wef="$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef"
mkdir -p "$wef"
cp "$(dirname "$0")/../dist/manifest.xml" "$wef/xlfly.manifest.xml"
echo "Installed $wef/xlfly.manifest.xml — restart Excel, then Home > Add-ins > xlfly."
