#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")/.."

VERSION=$(python3 - <<'PY'
import json,sys
print(json.load(open('manifest.json'))['version'])
PY
)

mkdir -p dist
OUT="dist/htmltabletoexcel-${VERSION}.zip"
rm -f "$OUT"
zip -r "$OUT" . \
  -x "dist/*" \
  -x ".git/*" \
  -x "*.DS_Store" \
  -x "__MACOSX/*" \
  -x "*/.*" \
  -x "tools/*"

echo "Packed -> $OUT"
