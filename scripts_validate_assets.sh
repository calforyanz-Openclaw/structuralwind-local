#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")"

TMPDIR="$(mktemp -d)"
trap 'rm -rf "$TMPDIR"' EXIT
mkdir -p "$TMPDIR/lib"

curl -L -s https://structuralwind.com -o "$TMPDIR/index.html"
curl -L -s https://structuralwind.com/about.html -o "$TMPDIR/about.html"
curl -L -s https://structuralwind.com/style.css -o "$TMPDIR/style.css"
curl -L -s 'https://structuralwind.com/script.js?v=ifc-archive-path' -o "$TMPDIR/script.js"
curl -L -s https://structuralwind.com/mesh-classifier.js -o "$TMPDIR/mesh-classifier.js"
curl -L -s https://structuralwind.com/lib/terrain-samples.js -o "$TMPDIR/lib/terrain-samples.js"
curl -L -s https://structuralwind.com/lib/elev-refine.js -o "$TMPDIR/lib/elev-refine.js"
curl -L -s https://structuralwind.com/lib/mh-topography.js -o "$TMPDIR/lib/mh-topography.js"

compare() {
  local live="$1"
  local localf="$2"
  if diff -q "$live" "$localf" >/dev/null; then
    printf 'OK  %s\n' "$localf"
  else
    printf 'DIFF %s\n' "$localf"
  fi
}

compare "$TMPDIR/index.html" public/index.html
compare "$TMPDIR/about.html" public/about.html
compare "$TMPDIR/style.css" public/style.css
compare "$TMPDIR/script.js" public/script.js
compare "$TMPDIR/mesh-classifier.js" public/mesh-classifier.js
compare "$TMPDIR/lib/terrain-samples.js" public/lib/terrain-samples.js
compare "$TMPDIR/lib/elev-refine.js" public/lib/elev-refine.js
compare "$TMPDIR/lib/mh-topography.js" public/lib/mh-topography.js
