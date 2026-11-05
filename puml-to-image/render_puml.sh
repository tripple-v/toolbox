#!/usr/bin/env bash
# Renders PlantUML files to images, locally, with the plantuml command line.
# Usage: render_puml.sh [-f png|svg|both] [-o OUTPUT_DIR] FILE.puml [FILE2.puml ...]
# Defaults: both formats, images written next to each source file.
set -euo pipefail

usage() {
  sed -n '2,4p' "$0" | sed 's/^# //'
  exit "${1:-1}"
}

format="both"
outdir=""
while getopts ":f:o:h" opt; do
  case "$opt" in
    f) format="$OPTARG" ;;
    o) outdir="$OPTARG" ;;
    h) usage 0 ;;
    *) usage ;;
  esac
done
shift $((OPTIND - 1))
[ "$#" -ge 1 ] || usage

case "$format" in
  png|svg) formats=("$format") ;;
  both) formats=(png svg) ;;
  *) echo "Unknown format: $format (png, svg or both)" >&2; exit 1 ;;
esac

if ! command -v plantuml >/dev/null 2>&1; then
  echo "plantuml is not installed. On macOS: brew install plantuml. On Windows: choco install plantuml. Elsewhere: https://plantuml.com/download" >&2
  exit 1
fi

for src in "$@"; do
  [ -f "$src" ] || { echo "Not a file: $src" >&2; exit 1; }
  dest="${outdir:-$(dirname "$src")}"
  mkdir -p "$dest"
  for fmt in "${formats[@]}"; do
    plantuml "-t$fmt" -o "$(cd "$dest" && pwd)" "$src"
  done
  base="$(basename "${src%.*}")"
  for fmt in "${formats[@]}"; do
    echo "$dest/$base.$fmt"
  done
done
