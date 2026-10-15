#!/usr/bin/env bash
# Render every slide of a deck to JPEG for visual QA.
# LibreOffice converts the deck to PDF, poppler rasterizes the pages.
# Prints the absolute paths of the created images (pass them to the image
# viewer / read tool).
#
# usage: render.sh deck.pptx [outdir] [dpi]
set -euo pipefail

deck=${1:?usage: render.sh deck.pptx [outdir] [dpi]}
out=${2:-render}
dpi=${3:-110}

abs_deck=$(realpath "$deck")
out=$(realpath -m "$out")
mkdir -p "$out"

command -v soffice >/dev/null || { echo "render: soffice (LibreOffice) not found" >&2; exit 1; }
command -v pdftoppm >/dev/null || { echo "render: pdftoppm (poppler-utils) not found" >&2; exit 1; }

# Isolated profile: avoids clashing with a running LibreOffice instance.
prof=${TMPDIR:-/tmp}/pi-pptx-soffice-profile
soffice --headless -env:UserInstallation="file://$prof" \
	--convert-to pdf --outdir "$out" "$abs_deck" >/dev/null

pdf="$out/$(basename "${abs_deck%.*}").pdf"
[ -f "$pdf" ] || { echo "render: PDF was not produced ($pdf)" >&2; exit 1; }

rm -f "$out"/slide-*.jpg
pdftoppm -jpeg -r "$dpi" "$pdf" "$out/slide"
ls -1 "$out"/slide-*.jpg
