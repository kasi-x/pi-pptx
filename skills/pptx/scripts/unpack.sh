#!/usr/bin/env bash
# Unpack a .pptx/.potx (a zip of XML) into a directory for OOXML editing.
# usage: unpack.sh deck.pptx [outdir]
set -euo pipefail

deck=${1:?usage: unpack.sh deck.pptx [outdir]}
out=${2:-unpacked}

mkdir -p "$out"
python3 -m zipfile -e "$deck" "$out"
echo "unpacked: $deck -> $out"
