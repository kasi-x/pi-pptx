#!/usr/bin/env bash
# Pack an unpacked directory back into a .pptx.
# Zips from INSIDE the directory and removes the target first, so deleted
# parts do not survive in the output.
# usage: pack.sh unpacked_dir out.pptx
set -euo pipefail

dir=${1:?usage: pack.sh unpacked_dir out.pptx}
out=${2:?usage: pack.sh unpacked_dir out.pptx}

command -v zip >/dev/null || { echo "pack: zip(1) not found" >&2; exit 1; }
out=$(realpath -m "$out")

rm -f "$out"
(cd "$dir" && zip -Xr "$out" . >/dev/null)
echo "packed: $dir -> $out"
