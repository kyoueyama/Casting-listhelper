#!/usr/bin/env bash
# Usage: ./run.sh "キャンペーン名" [--open]
# --open を付けると加工後の xlsx を Excel で開く

set -euo pipefail
cd "$(dirname "$0")"
source .venv/bin/activate

if [[ $# -lt 1 ]]; then
  echo "Usage: $0 \"キャンペーン名\" [--open]" >&2
  exit 2
fi

CAMPAIGN="$1"
OPEN=false
if [[ "${2-}" == "--open" ]]; then
  OPEN=true
fi

echo "==> downloading: $CAMPAIGN"
python download_astream.py --campaign "$CAMPAIGN"

LATEST=$(ls -t downloads/*.xlsx | head -n1)
echo "==> processing: $LATEST"
python process_excel.py "$LATEST"

PROCESSED=$(ls -t processed/*.xlsx | head -n1)
echo "==> done: $PROCESSED"

if $OPEN; then
  open "$PROCESSED"
fi
