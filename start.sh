#!/usr/bin/env bash
# Usage: ./start.sh
# Streamlit Web アプリ (Casting List Helper) を起動してブラウザで開く

set -euo pipefail
cd "$(dirname "$0")"
source .venv/bin/activate

exec streamlit run app.py "$@"
