#!/usr/bin/env bash
# ダブルクリックで Casting List Helper を起動してブラウザで開く。
# 初回は Python の環境を自動でセットアップする。
# 終了するときはこのターミナルウィンドウを閉じる。

cd "$(dirname "$0")" || exit 1
URL="http://localhost:8501"

# すでに起動中ならブラウザで開くだけ
if curl -s -o /dev/null "$URL"; then
  open "$URL"
  exit 0
fi

# 初回セットアップ
if [[ ! -x .venv/bin/streamlit ]]; then
  echo "==> 初回セットアップ中（数分かかります）..."
  python3 -m venv .venv || { echo "python3 が見つかりません"; read -r; exit 1; }
  .venv/bin/pip install -q --upgrade pip
  .venv/bin/pip install -q -r requirements.txt || { echo "インストールに失敗しました"; read -r; exit 1; }
fi

# 起動できたらブラウザを開く
(
  for _ in $(seq 1 30); do
    sleep 1
    if curl -s -o /dev/null "$URL"; then open "$URL"; break; fi
  done
) &

echo "==> Casting List Helper を起動しました: $URL"
echo "    終了するときはこのウィンドウを閉じてください。"
exec .venv/bin/streamlit run app.py --server.headless true --browser.gatherUsageStats false
