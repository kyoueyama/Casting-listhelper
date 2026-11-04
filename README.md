# Casting Listhelper

Astream のキャンペーン提案リスト Excel を自動取得し、業務用フォーマットに整形するツール。

## セットアップ（初回のみ）

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
playwright install chromium

cp .env.example .env
# .env を編集して ASTREAM_EMAIL / ASTREAM_PASSWORD を入れる
```

## 使い方

```bash
# 一気通貫（DL → 加工）
./run.sh "キャンペーン名"
./run.sh "キャンペーン名" --open      # 加工後ファイルをそのまま開く
```

個別にも実行できる:

```bash
source .venv/bin/activate
python download_astream.py --campaign "キャンペーン名"          # DL のみ
python download_astream.py --campaign "キャンペーン名" --headed # ブラウザ表示
python process_excel.py downloads/<file>.xlsx                  # 加工のみ
python process_excel.py downloads/<file>.xlsx --inspect        # 列構成だけ確認
```

## 加工の仕様（毎回適用）

1. A 列の左に空列を 2 列挿入（→ 元 A 列が C 列になる）
2. 1 行目の上に空行を 1 行追加（→ 元ヘッダーが 2 行目になる）
3. B2 に `No` を記入
4. 2 行目（ヘッダー行）に折り返し表示を適用
5. フォト列の画像を新しいフォト列に再配置
6. 列順を `名前 / フォト / アカウント名 / ...` に整える
7. `国` 列を削除
8. 下記 20 列だけ残してそれ以外を削除

残す列（この順序）:

```
名前, フォト, アカウント名, アカウントURL, フォロワー数,
いいねエンゲージメント, エンゲージメント,
フォロワーの男性割合, フォロワーの女性割合,
フォロワーの13-17歳の男性の割合, フォロワーの18-24歳の男性の割合,
フォロワーの25-34歳の男性の割合, フォロワーの35-44歳の男性の割合,
フォロワーの45-64歳の男性の割合,
フォロワーの13-17歳の女性の割合, フォロワーの18-24歳の女性の割合,
フォロワーの25-34歳の女性の割合, フォロワーの35-44歳の女性の割合,
フォロワーの45-64歳の女性の割合,
平均リール再生数
```

## 出力

- `downloads/<キャンペーン名>_YYYYMMDD_HHMMSS.xlsx` – Astream から落とした生データ
- `processed/<同名>_processed.xlsx` – 整形後

## 注意

- `.env` は `.gitignore` 済み。絶対にコミットしない
- 初回ログイン後は `.auth/state.json` にセッションが保存され、再ログインを回避する
- Astream の UI が変わるとセレクタが外れるので、その時はエラーメッセージと `debug/error_*.png` を見て修正
