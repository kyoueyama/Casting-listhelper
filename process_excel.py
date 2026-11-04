"""
Astream の提案リスト xlsx を業務用フォーマットに整形する。

毎回行う処理（仕様）:
  ① A列の左に空列を 2 列挿入（→ 元A列がC列になる）
  ② 1行目の上に空行を 1 行追加（→ 元ヘッダー行が 2 行目になる）
  ③ B2 に "No" を書き込む
  ④ 2行目（ヘッダー行）に折り返して表示を適用
  ⑤ 元「フォト」列に埋め込まれた画像を、新しい「フォト」列の各セルに再配置
  ⑥ 列順を: 名前, フォト, アカウント名, ... に整える
  ⑦ 「国」列を削除
  ⑧ KEEP_COLUMNS だけ残し、それ以外を削除

  python process_excel.py downloads/<file>.xlsx
  python process_excel.py downloads/<file>.xlsx -o out.xlsx
  python process_excel.py downloads/<file>.xlsx --inspect    # 列構成だけ見る
"""

from __future__ import annotations

import argparse
import io
import sys
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

DEFAULT_FONT_NAME = "メイリオ"

ROOT = Path(__file__).resolve().parent
PROCESSED_DIR = ROOT / "processed"

# 残す列（この順番で出力）
KEEP_COLUMNS = [
    "名前",
    "フォト",
    "アカウント名",
    "アカウントURL",
    "フォロワー数",
    "いいねエンゲージメント",
    "エンゲージメント",
    "フォロワーの男性割合",
    "フォロワーの女性割合",
    "フォロワーの13-17歳の男性の割合",
    "フォロワーの18-24歳の男性の割合",
    "フォロワーの25-34歳の男性の割合",
    "フォロワーの35-44歳の男性の割合",
    "フォロワーの45-64歳の男性の割合",
    "フォロワーの13-17歳の女性の割合",
    "フォロワーの18-24歳の女性の割合",
    "フォロワーの25-34歳の女性の割合",
    "フォロワーの35-44歳の女性の割合",
    "フォロワーの45-64歳の女性の割合",
    "平均リール再生数",
]

LEFT_PAD_COLS = 2  # 左に挿入する空列数
TOP_PAD_ROWS = 1   # 上に挿入する空行数
SRC_HEADER_ROW = 1  # 元データのヘッダー行


def inspect(input_path: Path) -> None:
    df = pd.read_excel(input_path)
    print(f"行数: {len(df)}")
    print(f"列数: {len(df.columns)}")
    print("--- columns ---")
    for i, c in enumerate(df.columns):
        print(f"  [{i:>2}] {c!r}  (dtype={df[c].dtype})")


def collect_images_by_row(src_ws, photo_col: int) -> dict[int, object]:
    """フォト列に紐づく画像を {元行番号: Image} で返す"""
    result: dict[int, object] = {}
    for img in list(src_ws._images):
        try:
            anchor_from = img.anchor._from
            col_1based = anchor_from.col + 1
            row_1based = anchor_from.row + 1
            if col_1based == photo_col:
                result[row_1based] = img
        except Exception:
            continue
    return result


def apply_transform_to_sheet(src_ws, dst_ws, keep_columns: list[str] | None = None) -> list[str]:
    """src_ws の内容を整形して dst_ws に書き込む。dst_ws は空のシート想定。

    keep_columns を指定すると、その順序で出力する。未指定なら KEEP_COLUMNS。
    戻り値: 元シートに存在しなかった列名のリスト（呼び出し側で警告表示に使える）
    """
    if keep_columns is None:
        keep_columns = KEEP_COLUMNS

    # 元シートのヘッダー → 列番号(1-indexed) のマップ
    headers = {}
    for cell in src_ws[SRC_HEADER_ROW]:
        if cell.value is not None:
            headers[str(cell.value).strip()] = cell.column

    keep_pairs: list[tuple[str, int]] = []
    missing: list[str] = []
    for name in keep_columns:
        if name in headers:
            keep_pairs.append((name, headers[name]))
        else:
            missing.append(name)
    if missing:
        print(f"WARN: 残す列に該当しないものがありました: {missing}", file=sys.stderr)

    photo_col_src = headers.get("フォト")
    images_by_src_row = (
        collect_images_by_row(src_ws, photo_col_src) if photo_col_src else {}
    )

    new_header_row = SRC_HEADER_ROW + TOP_PAD_ROWS  # = 2

    # ③ B2 = "No"
    dst_ws.cell(row=new_header_row, column=2, value="No")

    # ヘッダー書き込み（C2 から）
    new_photo_col = None
    for new_idx, (name, src_col) in enumerate(keep_pairs, start=LEFT_PAD_COLS + 1):
        dst_ws.cell(row=new_header_row, column=new_idx, value=name)
        if name == "フォト":
            new_photo_col = new_idx

    # データ行を写す（1セルずつ）
    last_src_row = src_ws.max_row
    for src_row in range(SRC_HEADER_ROW + 1, last_src_row + 1):
        new_row = src_row + TOP_PAD_ROWS
        # B 列に連番を入れる（1, 2, 3, ...）
        no = src_row - SRC_HEADER_ROW
        dst_ws.cell(row=new_row, column=2, value=no)
        for new_idx, (_, src_col) in enumerate(keep_pairs, start=LEFT_PAD_COLS + 1):
            src_cell = src_ws.cell(row=src_row, column=src_col)
            dst_cell = dst_ws.cell(row=new_row, column=new_idx, value=src_cell.value)
            if src_cell.number_format and src_cell.number_format != "General":
                dst_cell.number_format = src_cell.number_format

    # 行2: ラップ + 中央揃え / その他: 中央揃えのみ
    total_cols = LEFT_PAD_COLS + len(keep_pairs)
    last_new_row = last_src_row + TOP_PAD_ROWS
    header_align = Alignment(wrap_text=True, vertical="center", horizontal="center")
    center_align = Alignment(vertical="center", horizontal="center")
    default_font = Font(name=DEFAULT_FONT_NAME)
    for row in range(1, last_new_row + 1):
        for col in range(1, total_cols + 1):
            cell = dst_ws.cell(row=row, column=col)
            cell.font = default_font
            cell.alignment = header_align if row == new_header_row else center_align

    # ⑤ 画像を新フォト列に再配置
    if new_photo_col and images_by_src_row:
        new_photo_letter = get_column_letter(new_photo_col)
        for src_row, src_img in images_by_src_row.items():
            new_row = src_row + TOP_PAD_ROWS
            try:
                img_bytes = src_img._data()
                bio = io.BytesIO(img_bytes)
                new_img = XLImage(bio)
                # 元のサイズを保持（指定されていれば）
                if getattr(src_img, "width", None):
                    new_img.width = src_img.width
                if getattr(src_img, "height", None):
                    new_img.height = src_img.height
                dst_ws.add_image(new_img, f"{new_photo_letter}{new_row}")
            except Exception as e:
                print(
                    f"WARN: 画像の再配置に失敗 (row={src_row}): {e}",
                    file=sys.stderr,
                )

    # 行高・列幅を出来る限り引き継ぐ（画像の見栄え用）
    # 行: 元の各データ行高 → 新しい行 (+TOP_PAD_ROWS)
    for src_row in range(SRC_HEADER_ROW, last_src_row + 1):
        h = src_ws.row_dimensions[src_row].height
        if h:
            dst_ws.row_dimensions[src_row + TOP_PAD_ROWS].height = h

    # 列: 残す列 → 新列の幅
    for new_idx, (_, src_col) in enumerate(keep_pairs, start=LEFT_PAD_COLS + 1):
        src_letter = get_column_letter(src_col)
        new_letter = get_column_letter(new_idx)
        w = src_ws.column_dimensions[src_letter].width
        if w:
            dst_ws.column_dimensions[new_letter].width = w

    return missing


def transform(input_path: Path, output_path: Path) -> None:
    src = load_workbook(input_path)
    src_ws = src.active

    dst = Workbook()
    dst_ws = dst.active
    dst_ws.title = src_ws.title or "Sheet1"

    apply_transform_to_sheet(src_ws, dst_ws)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    dst.save(output_path)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("input", type=Path, help="入力 xlsx のパス")
    parser.add_argument("-o", "--output", type=Path, help="出力 xlsx のパス")
    parser.add_argument("--inspect", action="store_true", help="列構成だけ表示して終了")
    args = parser.parse_args()

    if not args.input.exists():
        print(f"ERROR: not found: {args.input}", file=sys.stderr)
        return 2

    if args.inspect:
        inspect(args.input)
        return 0

    out = args.output or PROCESSED_DIR / f"{args.input.stem}_processed.xlsx"
    transform(args.input, out)
    print(f"OK saved: {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
