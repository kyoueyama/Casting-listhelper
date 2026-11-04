"""Post_it_*.xlsx (Astream 生フォーマット) を 1 シートに統合する。

process_excel.py と同じ整形（左に2列パディング、上に1行パディング、KEEP_COLUMNS で絞る、
画像をフォト列に再配置）を行いつつ、複数ファイルを縦に連結する。
出所が分かるように「カテゴリー」列をフォト列の右に挿入する。
カテゴリー名はファイル名 (Post_it_<category>.xlsx) から取る。

  python merge_postit.py file1.xlsx file2.xlsx ... -o merged.xlsx
"""

from __future__ import annotations

import argparse
import io
import re
import sys
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

from process_excel import (
    DEFAULT_FONT_NAME,
    KEEP_COLUMNS,
    LEFT_PAD_COLS,
    PROCESSED_DIR,
    SRC_HEADER_ROW,
    TOP_PAD_ROWS,
    collect_images_by_row,
)

CATEGORY_COL_NAME = "カテゴリー"


def category_from_filename(path: Path) -> str:
    name = path.stem
    name = re.sub(r"^Post_it[_\-\s]*", "", name, flags=re.IGNORECASE)
    return name or path.stem


def build_output_columns() -> list[str]:
    """出力列: 名前, フォト, カテゴリー, アカウント名, ... の順。"""
    cols = list(KEEP_COLUMNS)
    # フォトの直後にカテゴリーを挿入（先頭付近に置いて視認性を上げる）
    if "フォト" in cols:
        i = cols.index("フォト")
        cols.insert(i + 1, CATEGORY_COL_NAME)
    else:
        cols.insert(0, CATEGORY_COL_NAME)
    return cols


def merge(inputs: list[Path], output_path: Path) -> None:
    out_cols = build_output_columns()
    dst = Workbook()
    dst_ws = dst.active
    dst_ws.title = "merged"

    header_row = SRC_HEADER_ROW + TOP_PAD_ROWS  # = 2

    # B2 = "No"
    dst_ws.cell(row=header_row, column=2, value="No")
    # ヘッダー（C2 から）
    new_photo_col = None
    new_category_col = None
    for new_idx, name in enumerate(out_cols, start=LEFT_PAD_COLS + 1):
        dst_ws.cell(row=header_row, column=new_idx, value=name)
        if name == "フォト":
            new_photo_col = new_idx
        if name == CATEGORY_COL_NAME:
            new_category_col = new_idx

    total_cols = LEFT_PAD_COLS + len(out_cols)
    dst_row = header_row  # 次に書き込む行 - 1
    no_counter = 0
    column_widths: dict[int, float] = {}

    # 出力列名→出力列番号(1-indexed) のマップ
    out_col_map = {
        name: idx for idx, name in enumerate(out_cols, start=LEFT_PAD_COLS + 1)
    }

    for src_path in inputs:
        if not src_path.exists():
            print(f"WARN: 見つからないのでスキップ: {src_path}", file=sys.stderr)
            continue
        category = category_from_filename(src_path)
        print(f"処理中: {src_path.name}  (カテゴリー: {category})")

        src = load_workbook(src_path)
        src_ws = src.active

        # 元のヘッダー → 列番号
        headers = {}
        for cell in src_ws[SRC_HEADER_ROW]:
            if cell.value is not None:
                headers[str(cell.value).strip()] = cell.column

        photo_col_src = headers.get("フォト")
        images_by_src_row = (
            collect_images_by_row(src_ws, photo_col_src) if photo_col_src else {}
        )

        # 列幅の引き継ぎ用（最後に書いた値で確定）
        for name in out_cols:
            if name == CATEGORY_COL_NAME:
                continue
            src_col = headers.get(name)
            if src_col is None:
                continue
            src_letter = get_column_letter(src_col)
            w = src_ws.column_dimensions[src_letter].width
            if w:
                column_widths[out_col_map[name]] = w

        last_src_row = src_ws.max_row
        for src_row in range(SRC_HEADER_ROW + 1, last_src_row + 1):
            # 空行（A 列もデータもない行）はスキップ
            row_is_empty = all(
                src_ws.cell(row=src_row, column=c).value in (None, "")
                for c in headers.values()
            )
            if row_is_empty and src_row not in images_by_src_row:
                continue

            dst_row += 1
            no_counter += 1
            dst_ws.cell(row=dst_row, column=2, value=no_counter)
            dst_ws.cell(row=dst_row, column=new_category_col, value=category)

            for name in out_cols:
                if name == CATEGORY_COL_NAME:
                    continue
                src_col = headers.get(name)
                if src_col is None:
                    continue
                src_cell = src_ws.cell(row=src_row, column=src_col)
                dst_cell = dst_ws.cell(
                    row=dst_row, column=out_col_map[name], value=src_cell.value
                )
                if src_cell.number_format and src_cell.number_format != "General":
                    dst_cell.number_format = src_cell.number_format

            # 行高を引き継ぐ
            h = src_ws.row_dimensions[src_row].height
            if h:
                dst_ws.row_dimensions[dst_row].height = h

            # 画像の再配置
            if new_photo_col and src_row in images_by_src_row:
                src_img = images_by_src_row[src_row]
                try:
                    img_bytes = src_img._data()
                    bio = io.BytesIO(img_bytes)
                    new_img = XLImage(bio)
                    if getattr(src_img, "width", None):
                        new_img.width = src_img.width
                    if getattr(src_img, "height", None):
                        new_img.height = src_img.height
                    new_photo_letter = get_column_letter(new_photo_col)
                    dst_ws.add_image(new_img, f"{new_photo_letter}{dst_row}")
                except Exception as e:
                    print(
                        f"WARN: 画像の再配置に失敗 ({src_path.name} row={src_row}): {e}",
                        file=sys.stderr,
                    )

    # スタイル: ヘッダー行は折り返し+中央寄せ、その他は中央寄せ
    last_new_row = dst_row
    header_align = Alignment(wrap_text=True, vertical="center", horizontal="center")
    center_align = Alignment(vertical="center", horizontal="center")
    default_font = Font(name=DEFAULT_FONT_NAME)
    for row in range(1, last_new_row + 1):
        for col in range(1, total_cols + 1):
            cell = dst_ws.cell(row=row, column=col)
            cell.font = default_font
            cell.alignment = header_align if row == header_row else center_align

    # 列幅
    for col_idx, w in column_widths.items():
        dst_ws.column_dimensions[get_column_letter(col_idx)].width = w
    # カテゴリー列の幅をそれっぽく
    if new_category_col is not None:
        cat_letter = get_column_letter(new_category_col)
        if not dst_ws.column_dimensions[cat_letter].width:
            dst_ws.column_dimensions[cat_letter].width = 14

    output_path.parent.mkdir(parents=True, exist_ok=True)
    dst.save(output_path)
    print(f"OK saved: {output_path}  ({no_counter} 行)")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("inputs", nargs="+", type=Path, help="入力 xlsx のパス（複数）")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="出力 xlsx のパス（デフォルト: processed/Post_it_merged.xlsx）",
    )
    args = parser.parse_args()

    out = args.output or PROCESSED_DIR / "Post_it_merged.xlsx"
    merge(args.inputs, out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
