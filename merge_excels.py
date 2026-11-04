"""
複数の Astream 提案リスト xlsx を 1 ブックにまとめる。
各入力ファイルは、整形（process_excel と同じ仕様）された上で別シートに配置される。

シート名はファイル名から決まる:
  CCZ_マラソン・ジョギング.xlsx → "マラソン・ジョギング"
  CCZ_yoga.xlsx                  → "yoga"

  python merge_excels.py file1.xlsx file2.xlsx ... -o merged.xlsx
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

from openpyxl import Workbook, load_workbook

from process_excel import PROCESSED_DIR, apply_transform_to_sheet

ILLEGAL_SHEET_CHARS = re.compile(r"[:\\/?*\[\]]")


def sheet_name_from_path(path: Path, used: set[str]) -> str:
    name = path.stem
    name = re.sub(r"^CCZ[_\-\s]*", "", name, flags=re.IGNORECASE)
    name = ILLEGAL_SHEET_CHARS.sub("_", name).strip()
    if not name:
        name = "Sheet"
    name = name[:31]

    # 重複回避
    base = name
    i = 2
    while name in used:
        suffix = f"_{i}"
        name = base[: 31 - len(suffix)] + suffix
        i += 1
    used.add(name)
    return name


def merge(inputs: list[Path], output_path: Path) -> None:
    dst = Workbook()
    # デフォルトで作られる空シートは最後に削除する
    default_sheet = dst.active

    used_names: set[str] = set()
    for src_path in inputs:
        if not src_path.exists():
            print(f"WARN: 見つからないのでスキップ: {src_path}", file=sys.stderr)
            continue
        print(f"処理中: {src_path.name}")
        src = load_workbook(src_path)
        src_ws = src.active

        sheet_name = sheet_name_from_path(src_path, used_names)
        dst_ws = dst.create_sheet(title=sheet_name)
        apply_transform_to_sheet(src_ws, dst_ws)

    if len(dst.sheetnames) > 1:
        dst.remove(default_sheet)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    dst.save(output_path)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("inputs", nargs="+", type=Path, help="入力 xlsx のパス（複数）")
    parser.add_argument(
        "-o", "--output", type=Path, help="出力 xlsx のパス（デフォルト: processed/CCZ_merged.xlsx）"
    )
    args = parser.parse_args()

    out = args.output or PROCESSED_DIR / "CCZ_merged.xlsx"
    merge(args.inputs, out)
    print(f"OK saved: {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
