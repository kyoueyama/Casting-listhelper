"""
Streamlit Web アプリ: Excel をアップロード → 残す列を選択/並び替え → 整形 → ダウンロード。

起動:
    streamlit run app.py
"""

from __future__ import annotations

import io
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import Workbook, load_workbook

from process_excel import KEEP_COLUMNS, apply_transform_to_sheet

st.set_page_config(page_title="Casting List Helper", layout="wide")
st.title("Casting List Helper")
st.caption("Astream の提案リスト Excel を整形します。残す列を選び、ダウンロードしてください。")

uploaded = st.file_uploader("Excel ファイル (.xlsx)", type=["xlsx"])

if uploaded is None:
    st.info("Excel をアップロードしてください。")
    st.stop()

file_bytes = uploaded.read()

# ヘッダーだけ読み出して列一覧を取得
try:
    wb_preview = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
except Exception as e:
    st.error(f"Excel を読み込めませんでした: {e}")
    st.stop()

ws_preview = wb_preview.active
headers: list[str] = []
for cell in ws_preview[1]:
    if cell.value is None:
        continue
    name = str(cell.value).strip()
    if name and name not in headers:
        headers.append(name)
wb_preview.close()

if not headers:
    st.error("1 行目にヘッダーが見つかりませんでした。")
    st.stop()

st.subheader(f"列の選択と並び替え（全 {len(headers)} 列）")
st.caption(
    "「残す」にチェックを入れた列だけが出力されます。"
    "「順序」の小さい順に左から並びます（同値の場合はアップロード順）。"
)

# 初期値: KEEP_COLUMNS にあるものは ON + その並び順、それ以外は OFF
default_rows = []
big = len(KEEP_COLUMNS) + 1
for name in headers:
    if name in KEEP_COLUMNS:
        order = KEEP_COLUMNS.index(name) + 1
        keep = True
    else:
        order = big
        big += 1
        keep = False
    default_rows.append({"残す": keep, "順序": order, "列名": name})

# data_editor のキーを uploaded.name に紐づけて、ファイル変更時に編集状態をリセット
editor_key = f"editor::{uploaded.name}::{len(headers)}"

edited = st.data_editor(
    pd.DataFrame(default_rows),
    column_config={
        "残す": st.column_config.CheckboxColumn("残す", default=False, width="small"),
        "順序": st.column_config.NumberColumn("順序", min_value=1, step=1, width="small"),
        "列名": st.column_config.TextColumn("列名", disabled=True, width="large"),
    },
    hide_index=True,
    use_container_width=True,
    num_rows="fixed",
    key=editor_key,
)

selected = (
    edited[edited["残す"] == True]  # noqa: E712
    .sort_values(["順序"], kind="stable")
)
keep_columns: list[str] = selected["列名"].tolist()

col_l, col_r = st.columns([2, 1])
with col_l:
    st.markdown(
        f"**出力される列 ({len(keep_columns)} 列)**: "
        + ("、".join(keep_columns) if keep_columns else "_（未選択）_")
    )
with col_r:
    run = st.button(
        "整形して出力",
        type="primary",
        disabled=len(keep_columns) == 0,
        use_container_width=True,
    )

if run:
    try:
        src_wb = load_workbook(io.BytesIO(file_bytes))
        src_ws = src_wb.active
        dst_wb = Workbook()
        dst_ws = dst_wb.active
        dst_ws.title = src_ws.title or "Sheet1"

        missing = apply_transform_to_sheet(src_ws, dst_ws, keep_columns=keep_columns)

        out_buf = io.BytesIO()
        dst_wb.save(out_buf)
        out_buf.seek(0)
    except Exception as e:
        st.error(f"整形中にエラーが発生しました: {e}")
        st.stop()

    if missing:
        st.warning(f"元データに存在しなかった列: {missing}")

    out_name = f"{Path(uploaded.name).stem}_processed.xlsx"
    st.success("整形が完了しました。")
    st.download_button(
        "ダウンロード",
        data=out_buf,
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
