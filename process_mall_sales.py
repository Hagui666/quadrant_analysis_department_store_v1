#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
從指定 Excel 檔案中讀取「地區型商場」「百貨」兩張工作表（雙重欄位，從第2列作為欄位名稱），
依規則篩選並輸出到新工作表「篩選結果」。

規則：
1) 有業績數據 == True
2) 2024 或 2025 任一欄位有值（只要其中之一有資料即可）
輸出欄位：
體系、商場名稱、區、縣市、行政區、地址、2024、2025
另外新增欄位：類型（資料來源工作表名稱）
"""

import argparse
import pandas as pd


SHEETS = ["地區型商場", "百貨"]
OUT_SHEET = "篩選結果"
KEEP_COLS = ["體系", "商場名稱", "區", "縣市", "行政區", "地址", "2024", "2025"]


def _to_bool_series(s: pd.Series) -> pd.Series:
    """
    把可能是 bool / 0-1 / 'true'/'false' / 'TRUE' 等各種型態轉成布林值。
    無法解析者視為 False。
    """
    if s.dtype == bool:
        return s.fillna(False)

    if pd.api.types.is_numeric_dtype(s):
        return (s.fillna(0) != 0)

    ss = s.astype(str).str.strip().str.lower()
    true_set = {"true", "1", "yes", "y", "是", "t"}
    false_set = {"false", "0", "no", "n", "否", "f", "nan", "none", ""}

    return ss.map(lambda x: True if x in true_set else (False if x in false_set else False)).fillna(False)


def _has_value_series(s: pd.Series) -> pd.Series:
    """
    判斷欄位是否「有值」：
    - 非 NaN
    - 非空字串
    """
    if pd.api.types.is_numeric_dtype(s):
        return s.notna()

    ss = s.astype(str)
    return s.notna() & (ss.str.strip() != "") & (ss.str.lower().str.strip() != "nan")


def process_file(input_path: str, output_path: str | None = None) -> str:
    from pathlib import Path

    in_path = Path(input_path)
    if not in_path.exists():
        raise FileNotFoundError(f"找不到檔案：{in_path}")

    if output_path is None:
        output_path = str(in_path.with_name(in_path.stem + "_篩選結果" + in_path.suffix))
    out_path = Path(output_path)

    xls = pd.ExcelFile(in_path)

    frames = []
    for sheet in SHEETS:
        if sheet not in xls.sheet_names:
            raise ValueError(f"工作表不存在：{sheet}（現有：{xls.sheet_names}）")

        # 雙重欄位：第2列當欄位名稱（header=1，0-based）
        df = pd.read_excel(in_path, sheet_name=sheet, header=1)
        df.columns = [str(c).strip() for c in df.columns]

        required = set(KEEP_COLS + ["有業績數據"])
        missing = [c for c in required if c not in df.columns]
        if missing:
            raise ValueError(f"工作表「{sheet}」缺少欄位：{missing}")

        # 1) 有業績數據 == True
        flag = _to_bool_series(df["有業績數據"])
        df1 = df.loc[flag].copy()

        # 2) 2024 或 2025 任一有值
        has_2024 = _has_value_series(df1["2024"])
        has_2025 = _has_value_series(df1["2025"])
        df2 = df1.loc[has_2024 | has_2025].copy()

        out_df = df2[KEEP_COLS].copy()
        out_df.insert(0, "類型", sheet)
        frames.append(out_df)

    result = pd.concat(frames, ignore_index=True)

    # 寫出：保留原工作表 + 新增「篩選結果」
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for s in xls.sheet_names:
            pd.read_excel(in_path, sheet_name=s, header=None).to_excel(
                writer, sheet_name=s, index=False, header=False
            )
        result.to_excel(writer, sheet_name=OUT_SHEET, index=False)

    return str(out_path)


def main():
    parser = argparse.ArgumentParser(
        description="篩選 Excel 檔案中「地區型商場」「百貨」工作表，輸出到新工作表「篩選結果」。"
    )
    parser.add_argument("input", help="輸入 Excel 檔案路徑，例如：商場年業績表.xlsx")
    parser.add_argument(
        "-o", "--output", default=None, help="輸出 Excel 檔案路徑（預設：在原檔名後加 _篩選結果）"
    )
    args = parser.parse_args()

    out = process_file(args.input, args.output)
    print(f"完成，輸出檔案：{out}")


if __name__ == "__main__":
    main()
