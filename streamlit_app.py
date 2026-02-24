#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
商場年業績象限分析（互動）
- 黑底介面（Streamlit），圖表白底（Plotly）
- 預設使用 repo 內建 Excel（與本檔同資料夾）：商場年業績表_含成長率.xlsx
- 亦可上傳 Excel 覆蓋內建資料
- 讀取工作表：篩選結果
- 計算：2025成長率 = (2025 - 2024) / 2024
- 篩選：類型 -> 區 -> 縣市（連動）
- 象限切分：平均值 / 中位數 / 自定義（自定義固定套用所有篩選情境）
- 圖表：散點圖（白底），顏色=縣市（固定顏色、不因篩選變動）；支援下載圖檔
- 圖表下方：
  1) 四象限統計表（象限分類 / 象限定義 / 商場數量）
  2) 明細表（含象限欄位）
"""

from __future__ import annotations

from pathlib import Path
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px


# =========================
# Config
# =========================
DEFAULT_FILE = "商場年業績表_含成長率.xlsx"
SHEET_NAME = "篩選結果"

TYPE_COL = "類型"
AREA_COL = "區"
CITY_COL = "縣市"

NAME_COL = "商場名稱"
SYS_COL = "體系"

COL_2024 = "2024"
COL_2025 = "2025"
GROWTH_COL = "2025成長率"  # internal numeric ratio (0.12 = 12%)


# =========================
# Theme (dark app)
# =========================
def inject_dark_theme():
    st.markdown(
        """
        <style>
          .stApp { background-color: #0e1117; color: #ffffff; }
          [data-testid="stSidebar"] { background-color: #0e1117; }
          h1, h2, h3, h4, h5, h6, p, div, span, label { color: #ffffff !important; }
          .note { font-size: 12px; opacity: 0.85; margin-top: -6px; }
        </style>
        """,
        unsafe_allow_html=True,
    )


def get_builtin_excel_path() -> Path:
    return Path(__file__).resolve().parent / DEFAULT_FILE


@st.cache_data(show_spinner=False)
def load_df(file_ref, sheet: str) -> pd.DataFrame:
    try:
        import openpyxl  # noqa: F401
    except Exception as e:
        raise RuntimeError("Missing optional dependency 'openpyxl'. Please add openpyxl to requirements.txt") from e

    df = pd.read_excel(file_ref, sheet_name=sheet)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def to_numeric_series(s: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(s):
        return s.astype(float)

    s2 = s.astype(str).str.strip()
    s2 = s2.replace({"": np.nan, "nan": np.nan, "None": np.nan, "NA": np.nan, "N/A": np.nan})
    is_pct = s2.str.contains("%", na=False)
    s2 = s2.str.replace(",", "", regex=False).str.replace("%", "", regex=False)
    out = pd.to_numeric(s2, errors="coerce")
    out = np.where(is_pct, out / 100.0, out)
    return pd.Series(out, index=s.index, dtype="float64")


def multiselect_with_all_sidebar(label: str, options, key_prefix: str, default_all: bool = True):
    """Sidebar 多選 + 全選/全不選；並確保選項變動時不會因 default 不在 options 而報錯"""
    options = list(options)
    st.sidebar.markdown(f"**{label}**")
    c1, c2 = st.sidebar.columns(2)
    key_ms = f"{key_prefix}_ms"

    if key_ms not in st.session_state:
        st.session_state[key_ms] = options[:] if default_all else []

    current = st.session_state.get(key_ms, [])
    current = [v for v in current if v in options]

    if default_all and len(current) == 0 and len(options) > 0:
        current = options[:]

    st.session_state[key_ms] = current

    with c1:
        if st.button("全選", key=f"{key_prefix}_btn_all"):
            st.session_state[key_ms] = options[:]
    with c2:
        if st.button("全不選", key=f"{key_prefix}_btn_none"):
            st.session_state[key_ms] = []

    sel = st.sidebar.multiselect(
        label="",
        options=options,
        default=st.session_state[key_ms],
        key=key_ms,
    )
    return sel


def compute_cut_values(df_plot: pd.DataFrame, mode: str, custom_growth_pct: float | None, custom_y: float | None):
    """
    mode: 平均值 / 中位數 / 自定義
    回傳 x_cut(成長率ratio) / y_cut(2025業績)
    """
    if len(df_plot) == 0:
        return None, None

    if mode == "平均值":
        return float(df_plot[GROWTH_COL].mean()), float(df_plot[COL_2025].mean())
    if mode == "中位數":
        return float(df_plot[GROWTH_COL].median()), float(df_plot[COL_2025].median())

    # 自定義（固定）
    x_cut = None if custom_growth_pct is None else float(custom_growth_pct) / 100.0
    y_cut = None if custom_y is None else float(custom_y)
    return x_cut, y_cut


def assign_quadrant(df_plot: pd.DataFrame, x_cut: float, y_cut: float) -> pd.Series:
    conds = [
        (df_plot[GROWTH_COL] >= x_cut) & (df_plot[COL_2025] >= y_cut),
        (df_plot[GROWTH_COL] < x_cut) & (df_plot[COL_2025] >= y_cut),
        (df_plot[GROWTH_COL] < x_cut) & (df_plot[COL_2025] < y_cut),
        (df_plot[GROWTH_COL] >= x_cut) & (df_plot[COL_2025] < y_cut),
    ]
    labels = ["第一象限", "第二象限", "第三象限", "第四象限"]
    return pd.Series(np.select(conds, labels, default="未分類"), index=df_plot.index)


def build_fixed_city_palette(city_order: list[str]) -> dict[str, str]:
    """依 city_order 建立固定顏色映射（同城市永遠同顏色）"""
    palette = (
        px.colors.qualitative.Plotly
        + px.colors.qualitative.D3
        + px.colors.qualitative.Set3
        + px.colors.qualitative.Dark24
        + px.colors.qualitative.Light24
    )
    return {c: palette[i % len(palette)] for i, c in enumerate(city_order)}


def main():
    st.set_page_config(page_title="商場年業績象限分析", layout="wide")
    inject_dark_theme()

    st.title("商場年業績象限分析（互動）")
    st.caption("圖表白底、介面黑底；象限分界可切換平均 / 中位數 / 自定義。")

    # =========================
    # Sidebar: Data source
    # =========================
    st.sidebar.header("資料來源")
    builtin_path = get_builtin_excel_path()

    uploaded = st.sidebar.file_uploader("覆蓋內建資料（可選）", type=["xlsx"])
    if uploaded is not None:
        file_ref = uploaded
        file_name = uploaded.name
    else:
        if not builtin_path.exists():
            st.error(
                f"找不到內建資料檔：{builtin_path}\n\n"
                "請確認 Excel 已推到 GitHub repo 且與 app 檔在同一資料夾，或改用上傳檔案。"
            )
            st.stop()
        file_ref = str(builtin_path)
        file_name = builtin_path.name

    sheet = st.sidebar.text_input("工作表名稱", value=SHEET_NAME)

    try:
        df = load_df(file_ref, sheet)
    except Exception as e:
        st.error(f"讀取資料失敗：{e}")
        st.stop()

    # =========================
    # Validate + prepare
    # =========================
    need_cols = [TYPE_COL, AREA_COL, CITY_COL, NAME_COL, SYS_COL, COL_2024, COL_2025]
    missing = [c for c in need_cols if c not in df.columns]
    if missing:
        st.error(f"缺少欄位：{missing}\n目前欄位：{list(df.columns)}")
        st.stop()

    df = df.copy()
    df[COL_2024] = to_numeric_series(df[COL_2024])
    df[COL_2025] = to_numeric_series(df[COL_2025])

    # 計算成長率（ratio）
    df[GROWTH_COL] = (df[COL_2025] - df[COL_2024]) / df[COL_2024]
    df.loc[df[COL_2024] == 0, GROWTH_COL] = np.nan

    # =========================
    # 固定顏色 & 圖例排序（以「全資料」的縣市為準，不因篩選變動）
    # =========================
    all_cities = sorted(df[CITY_COL].dropna().astype(str).unique().tolist())

    # 自訂縣市排序（可選）：若想固定特定順序，把 None 改成你的 list
    custom_city_order = None  # e.g. ["台北市","新北市",...]
    city_order = custom_city_order if custom_city_order else all_cities

    # 固定顏色映射：存入 session_state，避免因 rerun 改色
    if "city_color_map" not in st.session_state:
        st.session_state["city_color_map"] = build_fixed_city_palette(city_order)
    else:
        # 若資料來源換檔帶來新縣市，補色但不改舊色
        color_map = dict(st.session_state["city_color_map"])
        for c in city_order:
            if c not in color_map:
                color_map.update(build_fixed_city_palette([c]))
        st.session_state["city_color_map"] = color_map

    # =========================
    # Sidebar: filters (linked)
    # =========================
    st.sidebar.header("篩選器（連動）")

    # (1) 類型
    type_opts = sorted(df[TYPE_COL].dropna().astype(str).unique().tolist())
    type_pick = multiselect_with_all_sidebar("類型（多選）", type_opts, "type", default_all=True)
    fdf = df[df[TYPE_COL].astype(str).isin(type_pick)].copy() if type_pick else df.iloc[0:0].copy()

    # (2) 區
    area_opts = sorted(fdf[AREA_COL].dropna().astype(str).unique().tolist())
    area_pick = multiselect_with_all_sidebar("區（多選）", area_opts, "area", default_all=True)
    fdf = fdf[fdf[AREA_COL].astype(str).isin(area_pick)].copy() if area_pick else fdf.iloc[0:0].copy()

    # (3) 縣市
    city_opts = sorted(fdf[CITY_COL].dropna().astype(str).unique().tolist())
    city_pick = multiselect_with_all_sidebar("縣市（多選）", city_opts, "city", default_all=True)
    fdf = fdf[fdf[CITY_COL].astype(str).isin(city_pick)].copy() if city_pick else fdf.iloc[0:0].copy()

    # only keep rows with x/y
    fdf = fdf.dropna(subset=[GROWTH_COL, COL_2025]).copy()

    if len(fdf) == 0:
        st.warning("目前篩選結果為空，請調整左側篩選條件。")
        st.stop()

    st.caption(f"目前資料：{file_name}｜工作表：{sheet}｜篩選後筆數：{len(fdf):,}")

    # =========================
    # Sidebar: quadrant split
    # =========================
    st.sidebar.markdown("---")
    st.sidebar.header("象限切分")
    split_mode = st.sidebar.radio("切分方式", ["平均值", "中位數", "自定義"], index=0)

    if "custom_growth_pct" not in st.session_state:
        st.session_state["custom_growth_pct"] = 0.0
    if "custom_2025" not in st.session_state:
        st.session_state["custom_2025"] = 0.0

    custom_growth = None
    custom_y = None
    if split_mode == "自定義":
        st.sidebar.caption("自定義切分值會固定套用在所有篩選情境")
        st.session_state["custom_growth_pct"] = st.sidebar.number_input(
            "成長率分界（%）",
            value=float(st.session_state["custom_growth_pct"]),
            step=1.0,
            format="%.2f",
        )
        st.session_state["custom_2025"] = st.sidebar.number_input(
            "2025業績分界",
            value=float(st.session_state["custom_2025"]),
            step=1000.0,
            format="%.2f",
        )
        custom_growth = float(st.session_state["custom_growth_pct"])
        custom_y = float(st.session_state["custom_2025"])

    x_cut, y_cut = compute_cut_values(fdf, split_mode, custom_growth, custom_y)

    if x_cut is None or y_cut is None or np.isnan(x_cut) or np.isnan(y_cut):
        st.warning("目前資料不足以計算象限分界值（可能都缺 2025成長率或 2025業績）。")
        st.stop()

    # assign quadrant
    fdf = fdf.copy()
    fdf["象限"] = assign_quadrant(fdf, x_cut, y_cut)

    # =========================
    # Main: metrics
    # =========================
    c1, c2, c3 = st.columns(3)
    c1.metric("X 分界值（成長率）", f"{x_cut:.2%}")
    c2.metric("Y 分界值（2025業績）", f"{y_cut:,.2f}")
    c3.metric("商場數量", f"{len(fdf):,}")

    # =========================
    # Plot
    # =========================
    st.subheader("散點圖（可下載圖檔）")

    show_labels = st.sidebar.toggle("顯示資料標籤（商場名稱(體系)）", value=False)

    # label field (for point labels)
    fdf["_label"] = (
        fdf[NAME_COL].astype(str).fillna("").str.strip()
        + "("
        + fdf[SYS_COL].astype(str).fillna("").str.strip()
        + ")"
    )

    # Hover：顯示所有原始欄位（隱藏內部輔助欄位，避免干擾）
    hover_cols = {col: True for col in fdf.columns}
    for _hide in ["_label"]:
        if _hide in hover_cols:
            hover_cols[_hide] = False

    fig = px.scatter(
        fdf,
        x=GROWTH_COL,
        y=COL_2025,
        color=CITY_COL,  # ⭐ 顏色=縣市（圖例也會跟著變）
        text="_label" if show_labels else None,
        hover_data=hover_cols,
        labels={GROWTH_COL: "2025成長率", COL_2025: "2025業績", CITY_COL: "縣市"},
        title=f"散點圖（{split_mode}分界｜顏色=縣市）",
        category_orders={CITY_COL: city_order},  # ⭐ 圖例依縣市排序
        color_discrete_map=st.session_state["city_color_map"],  # ⭐ 固定縣市顏色
    )

    # quadrant lines
    fig.add_vline(
        x=x_cut,
        line_dash="dash",
        line_color="black",
        annotation_text=f"X分界: {x_cut:.2%}",
        annotation_position="top left",
    )
    fig.add_hline(
        y=y_cut,
        line_dash="dash",
        line_color="black",
        annotation_text=f"Y分界: {y_cut:,.2f}",
        annotation_position="bottom right",
    )

    # Enforce white chart + black text (avoid CSS interference)
    fig.update_layout(
        template="plotly_white",
        paper_bgcolor="white",
        plot_bgcolor="white",
        font=dict(color="black"),
        legend=dict(
            title=dict(text="縣市", font=dict(color="black")),
            font=dict(color="black", size=10),
            traceorder="normal",
        ),
        hovermode="closest",
        height=850,
        margin=dict(l=40, r=260, t=70, b=40),
    )

    fig.update_xaxes(
        tickformat=".0%",
        showgrid=False,
        title_font=dict(color="black"),
        tickfont=dict(color="black"),
        linecolor="black",
    )
    fig.update_yaxes(
        showgrid=False,
        title_font=dict(color="black"),
        tickfont=dict(color="black"),
        linecolor="black",
    )

    # Bigger dots + readable labels
    fig.update_traces(marker=dict(size=12))
    fig.update_traces(textfont=dict(color="black"))

    if show_labels:
        fig.update_traces(mode="markers+text", textposition="top center")
    else:
        fig.update_traces(mode="markers")

    # Enable image download (modebar)
    config = {
        "displaylogo": False,
        "toImageButtonOptions": {
            "format": "png",
            "filename": "quadrant_scatter",
            "height": 1200,
            "width": 1800,
            "scale": 2,
        },
    }

    st.plotly_chart(fig, use_container_width=True, config=config)

    # =========================
    # Quadrant summary table
    # =========================
    st.subheader("四象限統計")

    q_meta = pd.DataFrame(
        [
            ["第一象限", "高業績, 高成長"],
            ["第二象限", "高業績, 低成長"],
            ["第三象限", "低業績, 低成長"],
            ["第四象限", "低業績, 高成長"],
        ],
        columns=["象限分類", "象限定義"],
    )

    q_cnt = (
        fdf["象限"]
        .value_counts()
        .rename_axis("象限分類")
        .reset_index(name="商場數量")
    )

    q_table = q_meta.merge(q_cnt, on="象限分類", how="left").fillna({"商場數量": 0})
    q_table["商場數量"] = q_table["商場數量"].astype(int)

    st.dataframe(q_table, width="stretch")

    # =========================
    # Detail table under chart
    # =========================
    st.subheader("篩選後商場明細（含象限）")

    prefer_cols = [
        TYPE_COL, SYS_COL, NAME_COL, AREA_COL, CITY_COL,
        "行政區", "地址", COL_2024, COL_2025, GROWTH_COL, "象限"
    ]
    cols_exist = [c for c in prefer_cols if c in fdf.columns]
    df_table = fdf[cols_exist].copy()
    if GROWTH_COL in df_table.columns:
        df_table[GROWTH_COL] = df_table[GROWTH_COL].map(lambda v: "" if pd.isna(v) else f"{v:.2%}")

    st.dataframe(df_table, width="stretch")


if __name__ == "__main__":
    main()
