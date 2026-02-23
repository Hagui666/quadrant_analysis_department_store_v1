# streamlit_app.py
# -*- coding: utf-8 -*-
"""
Streamlit 互動式散點圖（黑底介面 + 層級篩選，圖表白底）：
- 讀取 Excel 的「篩選結果」工作表
- 計算「2025成長率」 = (2025 - 2024) / 2024
- 依層級篩選：類型 -> 區 -> 縣市（高層篩選會影響低層選項）
- 繪製散點圖：X=2025成長率(%)，Y=2025（y軸標籤：2025業績）
- 按鈕切換：在圖上顯示/隱藏「商場名稱」
- Tooltip 顯示：該點所有欄位（含成長率以百分比呈現）
- 表格中所有成長率欄位以百分比字串顯示
"""

import pandas as pd
import streamlit as st
import altair as alt

DEFAULT_FILE = "商場年業績表_含成長率.xlsx"
DEFAULT_SHEET = "篩選結果"


def _inject_dark_theme():
    # 強制黑底 + 白字（包含 sidebar）；圖表底色另外在 Altair 設為白底
    st.markdown(
        """
        <style>
        .stApp { background-color: #0e1117; color: #ffffff; }
        section[data-testid="stSidebar"] { background-color: #0e1117; color: #ffffff; }
        section[data-testid="stSidebar"] * { color: #ffffff !important; }
        .stMarkdown, .stText, .stCaption, .stMetric, .stDataFrame, .stTable { color: #ffffff; }
        /* Dataframe header text a bit brighter */
        div[data-testid="stDataFrame"] * { color: #ffffff; }
        </style>
        """,
        unsafe_allow_html=True,
    )


@st.cache_data(show_spinner=False)
def load_data(file_ref, sheet: str) -> pd.DataFrame:
    df = pd.read_excel(file_ref, sheet_name=sheet)

    # 清理欄位名稱
    df.columns = [str(c).strip() for c in df.columns]

    # 轉型
    for col in ["2024", "2025"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # 成長率
    if "2024" in df.columns and "2025" in df.columns:
        df["2025成長率"] = (df["2025"] - df["2024"]) / df["2024"]
        df.loc[df["2024"] == 0, "2025成長率"] = pd.NA  # 避免除以0

    return df


def toggle_show_labels():
    st.session_state["show_labels"] = not st.session_state.get("show_labels", False)


def to_percent_str(x, digits=2):
    if pd.isna(x):
        return ""
    try:
        return f"{float(x):.{digits}%}"
    except Exception:
        return ""


def main():
    st.set_page_config(page_title="商場業績分析", layout="wide")
    _inject_dark_theme()

    st.title("商場業績分析：2025成長率 vs 2025業績")

    # Sidebar：資料來源
    st.sidebar.header("資料來源")
    uploaded = st.sidebar.file_uploader("上傳 Excel 檔（可選）", type=["xlsx"])
    if uploaded is not None:
        file_ref = uploaded
        file_name_for_info = uploaded.name
    else:
        file_ref = DEFAULT_FILE
        file_name_for_info = DEFAULT_FILE

    sheet = st.sidebar.text_input("工作表名稱", value=DEFAULT_SHEET)

    try:
        df = load_data(file_ref, sheet)
    except Exception as e:
        st.error(f"讀取資料失敗：{e}")
        st.stop()

    required_cols = {"類型", "區", "縣市", "2024", "2025", "2025成長率", "商場名稱"}
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        st.error(f"資料缺少必要欄位：{missing}")
        st.stop()

    st.caption(f"目前檔案：{file_name_for_info}｜工作表：{sheet}｜總筆數：{len(df):,}")

    # === 層級篩選（類型 -> 區 -> 縣市）===
    st.sidebar.header("篩選條件（高層會影響低層選項）")

    type_options = sorted([x for x in df["類型"].dropna().unique().tolist()])
    selected_types = st.sidebar.multiselect("類型", options=type_options, default=type_options)
    df_f1 = df[df["類型"].isin(selected_types)] if selected_types else df.iloc[0:0]

    area_options = sorted([x for x in df_f1["區"].dropna().unique().tolist()])
    selected_areas = st.sidebar.multiselect("區", options=area_options, default=area_options)
    df_f2 = df_f1[df_f1["區"].isin(selected_areas)] if selected_areas else df_f1.iloc[0:0]

    city_options = sorted([x for x in df_f2["縣市"].dropna().unique().tolist()])
    selected_cities = st.sidebar.multiselect("縣市", options=city_options, default=city_options)
    df_filtered = df_f2[df_f2["縣市"].isin(selected_cities)] if selected_cities else df_f2.iloc[0:0]

    # 顯示選項
    st.sidebar.header("顯示選項")
    only_valid = st.sidebar.checkbox(
        "只顯示可計算正常成長率（2024/2025皆有值且2024≠0）",
        value=True,
    )

    valid_mask = df_filtered["2024"].notna() & df_filtered["2025"].notna() & (df_filtered["2024"] != 0)
    df_plot = df_filtered[valid_mask].copy() if only_valid else df_filtered.copy()

    # 指標
    col1, col2, col3 = st.columns(3)
    col1.metric("篩選後總筆數", f"{len(df_filtered):,}")
    col2.metric("可繪圖筆數", f"{len(df_plot):,}")
    if len(df_plot) > 0:
        col3.metric("2025業績(中位數)", f"{df_plot['2025'].median():,.2f}")
    else:
        col3.metric("2025業績(中位數)", "-")

    # 按鈕：切換商場名稱顯示
    if "show_labels" not in st.session_state:
        st.session_state["show_labels"] = False

    btn_col1, btn_col2 = st.columns([1, 5])
    with btn_col1:
        st.button(
            "顯示/隱藏商場名稱",
            on_click=toggle_show_labels,
            use_container_width=True,
        )
    with btn_col2:
        st.caption("提示：按下按鈕即可在散點圖上顯示每個點的商場名稱（資料多時可能較擁擠）。")

    # === 圖表 ===
    st.subheader("散點圖")
    if len(df_plot) == 0:
        st.info("目前篩選條件下沒有可繪圖的資料。")
    else:
        
        # tooltip 顯示所有欄位；數字千分位；成長率百分比
        tooltips = []
        for c in df_plot.columns:
            if c == "2025成長率":
                tooltips.append(alt.Tooltip("2025成長率:Q", format=".2%"))
            else:
                if pd.api.types.is_numeric_dtype(df_plot[c]):
                    tooltips.append(alt.Tooltip(f"{c}:Q", format=",.2f"))
                else:
                    tooltips.append(alt.Tooltip(f"{c}:N"))


        base = (
            alt.Chart(df_plot)
            .mark_circle(size=70)
            .encode(
                x=alt.X("2025成長率:Q", axis=alt.Axis(title="2025成長率", format="%")),
                y=alt.Y("2025:Q", axis=alt.Axis(title="2025業績")),
                color=alt.Color("商場名稱:N", legend=None),
                tooltip=tooltips,
            )
        )

        chart = base

        if st.session_state.get("show_labels", False):
            labels = (
                alt.Chart(df_plot)
                .mark_text(align="left", dx=6, dy=-6)
                .encode(
                    x=alt.X("2025成長率:Q"),
                    y=alt.Y("2025:Q"),
                    text=alt.Text("商場名稱:N"),
                )
            )
            chart = base + labels

        chart = chart.properties(height=520).configure_view(strokeWidth=0).configure(background="white")
        st.altair_chart(chart, use_container_width=True)

    # === 表格（成長率以百分比字串顯示）===
    st.subheader("資料表")
    show_cols = [c for c in ["類型", "體系", "商場名稱", "區", "縣市", "行政區", "地址", "2024", "2025", "2025成長率"] if c in df_filtered.columns]

    df_table = df_filtered[show_cols].copy()
    if "2025成長率" in df_table.columns:
        df_table["2025成長率"] = df_table["2025成長率"].apply(lambda v: to_percent_str(v, digits=2))

    st.dataframe(df_table, use_container_width=True, height=420)

    st.download_button(
        "下載篩選後資料（CSV）",
        data=df_table.to_csv(index=False).encode("utf-8-sig"),
        file_name="filtered_mall_sales.csv",
        mime="text/csv",
    )


if __name__ == "__main__":
    main()
