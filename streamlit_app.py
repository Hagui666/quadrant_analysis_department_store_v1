# streamlit_app.py
# -*- coding: utf-8 -*-
"""
Streamlit 互動式散點圖（黑底介面、圖表白底）：
- 預設讀取專案內建 Excel：商場年業績表_含成長率.xlsx（與本檔案同資料夾）
- 亦可在側欄上傳新的 Excel 以覆蓋內建資料
- 讀取工作表：篩選結果
- 計算「2025成長率」 = (2025 - 2024) / 2024
- 依層級篩選（高層影響低層選項）：類型 -> 區 -> 縣市
- 散點圖：X=2025成長率（百分比），Y=2025（y軸標籤：2025業績）
- 圖表：白底、無格線、座標軸文字黑色
- 顏色：以「商場名稱」分類（不同名稱不同顏色）
- Tooltip：顯示該點所有欄位；數字欄位千分位+2位小數；成長率以百分比
- 表格：成長率欄位以百分比字串呈現
"""

from __future__ import annotations

from pathlib import Path
import pandas as pd
import streamlit as st
import altair as alt

DEFAULT_FILE = "商場年業績表_含成長率.xlsx"
DEFAULT_SHEET = "篩選結果"


def _inject_dark_theme():
    """黑底介面（圖表底色另外在 Altair 設白底）"""
    st.markdown(
        """
        <style>
        .stApp { background-color: #0e1117; color: #ffffff; }
        section[data-testid="stSidebar"] { background-color: #0e1117; color: #ffffff; }
        section[data-testid="stSidebar"] * { color: #ffffff !important; }

        /* 普通文字 */
        .stMarkdown, .stText, .stCaption, .stMetric, .stTable { color: #ffffff; }

        /* dataframe */
        div[data-testid="stDataFrame"] * { color: #ffffff; }

        /* 小地方：讓 metric 標籤不要太灰 */
        [data-testid="stMetricLabel"] { color: #ffffff; }
        </style>
        """,
        unsafe_allow_html=True,
    )


def get_builtin_excel_path() -> Path:
    """回傳 repo 內建 Excel 的絕對路徑（與本檔案同資料夾）"""
    return Path(__file__).resolve().parent / DEFAULT_FILE


@st.cache_data(show_spinner=False)
def load_data(file_ref, sheet: str) -> pd.DataFrame:
    # 更友善的錯誤：Streamlit Cloud 常見是沒裝 openpyxl
    try:
        import openpyxl  # noqa: F401
    except Exception as e:
        raise RuntimeError(
            "缺少 openpyxl，無法讀取 .xlsx。請在 requirements.txt 加上 openpyxl 後重新部署。"
        ) from e

    df = pd.read_excel(file_ref, sheet_name=sheet)
    df.columns = [str(c).strip() for c in df.columns]

    for col in ["2024", "2025"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # 成長率：避免除以 0
    if "2024" in df.columns and "2025" in df.columns:
        df["2025成長率"] = (df["2025"] - df["2024"]) / df["2024"]
        df.loc[df["2024"] == 0, "2025成長率"] = pd.NA

    return df


def to_percent_str(x, digits: int = 2) -> str:
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

    # === 資料來源：預設用內建檔案，亦可上傳覆蓋 ===
    st.sidebar.header("資料來源")
    builtin_path = get_builtin_excel_path()

    uploaded = st.sidebar.file_uploader(
        "如需覆蓋內建資料，請上傳新的 Excel 檔（可選）",
        type=["xlsx"],
        help="未上傳時，會使用 GitHub 專案內建的 Excel 檔案。",
    )

    if uploaded is not None:
        file_ref = uploaded
        file_name_for_info = uploaded.name
    else:
        if not builtin_path.exists():
            st.error(
                f"找不到內建資料檔：{builtin_path}\n\n"
                "請確認已把 Excel 檔案一併推到 GitHub repo，且與 streamlit_app.py 在同一資料夾；\n"
                "或改用側欄上傳 Excel 檔案。"
            )
            st.stop()
        file_ref = str(builtin_path)
        file_name_for_info = builtin_path.name

    sheet = st.sidebar.text_input("工作表名稱", value=DEFAULT_SHEET)

    # === 讀取 ===
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

    st.caption(f"目前資料：{file_name_for_info}｜工作表：{sheet}｜總筆數：{len(df):,}")

    # === 層級篩選（類型 -> 區 -> 縣市）===
    st.sidebar.header("篩選條件（高層會影響低層選項）")

    type_options = sorted(df["類型"].dropna().unique().tolist())
    selected_types = st.sidebar.multiselect("類型", options=type_options, default=type_options)
    df_f1 = df[df["類型"].isin(selected_types)] if selected_types else df.iloc[0:0]

    area_options = sorted(df_f1["區"].dropna().unique().tolist())
    selected_areas = st.sidebar.multiselect("區", options=area_options, default=area_options)
    df_f2 = df_f1[df_f1["區"].isin(selected_areas)] if selected_areas else df_f1.iloc[0:0]

    city_options = sorted(df_f2["縣市"].dropna().unique().tolist())
    selected_cities = st.sidebar.multiselect("縣市", options=city_options, default=city_options)
    df_filtered = df_f2[df_f2["縣市"].isin(selected_cities)] if selected_cities else df_f2.iloc[0:0]

    st.sidebar.header("顯示選項")
    only_valid = st.sidebar.checkbox(
        "只顯示可計算正常成長率（2024/2025皆有值且2024≠0）",
        value=True,
    )

    valid_mask = df_filtered["2024"].notna() & df_filtered["2025"].notna() & (df_filtered["2024"] != 0)
    df_plot = df_filtered[valid_mask].copy() if only_valid else df_filtered.copy()

    # === 指標 ===
    c1, c2, c3 = st.columns(3)
    c1.metric("篩選後總筆數", f"{len(df_filtered):,}")
    c2.metric("可繪圖筆數", f"{len(df_plot):,}")
    c3.metric("2025業績(中位數)", f"{df_plot['2025'].median():,.2f}" if len(df_plot) else "-")

    # === 顯示/隱藏標籤 ===
    if "show_labels" not in st.session_state:
        st.session_state["show_labels"] = False

    def _toggle_labels():
        st.session_state["show_labels"] = not st.session_state["show_labels"]

    btn_col1, btn_col2 = st.columns([1, 5])
    with btn_col1:
        st.button("顯示/隱藏商場名稱", on_click=_toggle_labels, use_container_width=True)
    with btn_col2:
        st.caption("資料多時顯示名稱會較擁擠，可用此按鈕切換。")

    # === 散點圖 ===
    st.subheader("散點圖")
    if len(df_plot) == 0:
        st.info("目前篩選條件下沒有可繪圖的資料。")
    else:
        # tooltip：所有欄位
        tooltips = []
        for col in df_plot.columns:
            if col == "2025成長率":
                tooltips.append(alt.Tooltip("2025成長率:Q", format=".2%"))
            else:
                if pd.api.types.is_numeric_dtype(df_plot[col]):
                    tooltips.append(alt.Tooltip(f"{col}:Q", format=",.2f"))
                else:
                    tooltips.append(alt.Tooltip(f"{col}:N"))

        base = (
            alt.Chart(df_plot)
            .mark_circle(size=70)
            .encode(
                x=alt.X(
                    "2025成長率:Q",
                    axis=alt.Axis(title="2025成長率", format="%", grid=False, labelColor="black", titleColor="black"),
                ),
                y=alt.Y(
                    "2025:Q",
                    axis=alt.Axis(title="2025業績", grid=False, labelColor="black", titleColor="black"),
                ),
                # 顏色以商場名稱分類
                color=alt.Color("商場名稱:N", legend=None),
                tooltip=tooltips,
            )
        )

        chart = base

        if st.session_state["show_labels"]:
            labels = (
                alt.Chart(df_plot)
                .mark_text(align="left", dx=6, dy=-6, color="black")
                .encode(x="2025成長率:Q", y="2025:Q", text="商場名稱:N")
            )
            chart = base + labels

        # 圖表：白底、無格線、軸線與刻度黑色
        chart = (
            chart.properties(height=520)
            .configure_view(strokeWidth=0)
            .configure(background="white")
            .configure_axis(
                grid=False,
                labelColor="black",
                titleColor="black",
                tickColor="black",
                domainColor="black",
            )
        )

        st.altair_chart(chart, use_container_width=True)

    # === 表格（成長率百分比）===
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
