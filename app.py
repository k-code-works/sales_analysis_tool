
import io
import os
from datetime import date

import pandas as pd
import streamlit as st

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import japanize_matplotlib

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.chart import BarChart, DoughnutChart, LineChart
from openpyxl.chart.label import DataLabelList
from openpyxl.chart import Reference
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.image import Image as XLImage
from openpyxl.worksheet.page import PageMargins


# ==========================================
# 1. 基本設定
# ==========================================

st.set_page_config(
    page_title="売上分析ダッシュボード",
    page_icon="📊",
    layout="wide"
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CSV_FILE = os.path.join(BASE_DIR, "sales_app.csv")

COLUMNS = ["日付", "商品名", "数量", "単価"]

NAVY = "#17365D"
BLUE = "#1686B9"
TEAL = "#19A58B"
ORANGE = "#E3A343"


# ==========================================
# 2. 画面デザイン
# ==========================================

st.markdown("""
<style>
.stApp {
    background-color: #F4F7FB;
}
.block-container {
    padding-top: 4.5rem;
    padding-bottom: 2rem;
    max-width: 1450px;
}
[data-testid="stSidebar"] {
    background-color: #FFFFFF;
    border-right: 1px solid #E2E8F0;
}
.main-title {
    font-size: 29px;
    font-weight: 800;
    color: #17365D;
}
.sub-title {
    color: #718096;
    font-size: 13px;
    letter-spacing: 2px;
}
.kpi-card {
    background: white;
    border: 1px solid #E1E8F0;
    border-radius: 12px;
    padding: 19px;
    min-height: 115px;
    box-shadow: 0 3px 12px rgba(20,50,80,0.04);
}
.kpi-label {
    font-size: 13px;
    color: #64748B;
    font-weight: 600;
}
.kpi-value {
    font-size: 28px;
    font-weight: 800;
    color: #17365D;
    margin-top: 8px;
}
.kpi-caption {
    font-size: 11px;
    color: #94A3B8;
}
.section-title {
    font-size: 17px;
    font-weight: 700;
    color: #17365D;
    border-bottom: 2px solid #1686B9;
    padding-bottom: 8px;
    margin-bottom: 12px;
}
</style>
""", unsafe_allow_html=True)


# ==========================================
# 3. CSV読み込み
# ==========================================

def load_data():

    if not os.path.exists(CSV_FILE):
        return pd.DataFrame(columns=COLUMNS)

    try:
        df = pd.read_csv(
            CSV_FILE,
            encoding="utf-8-sig"
        )

        if not all(c in df.columns for c in COLUMNS):
            raise ValueError("CSVの列名が正しくありません。")

        df = df[COLUMNS].copy()

        df["日付"] = pd.to_datetime(
            df["日付"],
            errors="coerce"
        )

        df["数量"] = pd.to_numeric(
            df["数量"],
            errors="coerce"
        )

        df["単価"] = pd.to_numeric(
            df["単価"],
            errors="coerce"
        )

        df["商品名"] = (
            df["商品名"]
            .astype("string")
            .str.strip()
        )

        df = df.dropna(subset=COLUMNS)

        df = df[
            (df["商品名"] != "")
            & (df["数量"] > 0)
            & (df["単価"] >= 0)
            & (df["数量"] % 1 == 0)
            & (df["単価"] % 1 == 0)
            & (df["数量"] * df["単価"] <= 10**12)
        ].copy()

        df["日付"] = df["日付"].dt.strftime("%Y-%m-%d")
        df["数量"] = df["数量"].astype(int)
        df["単価"] = df["単価"].astype(int)

        return df.reset_index(drop=True)

    except Exception as e:
        st.error(f"CSV読み込みエラー：{e}")
        return pd.DataFrame(columns=COLUMNS)


# ==========================================
# 4. CSV保存
# ==========================================

def save_data(df):

    temp_file = CSV_FILE + ".tmp"

    try:
        df[COLUMNS].to_csv(
            temp_file,
            index=False,
            encoding="utf-8-sig"
        )

        os.replace(temp_file, CSV_FILE)

    except Exception:
        if os.path.exists(temp_file):
            os.remove(temp_file)

        raise


# ==========================================
# 5. データ検証
# ==========================================

def validate_data(df):

    result = df[COLUMNS].copy()

    result["日付"] = pd.to_datetime(
        result["日付"],
        errors="coerce"
    )

    result["数量"] = pd.to_numeric(
        result["数量"],
        errors="coerce"
    )

    result["単価"] = pd.to_numeric(
        result["単価"],
        errors="coerce"
    )

    result["商品名"] = (
        result["商品名"]
        .astype("string")
        .str.strip()
    )

    if result.isna().any().any():
        raise ValueError("空欄または不正な値があります。")

    if (result["商品名"] == "").any():
        raise ValueError("商品名を入力してください。")

    if (result["数量"] <= 0).any():
        raise ValueError("数量は1以上にしてください。")

    if (result["単価"] < 0).any():
        raise ValueError("単価は0以上にしてください。")

    if (
        (result["数量"] % 1 != 0).any()
        or (result["単価"] % 1 != 0).any()
    ):
        raise ValueError("数量と単価は整数にしてください。")

    if (result["数量"] * result["単価"] > 10**12).any():
        raise ValueError("売上金額が上限を超えています。")

    result["日付"] = result["日付"].dt.strftime("%Y-%m-%d")
    result["数量"] = result["数量"].astype(int)
    result["単価"] = result["単価"].astype(int)

    return result.reset_index(drop=True)


# ==========================================
# 6. 売上集計
# ==========================================

def calculate_sales(df):

    detail = df.copy()

    detail["売上金額"] = (
        detail["数量"] * detail["単価"]
    )

    ranking = (
        detail.groupby(
            "商品名",
            as_index=False
        )["売上金額"]
        .sum()
        .sort_values(
            "売上金額",
            ascending=False
        )
        .reset_index(drop=True)
    )

    ranking.insert(
        0,
        "順位",
        range(1, len(ranking) + 1)
    )

    daily = (
        detail.groupby(
            "日付",
            as_index=False
        )["売上金額"]
        .sum()
        .sort_values("日付")
        .reset_index(drop=True)
    )

    total = int(detail["売上金額"].sum())
    count = len(detail)
    products = detail["商品名"].nunique()

    return detail, ranking, daily, total, count, products

# ==========================================
# 7. グラフ日本語フォント
# ==========================================

# WindowsではYu Gothic / Meiryo、Linux環境ではNoto Sans CJK JPを優先候補にする。
plt.rcParams["font.family"] = [
    "Noto Sans CJK JP",
    "Yu Gothic",
    "Meiryo",
    "MS Gothic",
    "sans-serif",
]

plt.rcParams["axes.unicode_minus"] = False


# ==========================================
# 8. 商品別ランキンググラフ
# ==========================================

def create_ranking_chart(ranking):
    """商品名を左、売上金額を棒の右に表示する横棒グラフ。"""
    fig, ax = plt.subplots(figsize=(8.6, 3.1), dpi=100)

    if ranking.empty:
        ax.text(0.5, 0.5, "売上データがありません", ha="center", va="center")
        ax.axis("off")
        return fig

    chart_data = ranking.head(10).sort_values("売上金額", ascending=True)
    y_pos = list(range(len(chart_data)))
    values = chart_data["売上金額"].astype(float).tolist()
    bars = ax.barh(y_pos, values, color=BLUE, height=0.58)
    ax.set_yticks(y_pos)
    ax.set_yticklabels(chart_data["商品名"].astype(str), fontsize=9)
    ax.invert_yaxis()
    ax.set_title("商品別売上ランキング", fontsize=13, fontweight="bold", pad=16)
    ax.set_xlabel("売上金額（円）", fontsize=9, labelpad=7)
    ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f"¥{v:,.0f}"))
    ax.grid(axis="x", linestyle="--", alpha=0.25)
    ax.set_axisbelow(True)
    ax.bar_label(bars, labels=[f"¥{int(v):,}" for v in values], padding=4, fontsize=8)
    max_value = max(values) if values else 1
    ax.set_xlim(0, max_value * 1.24 if max_value > 0 else 1)
    ax.tick_params(axis="y", length=0, pad=7)
    for spine in ("top", "right", "left"):
        ax.spines[spine].set_visible(False)
    fig.subplots_adjust(left=0.22, right=0.96, top=0.82, bottom=0.20)
    return fig




# ==========================================
# 9. 商品別構成比グラフ
# ==========================================

def create_doughnut_chart(ranking):
    """商品別の構成比を、凡例と円グラフが重ならない配置で表示する。"""
    fig, ax = plt.subplots(figsize=(8.6, 3.1), dpi=100)
    positive = ranking[ranking["売上金額"] > 0].copy()

    if positive.empty:
        ax.text(0.5, 0.5, "構成比を表示できる売上がありません", ha="center", va="center")
        ax.axis("off")
        return fig

    palette = [BLUE, TEAL, ORANGE, "#7B61A8", "#D66B79", "#6C9A8B", "#A5A5A5", "#D5B85A"]
    wedges, _ = ax.pie(
        positive["売上金額"].astype(float), labels=None,
        startangle=90, counterclock=False,
        colors=[palette[i % len(palette)] for i in range(len(positive))],
        wedgeprops={"width": 0.40, "edgecolor": "white", "linewidth": 1.2},
    )
    ax.set_title("商品別売上構成比", fontsize=13, fontweight="bold", pad=14)
    total_amount = float(positive["売上金額"].sum())
    legend_labels = [
        f"{name}　{amount / total_amount:.1%}"
        for name, amount in zip(positive["商品名"].astype(str), positive["売上金額"])
    ]
    ax.legend(
        wedges, legend_labels, title="商品名・構成比",
        loc="center left", bbox_to_anchor=(1.02, 0.5),
        frameon=False, fontsize=8, title_fontsize=9,
        borderaxespad=0.0, labelspacing=0.8
    )
    ax.set_aspect("equal")
    fig.subplots_adjust(left=0.08, right=0.72, top=0.82, bottom=0.08)
    return fig




# ==========================================
# 10. 日別売上推移グラフ
# ==========================================

def create_daily_chart(daily):
    """日別売上を1本の折れ線で表示し、日付ラベルの重なりを抑える。"""
    fig, ax = plt.subplots(figsize=(8.6, 3.1), dpi=100)

    if daily.empty:
        ax.text(0.5, 0.5, "売上データがありません", ha="center", va="center")
        ax.axis("off")
        return fig

    chart_data = daily.sort_values("日付").reset_index(drop=True)
    x_values = list(range(len(chart_data)))
    y_values = chart_data["売上金額"].astype(float).tolist()
    ax.plot(x_values, y_values, color=BLUE, marker="o", linewidth=2.2, markersize=5)

    tick_count = min(7, len(chart_data))
    if tick_count == 1:
        tick_positions = [0]
    else:
        tick_positions = sorted(set(round(i * (len(chart_data) - 1) / (tick_count - 1)) for i in range(tick_count)))
    tick_labels = [pd.to_datetime(chart_data.iloc[i]["日付"]).strftime("%m/%d") for i in tick_positions]
    ax.set_xticks(tick_positions)
    ax.set_xticklabels(tick_labels, fontsize=8)
    ax.set_title("日別売上推移", fontsize=13, fontweight="bold", pad=14)
    ax.set_xlabel("日付", fontsize=9, labelpad=7)
    ax.set_ylabel("売上金額（円）", fontsize=9, labelpad=7)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda value, _: f"¥{value:,.0f}"))
    ax.grid(axis="y", linestyle="--", alpha=0.25)
    ax.set_axisbelow(True)
    ax.margins(x=0.06, y=0.20)
    ax.tick_params(axis="both", labelsize=8)
    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)
    fig.subplots_adjust(left=0.13, right=0.97, top=0.82, bottom=0.22)
    return fig





# ==========================================
# 11. Excelレポート作成
# ==========================================

def create_excel(df):
    """見出し帯とグラフが重ならない、縦方向に整理したExcelレポートを作成する。"""
    detail, ranking, daily, total, count, products = calculate_sales(df)
    wb = Workbook()
    ws = wb.active
    ws.title = "売上分析レポート"
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 85

    navy, blue, white = "17365D", "1686B9", "FFFFFF"
    pale_blue, text_gray = "DDEBF7", "64748B"
    border_color = "D9E2F3"
    thin_border = Border(
        left=Side(style="thin", color=border_color),
        right=Side(style="thin", color=border_color),
        top=Side(style="thin", color=border_color),
        bottom=Side(style="thin", color=border_color),
    )

    # E列は売上明細の「単価」列として使用するため、狭い空白列にしない。
    # 日付・商品名・数量・単価・売上金額がすべて読める幅を確保する。
    widths = {"A": 3, "B": 15, "C": 24, "D": 12, "E": 15,
              "F": 16, "G": 13, "H": 13, "I": 13, "J": 13,
              "K": 13, "L": 13, "M": 13}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    def fill_range(cell_range, color):
        for row in ws[cell_range]:
            for cell in row:
                cell.fill = PatternFill("solid", fgColor=color)

    def section_title(row_num, title):
        cell_range = f"B{row_num}:M{row_num}"
        ws.merge_cells(cell_range)
        cell = ws[f"B{row_num}"]
        cell.value = title
        cell.font = Font(name="Meiryo", size=12, bold=True, color=white)
        cell.alignment = Alignment(horizontal="left", vertical="center")
        fill_range(cell_range, navy)
        ws.row_dimensions[row_num].height = 25

    def table_header(row_num, start_col, headers):
        for col, label in enumerate(headers, start=start_col):
            cell = ws.cell(row_num, col, label)
            cell.font = Font(name="Meiryo", size=10, bold=True, color=navy)
            cell.fill = PatternFill("solid", fgColor=pale_blue)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
        ws.row_dimensions[row_num].height = 23

    # タイトル帯
    ws.merge_cells("B2:M2")
    ws["B2"] = "売上分析レポート"
    ws["B2"].font = Font(name="Meiryo", size=20, bold=True, color=white)
    ws["B2"].alignment = Alignment(horizontal="center", vertical="center")
    fill_range("B2:M2", navy)
    ws.row_dimensions[2].height = 42
    ws.merge_cells("B3:M3")
    ws["B3"] = "SALES PERFORMANCE REPORT　｜　作成日：" + date.today().strftime("%Y年%m月%d日")
    ws["B3"].font = Font(name="Meiryo", size=9, color=text_gray)
    ws["B3"].alignment = Alignment(horizontal="right", vertical="center")
    ws.row_dimensions[3].height = 22

    # KPIカード
    section_title(5, "売上サマリー")
    summary = [
        ("B7:D7", "B8:D9", "売上合計", total, '¥#,##0'),
        ("F7:H7", "F8:H9", "販売件数", count, '#,##0"件"'),
        ("J7:M7", "J8:M9", "商品種類", products, '#,##0"種類"'),
    ]
    for label_range, value_range, label, value, number_format in summary:
        ws.merge_cells(label_range)
        ws.merge_cells(value_range)
        label_cell = ws[label_range.split(":")[0]]
        value_cell = ws[value_range.split(":")[0]]
        label_cell.value = label
        label_cell.font = Font(name="Meiryo", size=10, bold=True, color=navy)
        label_cell.alignment = Alignment(horizontal="center", vertical="center")
        value_cell.value = int(value)
        value_cell.number_format = number_format
        value_cell.font = Font(name="Meiryo", size=18, bold=True, color=blue)
        value_cell.alignment = Alignment(horizontal="center", vertical="center")
        fill_range(label_range, pale_blue)
        fill_range(value_range, "F7FAFD")
        for area in (label_range, value_range):
            for row in ws[area]:
                for cell in row:
                    cell.border = thin_border
    ws.row_dimensions[7].height = 24
    ws.row_dimensions[8].height = 26
    ws.row_dimensions[9].height = 22

    # 商品ランキング（上位10件に限定し、左の表と右のグラフを分離）
    section_title(11, "商品別売上ランキング（上位10件）")
    table_header(12, 2, ["順位", "商品名", "売上金額"])
    top_ranking = ranking.head(10)
    for i, (_, item) in enumerate(top_ranking.iterrows(), start=1):
        row_num = 12 + i
        for col, value in enumerate([i, str(item["商品名"]), int(item["売上金額"])], start=2):
            cell = ws.cell(row_num, col, value)
            cell.font = Font(name="Meiryo", size=9, color="243B53")
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="left" if col == 3 else ("right" if col == 4 else "center"), vertical="center")
            if col == 4:
                cell.number_format = '¥#,##0'
            if i % 2 == 0:
                cell.fill = PatternFill("solid", fgColor="F7FAFC")
        ws.row_dimensions[row_num].height = 21

    def add_chart_image(fig, anchor):
        buffer = io.BytesIO()
        fig.savefig(buffer, format="png", dpi=120, facecolor="white")
        plt.close(fig)
        buffer.seek(0)
        image = XLImage(buffer)
        image.width = 680
        image.height = 250
        ws.add_image(image, anchor)

    if not ranking.empty:
        add_chart_image(create_ranking_chart(ranking), "F13")

    # グラフ間に独立した見出し行と余白を確保
    section_title(31, "商品別売上構成比")
    if not ranking.empty:
        add_chart_image(create_doughnut_chart(ranking), "F33")

    section_title(51, "日別売上推移")
    daily_table_header = 53
    table_header(daily_table_header, 2, ["日付", "売上金額"])
    for i, (_, item) in enumerate(daily.iterrows(), start=1):
        row_num = daily_table_header + i
        day_cell = ws.cell(row_num, 2, pd.to_datetime(item["日付"]).to_pydatetime())
        day_cell.number_format = "yyyy-mm-dd"
        amount_cell = ws.cell(row_num, 3, int(item["売上金額"]))
        amount_cell.number_format = '¥#,##0'
        for col in (2, 3):
            cell = ws.cell(row_num, col)
            cell.font = Font(name="Meiryo", size=9, color="243B53")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
            if i % 2 == 0:
                cell.fill = PatternFill("solid", fgColor="F7FAFC")
        ws.row_dimensions[row_num].height = 21

    if len(daily) >= 1:
        add_chart_image(create_daily_chart(daily), "F53")

    # 売上明細はグラフ画像の下に配置。日付列は十分な幅を確保。
    detail_start = max(70, 53 + len(daily) + 3)
    section_title(detail_start, "売上明細一覧")
    header_row = detail_start + 1
    data_start = header_row + 1
    table_header(header_row, 2, ["日付", "商品名", "数量", "単価", "売上金額"])

    for i, (_, item) in enumerate(detail.iterrows()):
        row_num = data_start + i
        values = [pd.to_datetime(item["日付"]).to_pydatetime(), str(item["商品名"]), int(item["数量"]), int(item["単価"]), int(item["売上金額"])]
        for col, value in enumerate(values, start=2):
            cell = ws.cell(row_num, col, value)
            cell.font = Font(name="Meiryo", size=9, color="243B53")
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="left" if col == 3 else "center", vertical="center")
            if col == 2:
                cell.number_format = "yyyy-mm-dd"
            elif col in (5, 6):
                cell.number_format = '¥#,##0'
            if i % 2 == 1:
                cell.fill = PatternFill("solid", fgColor="F7FAFC")
        ws.row_dimensions[row_num].height = 21

    total_row = data_start + len(detail)
    # 合計ラベルはB～E列を結合して表示し、狭いセルで文字が切れないようにする。
    ws.merge_cells(
        start_row=total_row, start_column=2,
        end_row=total_row, end_column=5
    )
    ws.cell(total_row, 2, "総売上")
    ws.cell(total_row, 6, int(total))
    ws.cell(total_row, 6).number_format = '¥#,##0'
    for col in range(2, 7):
        cell = ws.cell(total_row, col)
        cell.fill = PatternFill("solid", fgColor=pale_blue)
        cell.font = Font(name="Meiryo", size=10, bold=True, color=navy)
        cell.border = thin_border
        cell.alignment = Alignment(
            horizontal="right" if col in (2, 6) else "center",
            vertical="center"
        )

    ws.auto_filter.ref = f"B{header_row}:F{total_row - 1}" if len(detail) else f"B{header_row}:F{header_row}"
    ws.freeze_panes = None
    ws.print_area = f"B2:M{total_row}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.25, right=0.25, top=0.45, bottom=0.45, header=0.15, footer=0.15)
    ws.oddFooter.center.text = "売上分析レポート　｜　&P / &N"
    ws.oddFooter.center.size = 8
    ws.oddFooter.center.font = "Meiryo"

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()



# ==========================================
# 12. アプリ共通ヘッダー
# ==========================================

st.markdown(
    '<div class="main-title">売上分析ダッシュボード</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="sub-title">SALES PERFORMANCE ANALYTICS</div>',
    unsafe_allow_html=True
)

st.divider()


# ==========================================
# 13. データ読み込み・メニュー
# ==========================================

if "sales_data" not in st.session_state:

    st.session_state.sales_data = load_data()

df = st.session_state.sales_data.copy()

menu = st.sidebar.radio(
    "メニュー",
    [
        "ダッシュボード",
        "売上データ登録",
        "データ管理",
        "売上分析",
        "Excelレポート"
    ]
)

st.sidebar.divider()

st.sidebar.caption("売上管理システム")
st.sidebar.caption("Python / Streamlit")


# ==========================================
# 14. ダッシュボード
# ==========================================

if menu == "ダッシュボード":

    detail, ranking, daily, total, count, products = calculate_sales(df)

    st.subheader("売上サマリー")

    col1, col2, col3 = st.columns(3)

    with col1:

        st.markdown(
            f"""
            <div class="kpi-card">
                <div class="kpi-label">売上合計</div>
                <div class="kpi-value">¥{total:,}</div>
                <div class="kpi-caption">総売上金額</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with col2:

        st.markdown(
            f"""
            <div class="kpi-card">
                <div class="kpi-label">販売件数</div>
                <div class="kpi-value">{count:,}件</div>
                <div class="kpi-caption">登録データ件数</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with col3:

        st.markdown(
            f"""
            <div class="kpi-card">
                <div class="kpi-label">商品種類</div>
                <div class="kpi-value">{products:,}種類</div>
                <div class="kpi-caption">取扱商品数</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    st.write("")

    left, right = st.columns([1.3, 1])

    # 商品別ランキング

    with left:

        st.markdown(
            '<div class="section-title">商品別売上ランキング</div>',
            unsafe_allow_html=True
        )

        if not ranking.empty:

            st.dataframe(
                ranking,
                hide_index=True,
                use_container_width=True
            )

            fig = create_ranking_chart(ranking)

            st.pyplot(
                fig,
                use_container_width=True
            )

            plt.close(fig)

        else:

            st.info("売上データがありません。")

    # 商品別構成比

    with right:

        st.markdown(
            '<div class="section-title">商品別売上構成比</div>',
            unsafe_allow_html=True
        )

        if not ranking.empty:

            fig = create_doughnut_chart(ranking)

            st.pyplot(
                fig,
                use_container_width=True
            )

            plt.close(fig)

        else:

            st.info("売上データがありません。")

    st.divider()

    # 日別売上推移

    st.markdown(
        '<div class="section-title">日別売上推移</div>',
        unsafe_allow_html=True
    )

    if len(daily) >= 2:

        fig = create_daily_chart(daily)

        st.pyplot(
            fig,
            use_container_width=True
        )

        plt.close(fig)

    else:

        st.info(
            "日別売上推移を表示するには、"
            "異なる日付の売上データが2日分以上必要です。"
        )

    st.divider()

    # 最新の売上明細

    st.markdown(
        '<div class="section-title">最新の売上明細</div>',
        unsafe_allow_html=True
    )

    if not detail.empty:

        st.dataframe(
            detail.sort_values(
                "日付",
                ascending=False
            ),
            hide_index=True,
            use_container_width=True
        )

    else:

        st.info("売上データがありません。")


# ==========================================
# 15. 売上データ登録
# ==========================================

elif menu == "売上データ登録":

    st.subheader("売上データ登録")

    st.caption(
        "販売日・商品名・数量・単価を入力してください。"
    )

    with st.form("sales_form"):

        input_date = st.date_input(
            "販売日",
            value=date.today()
        )

        product = st.text_input(
            "商品名",
            placeholder="例：商品A"
        )

        col1, col2 = st.columns(2)

        with col1:

            quantity = st.number_input(
                "数量",
                min_value=1,
                max_value=1000000,
                value=1,
                step=1
            )

        with col2:

            price = st.number_input(
                "単価（円）",
                min_value=0,
                max_value=100000000,
                value=100,
                step=100
            )

        st.metric(
            "売上金額（予定）",
            f"¥{int(quantity * price):,}"
        )

        submitted = st.form_submit_button(
            "売上データを登録",
            use_container_width=True
        )

        if submitted:

            if not product.strip():

                st.error("商品名を入力してください。")

            elif quantity * price > 10**12:

                st.error("売上金額が上限を超えています。")

            else:

                new_row = pd.DataFrame([{
                    "日付": input_date.isoformat(),
                    "商品名": product.strip(),
                    "数量": int(quantity),
                    "単価": int(price)
                }])

                updated = pd.concat(
                    [df, new_row],
                    ignore_index=True
                )

                try:

                    updated = validate_data(updated)

                    save_data(updated)

                    st.session_state.sales_data = updated

                    st.success("売上データを登録しました。")

                    st.rerun()

                except Exception as e:

                    st.error(f"登録エラー：{e}")


# ==========================================
# 16. データ管理
# ==========================================

elif menu == "データ管理":

    st.subheader("売上データ管理")

    if df.empty:

        st.info("管理する売上データがありません。")

    else:

        st.caption(
            "表のセルを直接編集して保存できます。"
        )

        edited = st.data_editor(
            df,
            num_rows="fixed",
            use_container_width=True,
            hide_index=True,
            key="sales_editor"
        )

        # 変更保存

        if st.button(
            "変更を保存",
            use_container_width=True
        ):

            try:

                updated = validate_data(edited)

                save_data(updated)

                st.session_state.sales_data = updated

                st.success("変更を保存しました。")

                st.rerun()

            except Exception as e:

                st.error(f"保存できませんでした：{e}")

        st.divider()

        # データ削除

        st.subheader("売上データの削除")

        selected = st.multiselect(
            "削除するデータを選択してください。",
            options=list(range(len(df))),
            format_func=lambda x: (
                f"{x + 1}：{df.iloc[x]['商品名']} "
                f"（{df.iloc[x]['日付']}）"
            )
        )

        if st.button(
            "選択したデータを削除",
            type="secondary",
            use_container_width=True
        ):

            if not selected:

                st.warning("削除するデータを選択してください。")

            else:

                updated = df.drop(
                    index=selected
                ).reset_index(drop=True)

                try:

                    save_data(updated)

                    st.session_state.sales_data = updated

                    st.success("データを削除しました。")

                    st.rerun()

                except Exception as e:

                    st.error(f"削除エラー：{e}")


# ==========================================
# 17. 売上分析
# ==========================================

elif menu == "売上分析":

    st.subheader("売上分析")

    detail, ranking, daily, total, count, products = calculate_sales(df)

    if detail.empty:

        st.info("分析するデータがありません。")

    else:

        # 商品別ランキング

        st.markdown(
            '<div class="section-title">商品別売上ランキング</div>',
            unsafe_allow_html=True
        )

        st.dataframe(
            ranking,
            hide_index=True,
            use_container_width=True
        )

        fig = create_ranking_chart(ranking)

        st.pyplot(
            fig,
            use_container_width=True
        )

        plt.close(fig)

        st.divider()

        # 商品別構成比

        st.markdown(
            '<div class="section-title">商品別売上構成比</div>',
            unsafe_allow_html=True
        )

        fig = create_doughnut_chart(ranking)

        st.pyplot(
            fig,
            use_container_width=True
        )

        plt.close(fig)

        st.divider()

        # 日別売上推移

        st.markdown(
            '<div class="section-title">日別売上推移</div>',
            unsafe_allow_html=True
        )

        if len(daily) >= 2:

            fig = create_daily_chart(daily)

            st.pyplot(
                fig,
                use_container_width=True
            )

            plt.close(fig)

        else:

            st.info(
                "異なる日付の売上データが2日分以上必要です。"
            )

        st.divider()

        # 売上明細

        st.markdown(
            '<div class="section-title">売上明細</div>',
            unsafe_allow_html=True
        )

        st.dataframe(
            detail,
            hide_index=True,
            use_container_width=True
        )


# ==========================================
# 18. Excelレポート
# ==========================================

elif menu == "Excelレポート":

    st.subheader("Excelレポート出力")

    st.write(
        "売上サマリー・商品別ランキング・"
        "3種類のグラフと売上明細をまとめたExcelレポートを作成します。"
    )

    if df.empty:

        st.warning(
            "Excelレポートを作成するには、"
            "売上データを登録してください。"
        )

    else:

        detail, ranking, daily, total, count, products = calculate_sales(df)

        col1, col2, col3 = st.columns(3)

        col1.metric(
            "売上合計",
            f"¥{total:,}"
        )

        col2.metric(
            "販売件数",
            f"{count:,}件"
        )

        col3.metric(
            "商品種類",
            f"{products:,}種類"
        )

        st.divider()

        st.markdown(
            '<div class="section-title">出力内容</div>',
            unsafe_allow_html=True
        )

        st.markdown("""
        - 売上合計・販売件数・商品種類
        - 商品別売上ランキング
        - 商品別売上構成比（ドーナツグラフ）
        - 日別売上推移（折れ線グラフ）
        - 売上明細一覧
        """)

        try:

            excel_data = create_excel(df)

            st.download_button(
                label="📥 Excelレポートをダウンロード",
                data=excel_data,
                file_name="売上分析レポート.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                use_container_width=True
            )

        except Exception as e:

            st.error(f"Excel作成エラー：{e}")


# ==========================================
# 19. フッター
# ==========================================

st.divider()

st.caption(
    "売上分析ダッシュボード "
    "| Python / Streamlit / Pandas / Matplotlib / OpenPyXL"
)
