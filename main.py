"""
売上分析システム - Excelレポート作成

app.py と同じ sales_app.csv を読み込み、
ダッシュボード形式のExcelレポートを作成します。
"""

import csv
import io
from collections import defaultdict
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, DoughnutChart, LineChart
from openpyxl.chart import Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.worksheet.page import PageMargins
from openpyxl.drawing.image import Image as XLImage

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import japanize_matplotlib
from matplotlib.ticker import FuncFormatter

plt.rcParams["font.family"] = ["Noto Sans CJK JP", "Meiryo", "Yu Gothic", "MS Gothic", "sans-serif"]
plt.rcParams["axes.unicode_minus"] = False


# ==================================
# 1. ファイル・デザイン設定
# ==================================

BASE_DIR = Path(__file__).resolve().parent
CSV_FILE = BASE_DIR / "sales_app.csv"
EXCEL_FILE = BASE_DIR / "売上分析レポート.xlsx"

NAVY = "17365D"
BLUE = "1686B9"
TEAL = "19A58B"
PALE_BLUE = "DDEBF7"
PALE_GRAY = "F4F7FB"
WHITE = "FFFFFF"
TEXT = "243B53"
BORDER_COLOR = "D9E2F3"

FONT_NAME = "Meiryo"


# ==================================
# 2. CSV読み込み
# ==================================

def load_sales_data():
    """共通CSVファイルを読み込み、入力内容を確認する。"""
    sales_data = []
    error_count = 0
    required_columns = ["日付", "商品名", "数量", "単価"]

    try:
        with CSV_FILE.open("r", encoding="utf-8-sig", newline="") as file:
            reader = csv.DictReader(file)

            if not reader.fieldnames or not all(
                column in reader.fieldnames for column in required_columns
            ):
                raise ValueError(
                    "CSVの列名が正しくありません。"
                    "日付、商品名、数量、単価が必要です。"
                )

            for row_number, row in enumerate(reader, start=2):
                try:
                    sale_date = datetime.strptime(
                        row["日付"].strip(), "%Y-%m-%d"
                    ).date()
                    product = row["商品名"].strip()
                    quantity = int(row["数量"])
                    price = int(row["単価"])

                    if not product or quantity <= 0 or price < 0:
                        raise ValueError

                    amount = quantity * price
                    if amount > 10**12:
                        raise ValueError

                    sales_data.append({
                        "date": sale_date,
                        "product": product,
                        "quantity": quantity,
                        "price": price,
                        "amount": amount,
                    })

                except (ValueError, TypeError, AttributeError):
                    print(f"警告：{row_number}行目のデータを確認してください。")
                    error_count += 1

    except FileNotFoundError:
        raise FileNotFoundError(
            f"{CSV_FILE.name} が見つかりません。"
            "app.pyと同じフォルダに置いてください。"
        )

    return sales_data, error_count


# ==================================
# 3. 売上集計
# ==================================

def calculate_sales(sales_data):
    """商品別・日別の売上金額を集計する。"""
    product_sales = defaultdict(int)
    daily_sales = defaultdict(int)

    for data in sales_data:
        product_sales[data["product"]] += data["amount"]
        daily_sales[data["date"]] += data["amount"]

    ranking = sorted(
        product_sales.items(),
        key=lambda item: item[1],
        reverse=True
    )
    daily_ranking = sorted(daily_sales.items())
    total_sales = sum(data["amount"] for data in sales_data)

    return ranking, daily_ranking, total_sales, len(product_sales)


# ==================================
# 3-1. Excel用グラフ画像
# ==================================

def figure_to_excel_image(fig):
    """MatplotlibのグラフをExcelに貼れる画像へ変換する。"""
    buffer = io.BytesIO()
    fig.savefig(buffer, format="png", dpi=120, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    buffer.seek(0)
    image = XLImage(buffer)
    image.width = 680
    image.height = 250
    return image


def create_ranking_image(ranking):
    """商品名を左、売上金額を棒の右に表示する横棒グラフ。"""
    fig, ax = plt.subplots(figsize=(8.6, 3.1), dpi=100)
    if not ranking:
        ax.text(0.5, 0.5, "売上データがありません", ha="center", va="center")
        ax.axis("off")
        return figure_to_excel_image(fig)
    chart_data = ranking[:10][::-1]
    products = [item[0] for item in chart_data]
    amounts = [item[1] for item in chart_data]
    y = list(range(len(products)))
    bars = ax.barh(y, amounts, color="#1686B9", height=0.58)
    ax.set_yticks(y, labels=products)
    ax.invert_yaxis()
    ax.set_title("商品別売上ランキング", fontsize=13, fontweight="bold", pad=16)
    ax.set_xlabel("売上金額（円）", fontsize=9, labelpad=7)
    ax.xaxis.set_major_formatter(FuncFormatter(lambda value, _: f"¥{value:,.0f}"))
    ax.grid(axis="x", linestyle="--", alpha=0.25)
    ax.set_axisbelow(True)
    ax.bar_label(bars, labels=[f"¥{v:,}" for v in amounts], padding=4, fontsize=8)
    ax.set_xlim(0, max(amounts or [1]) * 1.24 if max(amounts or [1]) > 0 else 1)
    ax.tick_params(axis="y", labelsize=9, length=0, pad=7)
    for spine in ("top", "right", "left"):
        ax.spines[spine].set_visible(False)
    fig.subplots_adjust(left=0.22, right=0.96, top=0.82, bottom=0.20)
    return figure_to_excel_image(fig)




def create_doughnut_image(ranking):
    """円グラフと凡例を左右に分離し、文字の重なりを防ぐ。"""
    fig, ax = plt.subplots(figsize=(8.6, 3.1), dpi=100)
    positive = [(name, amount) for name, amount in ranking if amount > 0]
    if not positive:
        ax.text(0.5, 0.5, "構成比を表示できる売上がありません", ha="center", va="center")
        ax.axis("off")
        return figure_to_excel_image(fig)
    palette = ["#1686B9", "#19A58B", "#E3A343", "#7B61A8", "#D66B79", "#6C9A8B", "#A5A5A5", "#D5B85A"]
    wedges, _ = ax.pie(
        [amount for _, amount in positive], labels=None,
        startangle=90, counterclock=False,
        colors=[palette[i % len(palette)] for i in range(len(positive))],
        wedgeprops={"width": 0.40, "edgecolor": "white", "linewidth": 1.2}
    )
    ax.set_title("商品別売上構成比", fontsize=13, fontweight="bold", pad=14)
    total_amount = sum(amount for _, amount in positive)
    legend_labels = [f"{name}　{amount / total_amount:.1%}" for name, amount in positive]
    ax.legend(wedges, legend_labels, title="商品名・構成比",
              loc="center left", bbox_to_anchor=(1.02, 0.5), frameon=False,
              fontsize=8, title_fontsize=9, borderaxespad=0.0, labelspacing=0.8)
    ax.set_aspect("equal")
    fig.subplots_adjust(left=0.08, right=0.72, top=0.82, bottom=0.08)
    return figure_to_excel_image(fig)




def create_daily_image(daily_ranking):
    """日付を横軸、売上金額を1系列で表示する折れ線グラフ。"""
    fig, ax = plt.subplots(figsize=(8.6, 3.1), dpi=100)
    if not daily_ranking:
        ax.text(0.5, 0.5, "売上データがありません", ha="center", va="center")
        ax.axis("off")
        return figure_to_excel_image(fig)
    dates = [day.strftime("%m/%d") for day, _ in daily_ranking]
    values = [amount for _, amount in daily_ranking]
    x = list(range(len(dates)))
    ax.plot(x, values, color="#1686B9", marker="o", linewidth=2.2, markersize=5)
    tick_count = min(7, len(dates))
    if tick_count == 1:
        positions = [0]
    else:
        positions = sorted(set(round(i * (len(dates) - 1) / (tick_count - 1)) for i in range(tick_count)))
    ax.set_xticks(positions, labels=[dates[i] for i in positions])
    ax.set_title("日別売上推移", fontsize=13, fontweight="bold", pad=14)
    ax.set_xlabel("日付", fontsize=9, labelpad=7)
    ax.set_ylabel("売上金額（円）", fontsize=9, labelpad=7)
    ax.yaxis.set_major_formatter(FuncFormatter(lambda value, _: f"¥{value:,.0f}"))
    ax.grid(axis="y", linestyle="--", alpha=0.25)
    ax.set_axisbelow(True)
    ax.margins(x=0.06, y=0.20)
    ax.tick_params(axis="both", labelsize=8)
    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)
    fig.subplots_adjust(left=0.13, right=0.97, top=0.82, bottom=0.22)
    return figure_to_excel_image(fig)




# ==================================
# 4. Excelの共通デザイン設定
# ==================================

def set_cell_style(cell, bold=False, color=TEXT, fill=None,
                   horizontal="center", number_format=None):
    """セルの文字・背景・配置・罫線を設定する。"""
    cell.font = Font(
        name=FONT_NAME,
        size=10,
        bold=bold,
        color=color
    )
    cell.alignment = Alignment(
        horizontal=horizontal,
        vertical="center"
    )
    cell.border = Border(
        left=Side(style="thin", color=BORDER_COLOR),
        right=Side(style="thin", color=BORDER_COLOR),
        top=Side(style="thin", color=BORDER_COLOR),
        bottom=Side(style="thin", color=BORDER_COLOR)
    )

    if fill:
        cell.fill = PatternFill("solid", fgColor=fill)

    if number_format:
        cell.number_format = number_format


def add_section_title(sheet, cell_range, title):
    """濃紺の見出し行を作る。"""
    sheet.merge_cells(cell_range)
    cell = sheet[cell_range.split(":")[0]]
    cell.value = title
    cell.font = Font(name=FONT_NAME, size=12, bold=True, color=WHITE)
    cell.fill = PatternFill("solid", fgColor=NAVY)
    cell.alignment = Alignment(horizontal="left", vertical="center")

    for row in sheet[cell_range]:
        for item in row:
            item.fill = PatternFill("solid", fgColor=NAVY)

    sheet.row_dimensions[cell.row].height = 24


def add_table_header(sheet, row_number, first_col, headers):
    """表の見出し行を作る。"""
    for col, title in enumerate(headers, start=first_col):
        cell = sheet.cell(row_number, col, title)
        set_cell_style(
            cell,
            bold=True,
            color=NAVY,
            fill=PALE_BLUE
        )
    sheet.row_dimensions[row_number].height = 23


# ==================================
# 5. ダッシュボードシート作成
# ==================================

def create_dashboard(workbook, ranking, daily_ranking,
                     total_sales, sales_count, product_count):
    """見出し帯とグラフの領域を分離したダッシュボードを作成する。"""
    sheet = workbook.active
    sheet.title = "ダッシュボード"
    sheet.sheet_view.showGridLines = False
    sheet.sheet_view.zoomScale = 85

    widths = {"A": 3, "B": 15, "C": 23, "D": 16, "E": 3,
              "F": 13, "G": 13, "H": 13, "I": 13, "J": 13,
              "K": 13, "L": 13, "M": 13}
    for column, width in widths.items():
        sheet.column_dimensions[column].width = width

    # タイトル帯
    sheet.merge_cells("B2:M2")
    sheet["B2"] = "売上分析レポート"
    sheet["B2"].font = Font(name=FONT_NAME, size=20, bold=True, color=WHITE)
    sheet["B2"].fill = PatternFill("solid", fgColor=NAVY)
    sheet["B2"].alignment = Alignment(horizontal="center", vertical="center")
    for row in sheet["B2:M2"]:
        for cell in row:
            cell.fill = PatternFill("solid", fgColor=NAVY)
    sheet.row_dimensions[2].height = 42
    sheet.merge_cells("B3:M3")
    sheet["B3"] = "SALES PERFORMANCE REPORT"
    sheet["B3"].font = Font(name=FONT_NAME, size=9, color="718096")
    sheet["B3"].alignment = Alignment(horizontal="right", vertical="center")
    sheet.row_dimensions[3].height = 22

    add_section_title(sheet, "B5:M5", "売上サマリー")
    summary_items = [
        ("B7:D7", "B8:D9", "売上合計", total_sales, '¥#,##0'),
        ("F7:H7", "F8:H9", "販売件数", sales_count, '#,##0"件"'),
        ("J7:M7", "J8:M9", "商品種類", product_count, '#,##0"種類"')
    ]
    for label_range, value_range, label, value, number_format in summary_items:
        sheet.merge_cells(label_range)
        sheet.merge_cells(value_range)
        label_cell = sheet[label_range.split(":")[0]]
        label_cell.value = label
        label_cell.font = Font(name=FONT_NAME, size=10, bold=True, color=NAVY)
        label_cell.alignment = Alignment(horizontal="center", vertical="center")
        value_cell = sheet[value_range.split(":")[0]]
        value_cell.value = value
        value_cell.number_format = number_format
        value_cell.font = Font(name=FONT_NAME, size=18, bold=True, color=BLUE)
        value_cell.alignment = Alignment(horizontal="center", vertical="center")
        for target_range, fill_color in ((label_range, PALE_BLUE), (value_range, "F7FAFD")):
            for row in sheet[target_range]:
                for cell in row:
                    cell.fill = PatternFill("solid", fgColor=fill_color)
                    cell.border = Border(left=Side(style="thin", color=BORDER_COLOR),
                                         right=Side(style="thin", color=BORDER_COLOR),
                                         top=Side(style="thin", color=BORDER_COLOR),
                                         bottom=Side(style="thin", color=BORDER_COLOR))
    sheet.row_dimensions[7].height = 24
    sheet.row_dimensions[8].height = 27
    sheet.row_dimensions[9].height = 22

    # ランキング見出し・表。グラフは右側に独立配置。
    add_section_title(sheet, "B11:M11", "商品別売上ランキング（上位10件）")
    add_table_header(sheet, 12, 2, ["順位", "商品名", "売上金額"])
    top_ranking = ranking[:10]
    for rank_number, (product, amount) in enumerate(top_ranking, start=1):
        row_number = 12 + rank_number
        for col, value in enumerate([rank_number, product, amount], start=2):
            cell = sheet.cell(row_number, col, value)
            set_cell_style(cell, color=TEXT,
                           fill=PALE_GRAY if rank_number % 2 == 0 else WHITE,
                           horizontal="left" if col == 3 else ("right" if col == 4 else "center"),
                           number_format='¥#,##0' if col == 4 else None)
        sheet.row_dimensions[row_number].height = 22

    if ranking:
        sheet.add_image(create_ranking_image(ranking), "F13")

    # 構成比グラフ専用セクション。直前のグラフの下に余白を確保。
    add_section_title(sheet, "B31:M31", "商品別売上構成比")
    if ranking:
        sheet.add_image(create_doughnut_image(ranking), "F33")

    # 日別推移専用セクション。表は左、グラフは右。
    add_section_title(sheet, "B51:M51", "日別売上推移")
    add_table_header(sheet, 52, 2, ["日付", "売上金額"])
    for index, (sale_date, amount) in enumerate(daily_ranking, start=1):
        row_number = 52 + index
        date_cell = sheet.cell(row_number, 2, sale_date)
        date_cell.number_format = "yyyy-mm-dd"
        amount_cell = sheet.cell(row_number, 3, amount)
        amount_cell.number_format = '¥#,##0'
        for col in (2, 3):
            set_cell_style(sheet.cell(row_number, col),
                           fill=PALE_GRAY if index % 2 == 0 else WHITE,
                           horizontal="center")
        sheet.row_dimensions[row_number].height = 21
    if daily_ranking:
        sheet.add_image(create_daily_image(daily_ranking), "F53")

    sheet.freeze_panes = None
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.paperSize = sheet.PAPERSIZE_A3
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    sheet.page_margins = PageMargins(left=0.25, right=0.25, top=0.45, bottom=0.45)
    last_dashboard_row = max(66, 53 + len(daily_ranking) + 2)
    sheet.print_area = f"B2:M{last_dashboard_row}"
    sheet.oddFooter.center.text = "売上分析レポート  |  &P / &N"
    return sheet




# ==================================
# 6. 売上明細シート作成
# ==================================

def create_detail_sheet(workbook, sales_data):
    """売上明細を別シートに分け、検索・並べ替えしやすくする。"""
    sheet = workbook.create_sheet("売上明細")
    sheet.sheet_view.showGridLines = False
    sheet.sheet_view.zoomScale = 90

    headers = ["日付", "商品名", "数量", "単価", "売上金額"]
    for col, title in enumerate(headers, start=1):
        cell = sheet.cell(1, col, title)
        set_cell_style(cell, bold=True, color=NAVY, fill=PALE_BLUE)

    for index, data in enumerate(sales_data, start=2):
        values = [
            data["date"],
            data["product"],
            data["quantity"],
            data["price"],
            data["amount"]
        ]

        for col, value in enumerate(values, start=1):
            cell = sheet.cell(index, col, value)
            fill = PALE_GRAY if index % 2 == 0 else WHITE
            set_cell_style(
                cell,
                fill=fill,
                horizontal="left" if col == 2 else "center",
                number_format=(
                    "yyyy-mm-dd" if col == 1 else
                    ('¥#,##0' if col in (4, 5) else None)
                )
            )

    # 明細シートは1行目だけ固定する
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:E{max(1, len(sales_data) + 1)}"

    widths = {"A": 16, "B": 25, "C": 12, "D": 16, "E": 18}
    for column, width in widths.items():
        sheet.column_dimensions[column].width = width

    sheet.row_dimensions[1].height = 25
    return sheet


# ==================================
# 7. メイン処理
# ==================================

def main():
    try:
        sales_data, error_count = load_sales_data()

        if not sales_data:
            print("売上データがありません。先にapp.pyでデータを登録してください。")
            return

        ranking, daily_ranking, total_sales, product_count = calculate_sales(
            sales_data
        )

        workbook = Workbook()
        create_dashboard(
            workbook,
            ranking,
            daily_ranking,
            total_sales,
            len(sales_data),
            product_count
        )
        create_detail_sheet(workbook, sales_data)

        workbook.save(EXCEL_FILE)

        print("\n===== Excelレポート作成結果 =====")
        print(f"正常データ：{len(sales_data)}件")
        print(f"除外データ：{error_count}件")
        print(f"売上合計：{total_sales:,}円")
        print("作成したグラフ：商品別ランキング／商品別構成比／日別売上推移")
        print(f"保存先：{EXCEL_FILE.name}")

    except (FileNotFoundError, ValueError) as error:
        print(f"エラー：{error}")


if __name__ == "__main__":
    main()
