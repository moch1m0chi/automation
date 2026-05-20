import os
import xlwings as xw
import re
from pathlib import Path


#-----------------------------------------------
# 条件定義、今回はシートが支払通知書か否かの判定のみ
#-----------------------------------------------

target_keywords = ["支払通知書"]

#-----------------------------------------------
# 分岐ルール
#-----------------------------------------------

# DMM
# ファイル名指定 : 【○○様】支払通知書_n月.pdf
# 3枚の支払通知書をひとつのpdfにまとめる

# キュービクル  
# ファイル名指定 : 【○○様】請求内訳_ 2月.pdf

RULES = [
    {
        "keyword": "DMM",
        "filename": lambda company_name, month: f"【{company_name}】支払通知書_{month}月",
        "filter_sheets": lambda sheets: [s for s in sheets if "明細" not in s],
    },
    {
        "keyword": "キュービクル",
        "filename": lambda company_name, month: f"【{company_name}】請求内訳_{month}月",
        "filter_sheets": lambda sheets: sheets,
    }
]

#-----------------------------------------------
# 共通ディレクトリ指定
#-----------------------------------------------

base_dir = Path(__file__).resolve().parent
input_folder = base_dir / "checked"   # Excelが入ってるフォルダ
pdf_folder = base_dir / "pdf"         # PDF出力先のフォルダ

pdf_folder.mkdir(parents=True, exist_ok=True)   # フォルダがなかった場合は作成

app = xw.App(visible=False)
app.display_alerts = False

#-----------------------------------------------
# 関数
#-----------------------------------------------

# 制御系

# メインロジック
def export_payment_notification(ws, company_name, month, save_folder):
    payee = get_payee_name(ws)

    pay_notice_name = f"【{payee}様】支払通知書（{company_name}分）_{month}月"
    pn_pdf_path = create_pdf_file_path(save_folder, pay_notice_name)

    try:
        output_payment_notification(ws, str(pn_pdf_path))
        print(f"支払通知書出力 : {pay_notice_name}.pdf")
    except Exception as e:
        print(f"エラー: {ws.name} / {e}")

    return

def export_renamed_pdf(wb, rule, target_sheets, company_name, month, save_folder):
    target_sheets = rule["filter_sheets"](target_sheets)

    new_wb = create_temp_workbook(wb, target_sheets)

    new_pdf_name = replace_file_name_specified(rule, company_name, month)
    renamed_path = create_pdf_file_path(save_folder, new_pdf_name)

    output_pdf(new_wb, str(renamed_path))
    print(f"PDF出力 : {new_pdf_name}.pdf")

def export_pdf(wb, target_sheets, pdf_path, pdf_name):

    new_wb = create_temp_workbook(wb, target_sheets)

    output_pdf(new_wb, str(pdf_path))
    print(f"PDF出力 : {pdf_name}")

# 判定系
def is_black_tab(ws):
    try:
        return ws.api.Tab.Color == 0 and ws.api.Tab.ColorIndex != -4142
    except Exception as e:
        print(f"エラー内容: {e}")
        return False
    
def is_payment_notification(ws):
    return any(k in ws.name for k in target_keywords)

def is_out_of_scope(file_name):
    return (
        not file_name.endswith((".xlsx", ".xlsm")) 
        or file_name.startswith("~$")
    )

def find_rule(wb_name, RULES):
    for rule in RULES:
        if rule["keyword"] in wb_name:
            return rule
    return None

# 変換系
def clean_company(name): # 様は抜けて出力されるので注意！
    name = name.replace("\n", "")
    name = re.sub(r"様$", "", name)
    name = re.sub(r"\s+", "", name)
    name = re.sub("株式会社", "", name)
    name = re.sub("合同会社", "", name)
    name = re.sub("有限会社", "", name)
    return name

def replace_file_name_specified(rule, company_name, month):
    return rule["filename"](company_name, month)

def exclude_invisible_sheets(wb, target_sheets):
    valid_sheets = []

    for name in target_sheets:
        try:
            sheet = wb.sheets[name]
            if sheet.api.Visible == -1:
                valid_sheets.append(name)
        except:
            print(f"{name} は存在しない")

    return valid_sheets

# ユーティリティ

def get_month(file_name):
    month_match = re.search(r"(\d+)月", file_name)
    return str(month_match.group(1)) if month_match else None

def get_company_name(base_name): # 「○○会社様」の形で出力される
    pattern = re.compile(r"[【（]([^】(]+?様)[】_）]")
    match = pattern.search(base_name)
    return str(match.group(1))  if match else None

def get_payee_name(ws):
    payee = None
    values = ws.range("A1:B3").value

    for row in values:
        for cell in row:
            if isinstance(cell, str):
                if re.search(r'(株式会社|有限会社|合同会社|福祉会)', cell):
                    payee = cell
                    break
        if payee:
            break
    
    return clean_company(payee) if payee else None

def filter_sheets(rule, target_sheets):
    return rule["filter_sheets"](target_sheets)

def create_temp_workbook(wb, target_sheets):
    wb.sheets[target_sheets].api.Copy()
    new_wb = xw.books.active
    return new_wb

def create_pdf_file_path(save_folder, pdf_name):
    return save_folder / f"{pdf_name}.pdf"


# pdf出力
def output_payment_notification(ws, pdf_filepath):  # シート出力
    ws.api.ExportAsFixedFormat(0, pdf_filepath)

def output_pdf(new_wb, pdf_filepath):   # ブックをまとめて出力
    new_wb.api.ExportAsFixedFormat(0, pdf_filepath)

#-------------------------------------------------------------------------------------------------------------------------------

#-----------------------------------------------
# 実行部
#-----------------------------------------------

try:
    for file_path in input_folder.iterdir():
        file_name = file_path.name

        if is_out_of_scope(file_name):
            continue
        
        # pdf出力用の情報を抽出
        base_name = Path(file_name).stem  # 拡張子除去
        company_name = get_company_name(base_name)
        month = get_month(base_name)
        pdf_name = base_name + ".pdf"

        # 保存先の各社フォルダを作成
        save_folder = pdf_folder / company_name
        save_folder.mkdir(parents=True, exist_ok=True)
        
        pdf_path = save_folder / pdf_name

        wb = None

        try:
            print(f"\nPDF変換開始: {file_name}")

            # xlsxファイルを開く
            wb = app.books.open(str(file_path))

            # 支払通知書以外のシートをまとめるリストを初期化
            target_sheets = []

            for ws in wb.sheets:

                if is_black_tab(ws):    # 黒タブ=既に終わっている案件はスルー
                    print(f" →スキップ（黒タブ）: {ws.name}")
                    continue

                # 支払通知書は個別PDFで出力
                if is_payment_notification(ws):

                    export_payment_notification(ws, company_name, month, save_folder)

                # それ以外はまとめ用に追加
                target_sheets.append(ws.name)

            # 非表示シートを除外
            target_sheets = exclude_invisible_sheets(wb, target_sheets)

            # まとめたシートの一括出力
            if target_sheets:
                
                for rule in RULES:
                    if rule["keyword"] in wb.name:
                        # ファイル名の変更が必要な場合の処理
                        export_renamed_pdf(wb, rule, target_sheets, company_name, month, save_folder)
                        break
                else:
                    # デフォルト処理
                    export_pdf(wb, target_sheets, pdf_path, pdf_name)

                print("----pdf出力完了----")

        except Exception as e:
            print(f"エラー: {file_name} / {e}")

        finally:
            if wb:
                wb.close()

finally:
    app.quit()

print("全PDF変換完了！")