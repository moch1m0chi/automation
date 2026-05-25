import os
import xlwings as xw
import re
from pathlib import Path

#-----------------------------------------------
# 分岐ルール
#-----------------------------------------------

RULES = [
    {
        # DMM
        # ファイル名指定 : 【○○様】支払通知書_n月.pdf
        "keyword": "DMM",
        "match": lambda wb: "DMM" in wb.name,
        "filter_sheets": lambda sheets: [s for s in sheets if "明細" not in s],
        "build_filename": lambda c, m: f"【{c}】支払通知書_{m}月",
        "executor": "special"
    },
    {
        # キュービクル  
        # ファイル名指定 : 【○○様】請求内訳_ n月.pdf
        "keyword": "キュービクル",
        "match": lambda wb: "キュービクル" in wb.name,
        "filter_sheets": lambda sheets: sheets,
        "build_filename": lambda c, m: f"【{c}】請求内訳_{m}月",
        "executor": "special"
    }
]

DEFAULT_RULE = {
    "name": "default",
    "filter_sheets": lambda sheets: sheets,
    "build_filename": None,  # 使わないのでnone
    "executor": "default"
}

#-----------------------------------------------
# rules、インフラ初期化
#-----------------------------------------------

target_keywords = ["支払通知書"]

def build_paths(base_dir):
    input_folder = base_dir / "checked"   # Excelが入ってるフォルダ
    pdf_folder = base_dir / "pdf"         # PDF出力先のフォルダ

    pdf_folder.mkdir(parents=True, exist_ok=True)   # フォルダがなかった場合は作成

    return {
        "input" : input_folder,
        "output" : pdf_folder 
    }

def init_excel_app():
    app = xw.App(visible=False)
    app.display_alerts = False
    return app

#-----------------------------------------------
# 関数
#-----------------------------------------------

# 制御系

# メインロジック
def export_payment_notification(ws, company, month, save_folder):
    payee = get_payee_name(ws)

    pay_notice_name = f"【{payee}様】支払通知書（{company}分）_{month}月"
    pn_pdf_path = create_pdf_file_path(save_folder, pay_notice_name)

    try:
        output_payment_notification(ws, str(pn_pdf_path))
        log(f"支払通知書出力 : {pay_notice_name}.pdf")
    except Exception as e:
        log(f"エラー: {ws.name} / {e}")

    return

def export_renamed_pdf(wb, rule, target_sheets, company, month, save_folder, pdf_name):
    target_sheets = rule["filter_sheets"](target_sheets)
    new_wb = create_temp_workbook(wb, target_sheets)

    new_pdf_name = rule["build_filename"](company, month)
    renamed_path = create_pdf_file_path(save_folder, new_pdf_name)

    output_pdf(new_wb, str(renamed_path))
    log(f"PDF出力 : {new_pdf_name}.pdf")

def export_pdf(wb, rule, target_sheets, copany_name, month, save_folder, pdf_name):
    pdf_path = save_folder / pdf_name
    new_wb = create_temp_workbook(wb, target_sheets)

    output_pdf(new_wb, str(pdf_path))
    log(f"PDF出力 : {pdf_name}")

# 判定系
def is_black_tab(ws):
    try:
        return ws.api.Tab.Color == 0 and ws.api.Tab.ColorIndex != -4142
    except Exception as e:
        log(f"エラー内容: {e}")
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
    return DEFAULT_RULE

# 変換系
def clean_company(name): # 様は抜けて出力されるので注意！
    name = name.replace("\n", "")
    name = re.sub(r"様$", "", name)
    name = re.sub(r"\s+", "", name)
    name = re.sub("株式会社", "", name)
    name = re.sub("合同会社", "", name)
    name = re.sub("有限会社", "", name)
    return name

def replace_file_name_specified(rule, company, month):
    return rule["filename"](company, month)

def exclude_invisible_sheets(wb, target_sheets):
    valid_sheets = []

    for name in target_sheets:
        try:
            sheet = wb.sheets[name]
            if sheet.api.Visible == -1:
                valid_sheets.append(name)
        except:
            log(f"{name} は存在しない")

    return valid_sheets

def build_file_context(file_path, pdf_root):
    file_name = file_path.name
    base_name = Path(file_name).stem

    company = get_company_name(base_name)
    month = get_month(base_name)

    if not company:
        raise ValueError(f"会社名取得失敗: {base_name}")
    
    save_folder = pdf_root / company
    save_folder.mkdir(parents=True, exist_ok=True)
    
    pdf_name = base_name + ".pdf"
    pdf_path = save_folder / pdf_name

    return {
        "file_name" : file_name,
        "base_name" : base_name,
        "company" : company,
        "month" : month,
        "save_folder" : save_folder,
        "pdf_name" : pdf_name,
        "pdf_path" : pdf_path
    }

# ユーティリティ

def paths(base_dir):

    input_folder = base_dir / "checked"   # Excelが入ってるフォルダ
    pdf_folder = base_dir / "pdf"         # PDF出力先のフォルダ

    pdf_folder.mkdir(parents=True, exist_ok=True)   # フォルダがなかった場合は作成

    return {
        "input" : input_folder,
        "output" : pdf_folder
    }

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

def create_pn_file_name(payee, company, month):
    return f"【{payee}様】支払通知書（{company}分）_{month}月"

def create_pdf_file_path(save_folder, pdf_name):
    return save_folder / f"{pdf_name}.pdf"

def log(msg):
    print(msg)

# pdf出力
def output_payment_notification(ws, pdf_filepath):  # シート出力
    ws.api.ExportAsFixedFormat(0, pdf_filepath)

def output_pdf(new_wb, pdf_filepath):   # ブックをまとめて出力
    new_wb.api.ExportAsFixedFormat(0, pdf_filepath)

#-----------------------------------------------
# executors
#-----------------------------------------------

EXECUTORS = {
    "special": export_renamed_pdf,
    "default": export_pdf,
}

#-------------------------------------------------------------------------------------------------------------------------------

#-----------------------------------------------
# 実行部
#-----------------------------------------------

try:
    base_dir = Path(__file__).resolve().parent

    paths = build_paths(base_dir)

    app = init_excel_app()

    for file_path in paths["input"].iterdir():

        if is_out_of_scope(file_path.name):
            continue

        ctx = build_file_context(file_path, paths["output"])

        wb = None

        try:
            log(f"\nPDF変換開始: {ctx["file_name"]}")

            # xlsxファイルを開く
            wb = app.books.open(str(file_path))

            # 支払通知書以外のシートをまとめるリストを初期化
            target_sheets = []

            for ws in wb.sheets:

                if is_black_tab(ws):    # 黒タブ=既に終わっている案件はスルー
                    log(f" →スキップ（黒タブ）: {ws.name}")
                    continue

                # 支払通知書は個別PDFで出力
                if is_payment_notification(ws):

                    export_payment_notification(ws, ctx["company"], ctx["month"], ctx["save_folder"])

                # それ以外はまとめ用に追加
                target_sheets.append(ws.name)

            # 非表示シートを除外
            target_sheets = exclude_invisible_sheets(wb, target_sheets)

            # まとめたシートの一括出力
            if target_sheets:

                rule = find_rule(ctx["base_name"], RULES)
                
                executor = EXECUTORS[rule["executor"]]
                executor(wb, rule, target_sheets, ctx["company"], ctx["month"], ctx["save_folder"], ctx["pdf_name"])

                log("----pdf出力完了----")

        except Exception as e:
            log(f"エラー: {ctx["file_name"]} / {e}")

        finally:
            if wb:
                wb.close()

finally:
    app.quit()

log("\n全PDF変換完了！")