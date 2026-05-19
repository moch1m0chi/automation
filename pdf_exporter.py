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
RULES = [
    {
        "keyword": "DMM",
        "filename": lambda company_name, month: f"【{company_name}】支払通知書_{month}月",
        "exclude": lambda sheets: [s for s in sheets if "明細" not in s],
    },
    {
        "keyword": "キュービクル",
        "filename": lambda company_name, month: f"【{company_name}】請求内訳_{month}月",
        "exclude": lambda sheets: sheets,
    }
]

#-----------------------------------------------
# ディレクトリ指定
#-----------------------------------------------

base_dir = Path(__file__).resolve().parent
input_folder = base_dir / "checked"   # Excelが入ってるフォルダ
pdf_folder = base_dir / "pdf"         # PDF出力先のフォルダ

pdf_folder.mkdir(parents=True, exist_ok=True)

app = xw.App(visible=False)
app.display_alerts = False

#-----------------------------------------------
# 関数
#-----------------------------------------------

# プロセス系



# 判定系
def is_black_tab(ws):
    try:
        return ws.api.Tab.Color == 0 and ws.api.Tab.ColorIndex != -4142
    except:
        return False
    
def is_payment_notification(ws):
    return any(k in ws.name for k in target_keywords)

def is_out_of_scope(file_name):
    return (
        not file_name.endswith((".xlsx", ".xlsm")) 
        or file_name.startswith("~$")
    )
        

# 変換系

def get_month(file_name):
    month_match = re.search(r"(\d+)月", file_name)
    return str(month_match.group(1))

# ユーティリティ

def get_company_name(base_name):
    pattern = re.compile(r"[【（]([^】(]+?様)[】_）]")
    match = pattern.search(base_name)
    return str(match.group(1))

def clean_company(name):
    name = name.replace("\n", "")
    name = re.sub(r"様$", "", name)
    name = re.sub(r"\s+", "", name)
    name = re.sub("株式会社", "", name)
    name = re.sub("合同会社", "", name)
    name = re.sub("有限会社", "", name)
    return name

# pdf保存
def output_payment_notification(ws, pdf_filepath):
    ws.api.ExportAsFixedFormat(0, pdf_filepath)

def output_pdf(new_wb, pdf_filepath):
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

        file_path = input_folder / file_name

        base_name = Path(file_name).stem  # 拡張子除去
        pdf_name = base_name + ".pdf"
        
        company_name = get_company_name(base_name)

        save_folder = pdf_folder / company_name
        save_folder.mkdir(parents=True, exist_ok=True)
        
        pdf_path = save_folder / pdf_name

        wb = None

        try:
            print(f"\nPDF変換開始: {file_name}")

            wb = app.books.open(str(file_path))
            target_sheets = []

            for ws in wb.sheets:

                # 黒タブはスキップ
                if is_black_tab(ws):
                    print(f" →スキップ（黒タブ）: {ws.name}")
                    continue

                # 支払通知書は個別PDF
                if is_payment_notification(ws):
                    company = get_company_name(file_name)

                    month = get_month(file_name)

                    # 支払先情報をシートから取得
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
                    
                    payee = clean_company(payee)

                    pay_notice_name = f"【{payee}様】支払通知書（{company}様分）_{month}月"
                    pn_pdf_path = save_folder / f"{pay_notice_name}.pdf"


                    try:
                        output_payment_notification(ws, str(pn_pdf_path))
                        print(f"PDF出力 : {pay_notice_name}.pdf")
                    except Exception as e:
                        print(f"エラー: {ws.name} / {e}")

                    continue

                # それ以外はまとめ用に追加
                target_sheets.append(ws.name)

            valid_sheets = []

            for name in target_sheets:
                try:
                    sheet = wb.sheets[name]
                    if sheet.api.Visible == -1:
                        valid_sheets.append(name)
                except:
                    print(f"{name} は存在しない")

            target_sheets = valid_sheets

            if target_sheets:
                for rule in RULES:
                    if rule["keyword"] in wb.name:
                        target_sheets = rule["exclude"](target_sheets)
                        wb.sheets[target_sheets].api.Copy()
                        new_wb = xw.books.active

                        filename = rule["filename"](company_name, month)
                        path = save_folder / f"{filename}.pdf"

                        output_pdf(new_wb, str(path))
                        break
                else:
                    # デフォルト処理
                    wb.sheets[target_sheets].api.Copy()
                    new_wb = xw.books.active
                    output_pdf(new_wb, str(pdf_path))

                # if "DMM" in wb.name:
                #     target_sheets = [s for s in target_sheets if "明細" not in s]
                #     wb.sheets[target_sheets].api.Copy()
                #     new_wb = xw.books.active

                #     month = get_month(file_name)
                #     dmm_filename = f"【{company_name}】支払通知書_{month}月"
                #     dmm_path = save_folder / f"{dmm_filename}.pdf"

                #     output_pdf(new_wb, str(dmm_path))
                #     print(f"PDF出力 : {dmm_filename}.pdf")
                #     new_wb.close()

                # elif "キュービクル" in wb.name:
                    
                #     wb.sheets[target_sheets].api.Copy()
                #     new_wb = xw.books.active
                
                #     month = get_month(file_name)
                #     cm_filename = f"【{company_name}】請求内訳_{month}月"
                #     cm_path = save_folder / f"{cm_filename}.pdf"

                #     output_pdf(new_wb, str(cm_path))
                #     print(f"PDF出力 : {cm_filename}.pdf")
                #     new_wb.close()
                
                # else:
                #     wb.sheets[target_sheets].api.Copy()
                #     new_wb = xw.books.active
                #     output_pdf(new_wb, str(pdf_path))
                #     print(f"PDF出力 : {pdf_name}")
                #     new_wb.close()

                print("----pdf出力完了----")

        except Exception as e:
            print(f"エラー: {file_name} / {e}")

        finally:
            if wb:
                wb.close()

finally:
    app.quit()

print("全PDF変換完了！")