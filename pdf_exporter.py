import os
import xlwings as xw
import re

target_keywords = ["支払通知書"]

base_dir = os.path.dirname(os.path.abspath(__file__))
input_folder = os.path.join(base_dir, "checked")   # Excelが入ってるフォルダ
pdf_folder = os.path.join(base_dir, "pdf")  #pdf出力先のフォルダ

os.makedirs(pdf_folder, exist_ok=True)

app = xw.App(visible=False)
app.display_alerts = False

def is_black_tab(ws):
    try:
        return ws.api.Tab.Color == 0 and ws.api.Tab.ColorIndex != -4142
    except:
        return False
    
def is_payment_notification(ws):
    return any(k in ws.name for k in target_keywords)

def clean_company(name):
    name = name.replace("\n", "")
    name = re.sub(r"様$", "", name)
    name = re.sub(r"\s+", "", name)
    name = re.sub("株式会社", "", name)
    name = re.sub("合同会社", "", name)
    name = re.sub("有限会社", "", name)
    return name


try:
    for file_name in os.listdir(input_folder):
        if not file_name.endswith((".xlsx", ".xlsm")) or file_name.startswith("~$"):
            continue

        file_path = os.path.join(input_folder, file_name)
        pdf_name = file_name.replace(".xlsx", ".pdf").replace(".xlsm", ".pdf")
        pdf_path = os.path.join(pdf_folder, pdf_name)

        wb = None

        try:
            print(f"PDF変換開始: {file_name}")

            wb = app.books.open(file_path)
            target_sheets = []

            for ws in wb.sheets:

                # 黒タブはスキップ
                if is_black_tab(ws):
                    print(f"スキップ（黒タブ）: {ws.name}")
                    continue

                if "DMM" in wb.name:
                    if "明細" in ws.name:
                        continue

                # 支払通知書は個別PDF
                if is_payment_notification(ws):
                    company_match = re.search(r'【(.+?)様】', file_name)
                    company = str(company_match.group(1))
                    company_text = f"（{company}様分）"

                    month_match = re.search(r"(\d+)月", file_name)
                    month = str(month_match.group(1))

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
                    pn_pdf_path = os.path.join(pdf_folder, f"{pay_notice_name}.pdf")

                    try:
                        ws.api.ExportAsFixedFormat(0, pn_pdf_path)
                        print(f"支払通知書: {pay_notice_name}")
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

            wb.sheets[target_sheets].select()

            if target_sheets:

                if "DMM" in wb.name:
                    company_match = re.search(r'【(.+?)様】', file_name)
                    company = str(company_match.group(1))

                    month_match = re.search(r"(\d+)月", file_name)
                    month = str(month_match.group(1))
                    dmm_filename = f"【{company}様】支払通知書_{month}月"

                    dmm_path = os.path.join(pdf_folder, f"{dmm_filename}.pdf")
                    wb.app.api.ActiveSheet.ExportAsFixedFormat(0, dmm_path)
                    print("PDF出力 : ", dmm_filename, ".pdf")

                elif "キュービクル" in wb.name:
                    company_match = re.search(r'[（(](.+?)様', file_name)
                    company = str(company_match.group(1))
                
                    month_match = re.search(r"(\d+)月", file_name)
                    month = str(month_match.group(1))
                    cm_filename = f"【{company}様】請求内訳_{month}月"

                    cm_path = os.path.join(pdf_folder, f"{cm_filename}.pdf")
                    wb.app.api.ActiveSheet.ExportAsFixedFormat(0, cm_path)
                    print("PDF出力 : ", cm_filename, ".pdf")
                
                else:
                    wb.app.api.ActiveSheet.ExportAsFixedFormat(0, pdf_path)
                    print("PDF出力 : ", pdf_name,)

                print("----作業完了、次のファイルへ----")

        except Exception as e:
            print(f"エラー: {file_name} / {e}")

        finally:
            if wb:
                wb.close()

finally:
    app.quit()

print("全PDF変換完了！")