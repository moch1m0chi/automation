import os
import xlwings as xw

input_folder = "output"   # Excelが入ってるフォルダ
pdf_folder = "pdf"

os.makedirs(pdf_folder, exist_ok=True)

app = xw.App(visible=False)
app.display_alerts = False

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
            wb.api.ExportAsFixedFormat(0, pdf_path)

            print(f"PDF出力完了: {pdf_name}")

        except Exception as e:
            print(f"エラー: {file_name} / {e}")

        finally:
            if wb:
                wb.close()

finally:
    app.quit()

print("全PDF変換完了！")