import os
from datetime import datetime, timedelta
import xlwings as xw
import re
import calendar
    
base_dir = os.path.dirname(os.path.abspath(__file__))
input_folder = os.path.join(base_dir, "data")
output_folder = os.path.join(base_dir, "output")

os.makedirs(output_folder, exist_ok=True)

#正規表現(数値)/(数値)回目をコンパイル
pattern_a = re.compile(r"(\d{4})年(\d{1,2})月利用分")
pattern_b = re.compile(r"(\d+)/(\d+回目)")

#================================
#関数
#================================

#ファイル名の月繰り上げ
def increment_month_in_filename(file_name):
    match = re.search(r"(\d+)月", file_name)
    
    if match:
        month = int(match.group(1))
        new_month = month + 1

        if new_month > 12:
            new_month = 1
        
        # 置き換え
        new_name = re.sub(r"\d+月", f"{new_month}月", file_name)
        return new_name
    else:
        return file_name  # 月が見つからなければそのまま
    
def increment_year_month_text(text):
    def repl(match):
        year = int(match.group(1))
        month = int(match.group(2))

        month += 1
        if month > 12:
            month = 1
            year += 1

        return f"{year}年{month}月利用分"

    return pattern_a.sub(repl, text)

def add_one_month(dt):
    year = dt.year
    month = dt.month + 1
    if month > 12:
        month = 1
        year += 1

    last_day = calendar.monthrange(year, month)[1]
    day = min(dt.day, last_day)
    return dt.replace(year = year, month= month, day = day)

#================================
#メイン処理
#================================
app = xw.App(visible=False)
app.screen_updating = False
app.display_alerts = False

target_keywords = ["確定合意書", "DMM（秀", "御請求書"]
#SOURCE_SHEET = ["DMM（秀商）", "報酬額確定合意書"]
DATE_COLUMNS = [1,7]

try:
    for file_name in os.listdir(input_folder):
        if not file_name.endswith((".xlsx", ".xlsm")) or file_name.startswith("~$"):
            continue
            
        file_path = os.path.join(input_folder, file_name)
        print(f"\n処理開始: {file_name}")

        wb = None

        try:
            wb = app.books.open(file_path)

            #================================
            #データの取得
            #================================

            all_data = {}
            
            for ws in wb.sheets:
                #if ws.name not in SOURCE_SHEET:
                    #continue

                if any(k in ws.name for k in target_keywords):
                    ur = ws.used_range
                    all_data[ws.name] = {
                        "values": ur.value,
                        "formats": ur.number_format,
                        "formulas": ur.formula,
                        "base_row": ur.row,
                        "base_col": ur.column
                        }

            #================================
            #月の繰り上げ
            #================================

            for ws in wb.sheets:
                if ws.name not in all_data:
                    continue

                print("対象シート:", ws.name)

                data = all_data[ws.name]
                values = data["values"]
                formats = data["formats"]
                formulas = data["formulas"]
                base_row = data["base_row"]
                base_col = data["base_col"]

                if values is None:
                    continue
                if formats is None:
                    formats = [[""] * len(values[0]) for _ in range(len(values))]
                if formulas is None:
                    formulas = [[None]*len(values[0]) for _ in range(len(values))]

                if not isinstance(values, list):
                    continue
                if not isinstance(values[0], list):
                    values = [values]

                if not isinstance(formats, list):
                    formats = [formats]
                if not isinstance(formats[0], list):
                    formats = [[f] for f in formats]

                if not isinstance(formulas, list):
                    formulas = [formulas]
                if not isinstance(formulas[0], list):
                    formulas = [[f] for f in formats]

                max_col = max(len(row) for row in values)

                for r, row in enumerate(values):
                    for c in range(max_col):
                        val = row[c] if c < len(row) else None

                        if val is None:
                            continue

                        if c not in DATE_COLUMNS:
                            continue

                        # 安全取得
                        fmt = ""
                        formula = None

                        if r < len(formats) and c < len(formats[r]):
                            fmt = str(formats[r][c]).lower()

                        if r < len(formulas) and c < len(formulas[r]):
                            formula = formulas[r][c]

                        if formula:
                            continue

                        if not isinstance(val, (datetime, int, float, str)):
                            continue

                        if isinstance(val, datetime):
                            new_val = add_one_month(val)

                        elif isinstance(val, (int, float)):
                            if not (40000 < val <50000):
                                continue
                            dt = datetime(1899, 12, 30) + timedelta(days=val)
                            new_val = add_one_month(dt)
                    
                        elif isinstance(val, str):
                            try:
                                dt = datetime.strptime(val, "%Y/%m/%d")
                                new_val = add_one_month(dt)
                            except:
                                continue

                        else:
                            continue

                        ws.cells(base_row+r, base_col+c).value = new_val
                        print(new_val, "を入力")

            #================================
            #2026年〇月利用分を更新
            #================================
            for ws in wb.sheets:
                values = ws.used_range.value

                if not values:
                    continue

                for r, row in enumerate(values):
                    for c, val in enumerate(row):
                        if isinstance(val, str):
                            new_val = increment_year_month_text(val)

                            if new_val != val:
                                ws.cells(r+1, c+1).value = new_val
                                print(val, "→", new_val)

            #================================
            #n/24回目の更新
            #================================
            for ws in wb.sheets:
                values = ws.range("A1:N400").value

                if not values:
                    continue

                for r, row in enumerate(values):  #行を抜き出し
                    for c, val in enumerate(row):    #抜き出した行のセルを走査
                        if isinstance(val, str):  #セルの値value)はstr型(文字列)？
                            match = pattern_b.search(val)   #Matchオブジェクト
                            if match:   #matchの中身があるとTrueとして判定される。NoneだとFalse扱い
                                print("書換え対象:", match.group(0))
                                left = int(match.group(1)) + 1  #matchオブジェクトのmatch(1)、ここでは(/d+)に相当する部分
                                right = match.group(2)

                                ws.cells(r+1, c+1).value = f"{left}/{right}"   #f文字列
                                if r<3:
                                    print(f"{left}/{right}に更新")

            new_file_name = increment_month_in_filename(file_name)  #ファイル名の月を繰り上げ
            output_path = os.path.join(output_folder, new_file_name)

            wb.save(output_path)
            print(f"保存完了:{new_file_name}")

        except Exception as e: 
            print(f"エラー発生:{file_name}")
            print(f"内容:{e}") 

        finally: 
            try:
                if wb: 
                    wb.close()
            except:
                print("えらー")
                pass
finally:
    app.quit()

print("\n全処理完了！")