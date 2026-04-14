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

def add_one_month(dt):
    year = dt.year
    month = dt.month + 1
    if month > 12:
        month = 1
        year += 1

    last_day = calendar.monthrange(year, month)[1]
    day = min(dt.day, last_day)
    return dt.replace(year = year, month= month, day = day)

def update_month(val):
    new_val = None

    if isinstance(val, datetime):
        new_val = add_one_month(val)
        return new_val

    elif isinstance(val, str):
        try:
            dt = datetime.strptime(val, "%Y/%m/%d")
            new_val = add_one_month(dt)
            return new_val
        except:
            return

    else:
        return

def is_target_sheet(ws):
    return any(k in ws.name for k in target_keywords)

def get_allcells_in_target_sheet(ws):
    if not is_target_sheet(ws):
        return None
    
    ur = ws.used_range

    values = ur.value
    formats = ur.number_format
    formulas = ur.formula

    if not values:
        return None
    
    if not isinstance(values, list):
        values = [[values]]
    elif not isinstance(values[0], list):
        values = [values]

    rows = len(values)
    cols = len(values[0])

    if formats is None:
        formats = [[""] * cols for _ in range(rows)]
    else:
        if not isinstance(formats, list):
            formats = [[formats]]
        else:
            formats = [formats]
    
    if formulas is None:
        formulas = [[None] * cols for _ in range(rows)]
    else:
        if not isinstance(formulas, list):
            formulas = [[formulas]]
        elif not isinstance(formulas[0], list):
            formulas = [formulas]

    return {
        "values": values,
        "formats": formats,
        "formulas": formulas,
        "base_row": ur.row,
        "base_col": ur.column
        }

def read_each_data(data):
    if not data:
        return None, None, None, None, None
    
    values = data.get("values")
    formats = data.get("formats")
    formulas = data.get("formulas")
    base_row = data.get("base_row")
    base_col = data.get("base_col")

    if not values:
        return None, None, None, None, None
    
    if not isinstance(values, list):
        values = [[values]]
    if not isinstance(values[0], list):
        values = [values]
    
    if formats is None:
        formats = [[""] * len(values[0]) for _ in range(len(values))]
    else:
        if not isinstance(formats, list):
            formats = [[formats]]
        elif not isinstance(formats[0], list):
            formats = [formats]

    if formulas is None:
        formulas = [[None]*len(values[0]) for _ in range(len(values))]
    else:
        if not isinstance(formulas, list):
            formulas = [[formulas]]
        elif not isinstance(formulas[0], list):
            formulas = [formulas]
    
    return values, formats, formulas, base_row, base_col

def is_target_column(c):
    if c in DATE_COLUMNS:
        return True
    
def is_date_like(val):

    if isinstance(val, datetime):
        return True

    elif isinstance(val, (int, float)):
        return True  # Excelシリアルの可能性

    elif isinstance(val, str):
        try:
            datetime.strptime(val, "%Y/%m/%d")
            return True
        except:
            pass
    
    return False

def is_formula_cell(ws, base_row, r, base_col, c):
    cell_formula = ws.cells(base_row + r, base_col + c).formula
    return isinstance(cell_formula, str) and cell_formula.startswith("=")

def write_update_month_to_sheet(ws, base_row, r, base_col, c, val):
    if update_month(val) is not None:
        ws.cells(base_row + r, base_col + c).value = update_month(val)
        print(" ",update_month(val), "を入力")

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

def update_usage_text(ws):
    values = ws.used_range.value

    if not values:
        return

    for r, row in enumerate(values):
        for c, val in enumerate(row):
            if isinstance(val, str):
                old_val = val
                new_val = increment_year_month_text(val)
            
                if new_val != val:
                    ws.cells(r+1, c+1).value = new_val
                    print(" ", ws.name, "シート:", old_val, "→", new_val,"に更新")

def get_allcells_without_fmt(ws):
    ur = ws.range("A1:N400")
    values = ur.value
    formulas = ur.formula

    if not values:
        return None
    
    if not isinstance(values, list):
        values = [[values]]
    elif not isinstance(values[0], list):
        values = [values]

    rows = len(values)
    cols = len(values[0])
    
    if formulas is None:
        formulas = [[None] * cols for _ in range(rows)]
    else:
        if not isinstance(formulas, list):
            formulas = [[formulas]]
        elif not isinstance(formulas[0], list):
            formulas = [formulas]

    return {
        "values": values,
        "formulas": formulas,
        "base_row" : ur.row,
        "base_col" : ur.column
    }

def read_each_data_without_fmt(data):
    if not data:
        return None, None, None, None
    
    values = data.get("values")
    formulas = data.get("formulas")
    base_row = data.get("base_row")
    base_col = data.get("base_col")

    if not values:
        return None, None, None, None
    
    if not isinstance(values, list):
        values = [[values]]
    if not isinstance(values[0], list):
        values = [values]

    if formulas is None:
        formulas = [[None]*len(values[0]) for _ in range(len(values))]
    else:
        if not isinstance(formulas, list):
            formulas = [[formulas]]
        elif not isinstance(formulas[0], list):
            formulas = [formulas]
    
    return values, formulas, base_row, base_col

def is_project_sheet(ws):
    return "案件" in ws.name

def is_black_tab(ws):
    try:
        color = ws.api.Tab.Color
        return color == 0 and ws.api.Tab.ColorIndex != -4142
    except:
        return False

#================================
#メイン処理
#================================
app = xw.App(visible=False)
app.screen_updating = False
app.display_alerts = False

target_keywords = ["確定合意書", "DMM（秀", "御請求書",]
#SOURCE_SHEET = ["DMM（秀商）", "報酬額確定合意書"]
DATE_COLUMNS = [1,7]
DATE_COLUMNS_2 = [4, 10]

try:
    for file_name in os.listdir(input_folder):
        if not file_name.endswith((".xlsx", ".xlsm")) or file_name.startswith("~$"):
            continue
            
        file_path = os.path.join(input_folder, file_name)
        print(f"\n処理開始: {file_name}")

        wb = None

        #================================
        #メイン処理
        #================================

        try:
            wb = app.books.open(file_path)

            #================================
            #月の繰り上げ
            #================================

            data = {}
            print("【処理1】")
            
            for ws in wb.sheets:
                data = get_allcells_in_target_sheet(ws)

                if is_target_sheet(ws):
                    result = read_each_data(data)

                    if not result[0]:
                        continue

                    values, formats, formulas, base_row, base_col = result

                    if not values:
                        continue

                    if not isinstance(values, list):
                        continue

                    for r, row in enumerate(values):
                        for c, val in enumerate(row):
                            val = row[c] if c < len(row) else None

                            if val is None:
                                continue

                            if not is_target_column(c):
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

                            if not is_date_like(val):
                                continue

                            if is_formula_cell(ws, base_row, r, base_col, c):
                                continue

                            write_update_month_to_sheet(ws, base_row, r, base_col, c, val)

            #================================
            #2026年〇月利用分を更新
            #================================
            print("【処理2】")
            for ws in wb.sheets:
                update_usage_text(ws)

            #================================
            #n/24回目の更新
            #================================
            print("【処理3】")
            processed_cells = set()

            for ws in wb.sheets:
                data = get_allcells_without_fmt(ws)

                is_completed_sheet = False
                
                result = read_each_data_without_fmt(data)

                if not result[0]:
                    continue

                values, formulas, base_row, base_col = result

                if not values:
                    continue

                if formulas is None:
                    formulas = [[None]*len(values[0]) for _ in range(len(values))]

                if not isinstance(values, list):
                    continue
                if not isinstance(values[0], list):
                    values = [values]

                if not isinstance(formulas, list):
                    formulas = [formulas]
                if not isinstance(formulas[0], list):
                    formulas = [[f] for f in formulas]

                for r, row in enumerate(values):  #行を抜き出し
                    for c, val in enumerate(row):    #抜き出した行のセルを走査
                        if c not in DATE_COLUMNS_2:
                            continue

                        formula = None

                        if formulas and r < len(formulas) and c < len(formulas[r]):
                            formula = formulas[r][c]

                        if not isinstance(val, (str, datetime)):
                            continue
                        

                        if isinstance(val, str):  #セルの値value)はstr型(文字列)？
                            match = pattern_b.search(val)   #Matchオブジェクト

                            if match:   #matchの中身があるとTrueとして判定される。NoneだとFalse扱い
                                cell_key = (ws.name, r, c)
                                if cell_key in processed_cells:
                                    continue
                                
                                if formula is None:
                                    #fallback
                                        cell_formula = ws.cells(base_row + r, base_col + c).formula

                                        if isinstance(cell_formula, str) and cell_formula.startswith("="):
                                            continue
                                else:
                                    if isinstance(formula, str) and formula.startswith("="):
                                        continue

                                old_val = val

                                left = int(match.group(1))  #matchオブジェクトのmatch(1)、ここでは(/d+)に相当する部分
                                right = match.group(2)
                                right_num = int(right.replace("回目", ""))

                                if left == right_num:
                                    is_completed_sheet = True
                                else:
                                    new_left = left + 1
                                    text = f"{new_left}/{right}"
                                    result = re.sub(r"(\d+)/(\d+回目)", text, val)
                                    ws.cells(base_row + r, base_col + c).value = result   #f文字列g
                                    print("  更新完了:", ws.name, "シート", old_val,"→", result, "に更新")
                                    processed_cells.add(cell_key)

                if is_completed_sheet and is_project_sheet(ws) and not is_black_tab(ws):
                    ws.api.Tab.Color = 0
                    print(ws.name, "は完了状態 → タブ色を変更")
                    continue

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