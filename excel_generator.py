import os
from datetime import datetime, timedelta
import xlwings as xw
import re
import calendar
from pathlib import Path
import argparse
    
#================================
#ディレクトリ指定
#================================
# base_dir = Path(__file__).resolve().parent
# input_folder = base_dir / "data"
# output_folder = base_dir / "output"
# output_folder.mkdir(parents=True, exist_ok=True)

#正規表現コンパイル
pattern_a = re.compile(r"(\d{4})年(\d{1,2})月利用分")
pattern_b = re.compile(r"(\d+)/(\d+回目)")
pattern_c = re.compile(r"(\d{4})年(\d{1,2})月末")

#================================
#インフラ初期化
#================================
def init_excel_app():
    app = xw.App(visible=False)
    app.screen_updating = False
    app.display_alerts = False
    return app

#================================
#クラス定義
#================================

class ExcelProcessor:
    def __init__(self, wb):
        self.wb = wb
        self.processed_cells = set()

        self.target_keywords = ["確定合意書", "DMM（秀", "御請求書","支払", "ガソリン販売"]
        self.DATE_COLUMNS = [1, 6, 7, 9]
        self.DATE_COLUMNS_2 = [4, 10]

    #================================
    #エントリーポイント
    #================================
    def run(self):
        self.log("【処理1】") # 処理1 : 確定合意書シート等の日付セルを次の月に繰り上げる
        self.process_sheets(
            self.is_target_sheet, # プロセス実行の条件
            self.update_month_on_sheets # 条件を満たすとき実行
        )

        self.log("【処理2】") # 処理2 : ガソリン代シート等の「YYYY年M月分」の月を次の月に繰り上げる
        self.process_sheets(
            lambda ws: True, 
            self.update_usage_text
        )

        self.log("【処理3】") # 処理3 : 案件シート等の「n/24回目」の形式の回数をひとつ繰り上げる
        def action_counts(ws):
            is_completed = self.update_counts_on_sheets(ws, self.processed_cells)
            if is_completed and self.is_project_sheet(ws) and not self.is_black_tab(ws):
                self.change_tab_color(ws)

        self.process_sheets(
            lambda ws: True, 
            action_counts
        )

        self.log("【処理4】") #処理4 : 支払予定日の月を次の月に繰り上げる
        self.process_sheets(
            lambda ws: True, 
            self.update_payday
        )
    
    #================================
    #制御系
    #================================
    def process_sheets(self, condition_func, action_func): # Excelブック中の各シートについて条件を満たしたらアクションを行う
        for ws in self.wb.sheets:
            if condition_func(ws):
                action_func(ws)
        
    #================================
    #メインロジック
    #================================
    def update_month_on_sheets(self,ws): # 処理1のメインプロセス

        if self.is_target_sheet(ws):
            data = self.get_sheet_matrix(ws)
            values, formats, formulas, base_row, base_col = data
            if not data:
                return

            if not values:
                return

            if not isinstance(values, list):
                return

            for r, row in enumerate(values): 
                for c, val in enumerate(row):
                    self.process_month_update(ws, r, c, val, data)

    def update_usage_text(self, ws):
        data = self.get_sheet_matrix(ws)
        if not data:
            return

        values, formats, formulas, base_row, base_col = data

        for r, row in enumerate(values):
            for c, val in enumerate(row):
                if isinstance(val, str):
                    old_val = val
                    new_val = self.increment_year_month_text(val)
                
                    if new_val != val:
                        self.write_update_usage_text(ws, old_val, new_val, base_row, base_col, r, c)

    def update_counts_on_sheets(self, ws, processed_cells):
        sheetdata = self.get_allcells_without_fmt(ws)

        is_completed_sheet = False
        
        data = self.read_each_data_without_fmt(sheetdata)

        if not data[0]:
            return False

        for r, row in enumerate(sheetdata["values"]):  #行を抜き出し
            for c, val in enumerate(row):    #抜き出した行のセルを走査
        
                values, formulas, base_row, base_col = data

                if c not in self.DATE_COLUMNS_2:
                    continue

                formula = self.normalized_formula(formulas, r, c)

                if not isinstance(val, (str, datetime)):
                    continue
                
                if isinstance(val, str):  #セルの値value)はstr型(文字列)？
                    match = pattern_b.search(val)   #Matchオブジェクト

                    if match:   #matchの中身があるとTrueとして判定される。NoneだとFalse扱い
                        cell_key = (ws.name, r, c)
                        
                        old_val = val
                        new_val = self.transform_count_text(val, match)

                        if cell_key in self.processed_cells:
                            continue
                        
                        if self.is_like_formula(formula, ws, base_row, base_col, r, c):
                            continue

                        if self.is_already_finished(match):
                            is_completed_sheet = True
                        else:
                            self.write_update_counts_to_sheet(ws, base_row, base_col, r, c, old_val, new_val)
        
        return is_completed_sheet

    def update_payday(self, ws):
        data = self.get_sheet_matrix(ws)
        if not data:
            return

        values, formats, formulas, base_row, base_col = data

        for r, row in enumerate(values):
            for c, val in enumerate(row):

                if not self.should_update_payday(values, r, c, val):
                    continue

                formula = self.normalized_formula(formulas, r, c)
                if self.is_like_formula(formula, ws, base_row, base_col, r, c):
                    continue

                new_val = self.increment_payday(val)
                old_val = val
                if new_val == val:
                    continue

                self.write_update_payday(ws, old_val, new_val, base_row, base_col, r, c)

    #================================
    #判定系
    #================================
    def is_target_sheet(self, ws): # 更新対象のシートか判定
        return any(k in ws.name for k in self.target_keywords)
    
    def is_target_column(self, c): # データの入っている列か判定。エクセル内部では日付も普通の数値とされるため
        return c in self.DATE_COLUMNS
    
    def is_date_like(self, val): # 日付データっぽいか。日付型だけでなくstr型で入力されている場合の対策

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
    
    def is_formula_cell(self, ws, base_row, base_col, r, c): # 変数セルか判定。セル参照している場合に参照先+該当セルで二重に日付が変更されることへの対策
        cell_formula = ws.cells(base_row + r, base_col + c).formula
        return isinstance(cell_formula, str) and cell_formula.startswith("=")
    
    def is_project_sheet(self, ws): # シート名に案件の文字があるか
        return "案件" in ws.name
    
    def is_black_tab(self, ws): # シートのタブが黒=更新対象外か
        try:
            color = ws.api.Tab.Color
            return color == 0 and ws.api.Tab.ColorIndex != -4142
        except Exception as e:
            self.log(f"エラー内容: {e}")
            return False
        
    def is_like_formula(self, formula, ws, base_row, base_col, r, c): # 変数セルっぽいか判定isformulacellでうまくいかなかったので実装、いずれ統一したい
        if formula is None:
            #fallback 
            cell_formula = ws.cells(base_row + r, base_col + c).formula
            return isinstance(cell_formula, str) and cell_formula.startswith("=")
        
        else:
            return isinstance(formula, str) and formula.startswith("=")
        
    def is_already_finished(self, match): # 対象の案件シートの回数n/24回目は終了しているか
        left, right = self.get_count_in_cell(match)
        right_num = int(right.replace("回目", ""))
        return left == right_num

    def is_year_month_like(self, val):
        return self.parse_year_month(val) is not None

    def should_update_payday(self, values, r, c, val):
        if not isinstance(val, str):
            return False

        if r == 0:
            return False

        upper_val = values[r-1][c]

        return self.is_year_month_like(upper_val)

    #================================
    #変換系
    #================================

    def normalized_fmt(self, formats, r, c): #フォーマット標準化
        fmt = ""
        if r < len(formats) and c < len(formats[r]):
            fmt = str(formats[r][c]).lower()
        return fmt

    def normalized_formula(self, formulas, r, c): # 数式標準化
        formula = None
        if r < len(formulas) and c < len(formulas[r]):
            formula = formulas[r][c]
        return formula
    
    def transform_date_and_month(self, val): # 月日を変換
        new_val = None

        if isinstance(val, datetime):
            new_val = self.add_one_month(val)
            return new_val

        elif isinstance(val, str):
            try:
                dt = datetime.strptime(val, "%Y/%m/%d")
                new_val = self.add_one_month(dt)
                return new_val
            except:
                return

        else:
            return
    
    def transform_month_in_filename(self, file_name: str) -> str: # ファイル名の月を変換
        match = re.search(r"(\d+)月", file_name)

        if not match:
            return file_name

        month = int(match.group(1))
        new_month = self.increment_month(month)

        return re.sub(r"\d+月", f"{new_month}月", file_name)
    
    def transform_count_text(self, val: str, match): #n/24回目のテキストを変換
        left, right = self.get_count_in_cell(match)
        new_left, right = self.increment_count(left, right)

        text = f"{new_left}/{right}"
        return re.sub(r"(\d+)/(\d+回目)", text, val)
    
    def transform_payday(self, val): # 支払日を変換
        return self.increment_payday(val)
    
    def parse_year_month(self, val):
        # ① datetime
        if isinstance(val, datetime):
            return val.year, val.month

        # ② Excelシリアル値
        if isinstance(val, (int, float)):
            if not (1 <= val <= 60000):
                return None

            try:
                dt = datetime(1899, 12, 30) + timedelta(days=val)
                return dt.year, dt.month
            except:
                return None

        # ③ 文字列（2024年5月）
        if isinstance(val, str):
            m = re.search(r"(\d{4})年(\d{1,2})月", val)
            if not m:
                return None

            year = int(m.group(1))
            month = int(m.group(2))

            if not (1 <= month <= 12):
                return None

            return year, month

        return None
        
    #================================
    #書き込み系
    #================================
    def write_cell(self, ws, base_row, base_col, r, c, new_val): # Excelのセルに直接書き込み
        ws.cells(base_row + r, base_col + c).value = new_val
    
    def write_update_month_to_sheet(self, ws, base_row, base_col, r, c, val): #更新した月を
        if self.transform_date_and_month(val) is not None:
            new_val = self.transform_date_and_month(val)
            self.write_cell(ws, base_row, base_col, r, c, new_val)
            self.log(f" {new_val}を入力")

    def process_month_update(self, ws, r, c, val, data): # 判定式を適用
        values, formats, formulas, base_row, base_col = data
        # val = row[c] if c < len(row) else None

        if val is None:
            return

        if not self.is_target_column(c):
            return

        if self.normalized_fmt(formats, r, c):
            return
        # if "yy" not in self.normalized_fmt(formats, r, c):
        #     return

        if self.normalized_formula(formulas, r, c):
            return

        if not isinstance(val, (datetime, int, float, str)):
            return

        if not self.is_date_like(val):
            return

        if self.is_formula_cell(ws, base_row, base_col, r, c):
            return

        self.write_update_month_to_sheet(ws, base_row, base_col, r, c, val)

    def write_update_counts_to_sheet(self, ws, base_row, base_col, r, c, old_val, new_val):
        self.write_cell(ws, base_row, base_col, r, c, new_val)
        self.log(f"  更新完了: {ws.name} シート {old_val} → {new_val} に更新")

    def write_update_payday(self, ws, old_val, new_val, base_row, base_col, r, c):
        self.write_cell(ws, base_row, base_col, r, c, new_val)
        self.log(f"  更新完了: {ws.name} シート {old_val} → {new_val} に更新")

    def write_update_usage_text(self, ws, old_val, new_val, base_row, base_col, r, c):
        self.write_cell(ws, base_row, base_col, r, c, new_val)
        self.log(f"  更新完了: {ws.name} シート {old_val} → {new_val} に更新")
    
    #================================
    #ユーティリティ
    #================================
    
    def get_sheet_matrix(self, ws): # シートの行列を取得
        sheetdata = self.get_allcells(ws)
        if not sheetdata:
            return None
        return self.read_each_data(sheetdata)
    
    def get_allcells(self, ws): # シートの全セルを取得
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
    
    def get_allcells_in_target_sheet(self, ws): # 対象セルのセルをすべて取得、月更新用に使用
        if not self.is_target_sheet(ws):
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

    def read_each_data(self, sheetdata): # 各データの読み込み
        if not sheetdata:
            return None, None, None, None, None
        
        values = sheetdata.get("values")
        formats = sheetdata.get("formats")
        formulas = sheetdata.get("formulas")
        base_row = sheetdata.get("base_row")
        base_col = sheetdata.get("base_col")

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

    def get_allcells_without_fmt(self, ws): #フォーマット形式以外のセルをすべて取り込み、今思えば別にいらなかった
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

    def read_each_data_without_fmt(self, data): #フォーマット形式以外のデータをすべて読み込み、今思えば別にいらなかった
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

    def get_count_in_cell(self, match): # セルの回数を取得、leftがn、rightが/24回目
        left = int(match.group(1))  #matchオブジェクトのmatch(1)、ここでは(/d+)に相当する部分
        right = match.group(2)
        return left, right

    def change_tab_color(self, ws): #シートタブ色を変更、案件シートが終了している際に使用する
        ws.api.Tab.Color = 0
        self.log(f"{ws.name} は完了状態 → タブ色を変更")

 
    def add_one_month(self, dt): # ひと月繰り上げ、セルの日付が月末であれば更新後も月末日を維持する
        year = dt.year
        month = dt.month + 1
        if month > 12:
            month = 1
            year += 1

        # 元の日付が月末か判定
        last_day_current = calendar.monthrange(dt.year, dt.month)[1]
        is_month_end = dt.day == last_day_current

        # 次の月の末日
        last_day_next = calendar.monthrange(year, month)[1]

        if is_month_end:
            day = last_day_next  # ← 月末なら次も月末
        else:
            day = min(dt.day, last_day_next)

        return dt.replace(year=year, month=month, day=day)
    
    def increment_year_month_text(self, text):
        def repl(match):
            year = int(match.group(1))
            month = int(match.group(2))

            month += 1
            if month > 12:
                month = 1
                year += 1

            return f"{year}年{month}月利用分"

        return pattern_a.sub(repl, text)
    
    def increment_payday(self, text):
        def repl(match):
            year = int(match.group(1))
            month = int(match.group(2))

            month += 1
            if month > 12:
                month = 1
                year += 1

            return f"{year}年{month}月末"
        return pattern_c.sub(repl, text)
    
    def increment_month(self, month: int) -> int:
        month += 1
        if month > 12:
            month = 1
        return month
    
    def increment_count(self, left: int, right: str):
        new_left = left + 1
        return new_left, right
    
    def log(self, msg):
        print(msg)

    def save_excel(self, file_name, output_folder, wb):
        new_file_name = self.transform_month_in_filename(file_name)  #ファイル名の月を繰り上げ
        output_path = os.path.join(output_folder, new_file_name)

        wb.save(output_path)
        self.log(f"保存完了 : {new_file_name}")

    #GUI化用(未実装)
    def run_job(input_folder, output_folder, log_func=None):
    
    #     app = xw.App(visible=False)
    #     app.screen_updating = False
    #     app.display_alerts = False

    #     try:
    #         for file_name in os.listdir(input_folder):
    #             if not file_name.endswith((".xlsx", ".xlsm")) or file_name.startswith("~$"):
    #                 continue
                    
    #             file_path = os.path.join(input_folder, file_name)

    #             if log_func:
    #                     log_func(f"処理開始: {file_name}")


    #             wb = None

    #             #================================
    #             #メイン処理
    #             #================================

    #             try:
    #                 wb = app.books.open(file_path)
    #                 data = {}

    #                 #================================
    #                 #月の繰り上げ
    #                 #================================

    #                 print("【処理1】")
                    
    #                 for ws in wb.sheets:
    #                     sheetdata = get_allcells_in_target_sheet(ws)

    #                     if is_target_sheet(ws):
    #                         data = read_each_data(sheetdata)

    #                         if not sheetdata:
    #                             continue

    #                         if not sheetdata["values"]:
    #                             continue

    #                         if not isinstance(sheetdata["values"], list):
    #                             continue

    #                         for r, row in enumerate(sheetdata["values"]):
    #                             for c, val in enumerate(row):
    #                                 process_month_update(ws, r, c, val, data)

    #                 #================================
    #                 #2026年〇月利用分を更新
    #                 #================================
    #                 print("【処理2】")

    #                 for ws in wb.sheets:
    #                     update_usage_text(ws)

    #                 #================================
    #                 #n/24回目の更新
    #                 #================================
    #                 print("【処理3】")

    #                 processed_cells = set()

    #                 for ws in wb.sheets:
    #                     data = get_allcells_without_fmt(ws)

    #                     is_completed_sheet = False
                        
    #                     result = read_each_data_without_fmt(data)

    #                     if not result[0]:
    #                         continue

    #                     values, formulas, base_row, base_col = result

    #                     if not values:
    #                         continue

    #                     if formulas is None:
    #                         formulas = [[None]*len(values[0]) for _ in range(len(values))]

    #                     if not isinstance(values, list):
    #                         continue
    #                     if not isinstance(values[0], list):
    #                         values = [values]

    #                     if not isinstance(formulas, list):
    #                         formulas = [formulas]
    #                     if not isinstance(formulas[0], list):
    #                         formulas = [[f] for f in formulas]

    #                     for r, row in enumerate(values):  #行を抜き出し
    #                         for c, val in enumerate(row):    #抜き出した行のセルを走査
    #                             if c not in DATE_COLUMNS_2:
    #                                 continue

    #                             formula = None

    #                             if formulas and r < len(formulas) and c < len(formulas[r]):
    #                                 formula = formulas[r][c]

    #                             if not isinstance(val, (str, datetime)):
    #                                 continue
                                

    #                             if isinstance(val, str):  #セルの値value)はstr型(文字列)？
    #                                 match = pattern_b.search(val)   #Matchオブジェクト

    #                                 if match:   #matchの中身があるとTrueとして判定される。NoneだとFalse扱い
    #                                     cell_key = (ws.name, r, c)
    #                                     if cell_key in processed_cells:
    #                                         continue
                                        
    #                                     if formula is None:
    #                                         #fallback
    #                                             cell_formula = ws.cells(base_row + r, base_col + c).formula

    #                                             if isinstance(cell_formula, str) and cell_formula.startswith("="):
    #                                                 continue
    #                                     else:
    #                                         if isinstance(formula, str) and formula.startswith("="):
    #                                             continue

    #                                     old_val = val

    #                                     left = int(match.group(1))  #matchオブジェクトのmatch(1)、ここでは(/d+)に相当する部分
    #                                     right = match.group(2)
    #                                     right_num = int(right.replace("回目", ""))

    #                                     if left == right_num:
    #                                         is_completed_sheet = True
    #                                     else:
    #                                         new_left = left + 1
    #                                         text = f"{new_left}/{right}"
    #                                         result = re.sub(r"(\d+)/(\d+回目)", text, val)
    #                                         ws.cells(base_row + r, base_col + c).value = result   #f文字列g
    #                                         print("  更新完了:", ws.name, "シート", old_val,"→", result, "に更新")
    #                                         processed_cells.add(cell_key)

    #                     if is_completed_sheet and is_project_sheet(ws) and not is_black_tab(ws):
    #                         ws.api.Tab.Color = 0
    #                         print(ws.name, "は完了状態 → タブ色を変更")
    #                         continue

    #                 #================================
    #                 #Excelファイルの保存
    #                 #================================
    #                 save_excel(file_name, output_folder, wb)

    #             except Exception as e: 
    #                 if log_func:
    #                     log_func(f"エラー: {file_name} / {e})
    #             finally: 
    #                 try:
    #                     if wb: 
    #                         wb.close()
    #                 except:
    #                     print("えらー")
    #                     pass
    #     finally:
    #         app.quit()

    #     ("\n全処理完了！")
        pass

#================================
#アプリケーション制御
#================================
def run_batch(input_folder, output_folder):
    app = init_excel_app()

    try:
        for file_path in input_folder.iterdir():
            file_name = file_path.name
            if not file_name.endswith((".xlsx", ".xlsm")) or file_name.startswith("~$"):
                continue
                
            file_path = os.path.join(input_folder, file_name)

            print(f"\n処理開始: {file_name}")
            wb = None

            try:
                wb = app.books.open(file_path)
                data = {}

                processor = ExcelProcessor(wb)
                processor.run()

                processor.save_excel(file_name, output_folder, wb)
                # processor.export_pdf(file_name, output_folder)

            except Exception as e: #エラー時のメッセージ表示
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



def parse_args():
    parser = argparse.ArgumentParser()

    parser.add_argument("--input", default="data")
    parser.add_argument("--output", default="output")

    return parser.parse_args()

#================================
#エントリーポイント
#================================

def main():
    args = parse_args()

    base_dir = Path(__file__).resolve().parent
    input_folder = base_dir / args.input
    output_folder = base_dir / args.output
    output_folder.mkdir(parents=True, exist_ok=True)

    run_batch(input_folder, output_folder)

if __name__ == "__main__":
    main()