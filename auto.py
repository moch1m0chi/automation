import os
from datetime import datetime
import xlwings as xw
import re
import calendar
    
input_folder = "data"
output_folder = "output"
os.makedirs(output_folder, exist_ok=True)

#正規表現(数値)/(数値)回目をコンパイル
pattern = re.compile(r"(\d+)/(\d+回目)")

#ファイル名の月繰り上げ
def increment_month_in_filename(file_name):
    match = re.search(r"(\d+)月", file_name)
    
    if match:
        month = int(match.group(1))
        new_month = month + 1
        
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
    try:
        return dt.replace(year=year, month = month)
    except:
        last_day = calendar.monthrange(year, month)[1]
        return dt.replace(year = year, month= month, day = last_day)

#================================
#メイン処理
#================================
app = xw.App(visible=False)
app.screen_updating = False
app.display_alerts = False

for file_name in os.listdir(input_folder):
    if not file_name.endswith((".xlsx", ".xlsm")) or file_name.startswith("~$"):
        continue
        
    file_path = os.path.join(input_folder, file_name)
    print(f"\n処理開始: {file_name}")

    app = None
    wb = None

    try:
        app = xw.App(visible=False)
        app.screen_updating = False

        try:
            wb =app.books.open(file_path)
        except Exception as e:
            print(f"開けない: {file_name}")
            print(file_path)
            print(e)
            continue

        
        if "報酬額確定合意書" in [s.name for s in wb.sheets]:
            ws = wb.sheets["報酬額確定合意書"]

            #範囲限定
            values = ws.range("A1:C50").value

            for r,row in enumerate(values):
                for c, val in enumerate(row):
                    if isinstance(val, datetime):
                        values[r][c] = add_one_month(val)
                
                    elif isinstance(val, str):
                        try:
                            dt = datetime.strptime(val, "%Y/%m/%d")
                            values[r][c] = add_one_month(dt).strftime("%Y/%m/%d")
                        except:
                            pass
            ws.range("A1").value = values        

            #================================
            #n/24回目の更新
            #================================
            for ws in wb.sheets:  # 全シート対象
                values = ws.range("A1:N400").value

                for r, row in enumerate(values):  #行を抜き出し
                    for c, val in enumerate(row):    #抜き出した行のセルを走査
                        if isinstance(val, str):  #セルの値value)はstr型(文字列)？
                            match = pattern.search(val)   #Matchオブジェクト
                            if match:   #matchの中身があるとTrueとして判定される。NoneだとFalse扱い
                                left = int(match.group(1)) + 1  #matchオブジェクトのmatch(1)、ここでは(/d+)に相当する部分
                                right = match.group(2)

                                values[r][c] = f"{left}/{right}"   #f文字列
                            #print(f"{ws.name} {cell.address}: {value} → {new_value}")
                ws.range("A1").value = values

        new_file_name = increment_month_in_filename(file_name)  #ファイル名の月を繰り上げ
        output_path = os.path.join(output_folder, new_file_name)

        wb.save(output_path)
        print(f"保存完了:{new_file_name}")

    except Exception as e:
        print(f"エラー発生:{file_name}")
        print(f"内容:{e}")

    finally:
        if wb:
            wb.close()
        if app:
            app.quit()

print("\n全処理完了！")