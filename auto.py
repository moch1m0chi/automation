import os
from openpyxl import load_workbook
import re

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
    
#元データを入れるフォルダの指定
input_folder = "data"
#処理したファイルの出力先フォルダを指定、なければoutputフォルダを作る
output_folder = "output"
os.makedirs(output_folder, exist_ok=True)

#正規表現(数値)/(数値)回目をコンパイル
pattern = re.compile(r"(\d+)/(\d+回目)")

#n/24回目を探してn+1/24回目に書き換える
for file_name in os.listdir(input_folder):
    if file_name.endswith((".xlsx", ".xlsm")):
        file_path = os.path.join(input_folder, file_name)
        print(f"処理中: {file_name}")

        wb = load_workbook(file_path)

        for ws in wb.worksheets:  # 全シート対象
            for row in ws.iter_rows():
                for cell in row:
                    value = cell.value

                    if isinstance(value, str):
                        match = pattern.search(value)
                        if match:
                            left = int(match.group(1)) + 1
                            right = match.group(2)

                            new_value = f"{left}/{right}"
                            print(f"{ws.title} {cell.coordinate}: {value} → {new_value}")

                            cell.value = new_value

        new_file_name = increment_month_in_filename(file_name)

        output_path = os.path.join(output_folder, new_file_name)
        wb.save(output_path)

print("完了！")