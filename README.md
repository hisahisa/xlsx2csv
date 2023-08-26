# xlsx2csv
使い方

```python
import pandas as pd
import zipfile
import io
import time
import xlsx2csv

# 開始
start_time = time.perf_counter()

file_name = 'tenpo_shohin_pattern3.xlsx'
excel_sheet = 'sheet1'

with open(file_name, 'rb') as file_obj:
    with zipfile.ZipFile(file_obj, 'r') as zip_ref:
        # エクセルファイルを展開してオブジェクトを取得
        s_list = [f'xl/worksheets/{excel_sheet}.xml']
        sheet_list = [i for i in zip_ref.namelist() if i in s_list]
        solve_bytes = 'xl/sharedStrings.xml'
        bytes_obj = zip_ref.read([i for i in zip_ref.namelist() if i in sheet_list][0])
        bytes_solve_obj = zip_ref.read(solve_bytes)
        str_csv = xlsx2csv.read_excel(bytes_obj, bytes_solve_obj)
        df_generator = pd.read_csv(io.StringIO(str_csv), chunksize=10000)
        for df in df_generator:
            print(df)

end_time = time.perf_counter()
# 経過時間を出力(秒)
elapsed_time = end_time - start_time
print(elapsed_time)

```