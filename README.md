# xlsx2csv
使い方
```asm
import pandas as pd
import zipfile
import io
import time
import xlsx2csv

# 開始
start_time = time.perf_counter()

with zipfile.ZipFile('tenpo_shohin_pattern3.xlsx', 'r') as zip_ref:
    # エクセルファイルを展開してオブジェクトを取得
    f_list = ['xl/worksheets/sheet1.xml']
    sheet_list = [i for i in zip_ref.namelist() if i in f_list]
    str_resolve_file = 'xl/sharedStrings.xml'
    file_obj = zip_ref.read([i for i in zip_ref.namelist() if i in sheet_list][0])
    str_resolve_obj = zip_ref.read(str_resolve_file)
    x = xlsx2csv.read_excel(file_obj, str_resolve_obj)
    r = pd.read_csv(io.StringIO(x), chunksize=10000)
    for i in r:
        print(i)


end_time = time.perf_counter()
# 経過時間を出力(秒)
elapsed_time = end_time - start_time
print(elapsed_time)

```