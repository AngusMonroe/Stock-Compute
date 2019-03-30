import xlrd, xlwt
from xlutils.copy import copy
import datetime
import sys
import io

day_range = 30

def setup_io():
    sys.stdout = sys.__stdout__ = io.TextIOWrapper(sys.stdout.detach(), encoding='utf-8', line_buffering=True)
    sys.stderr = sys.__stderr__ = io.TextIOWrapper(sys.stderr.detach(), encoding='utf-8', line_buffering=True)
setup_io()

search_book = xlrd.open_workbook('股票历史收益率查询清单.xlsx')
search_sheet = search_book.sheet_by_index(0)
rows = search_sheet.nrows
cols = search_sheet.ncols

result_book = xlrd.open_workbook('查询结果.xlsx')
new_book = copy(result_book)
result_sheet = new_book.get_sheet(0)
refer_sheet = result_book.sheet_by_index(0)
result_rows = refer_sheet.nrows
result_cols = refer_sheet.ncols

for row in range(result_rows):
    print(row)
    try:
        if row == 0:
            continue
        # if row > 2:
        #     break
        row_value = refer_sheet.row_values(row)
        aim_date = datetime.datetime.strptime(row_value[3],"%Y-%m-%d")  # 字符串转化为date形式

        for r in range(rows):
            if r == 0:
                continue
            r_value = search_sheet.row_values(r)
            search_date = datetime.datetime.strptime(r_value[2],"%Y-%m-%d")
            if row_value[1] == r_value[1] and aim_date <= search_date:
                sum = 0.0
                for i in range(r - day_range, r + day_range + 1):
                    val = search_sheet.row_values(i)[4]
                    # print(val)
                    sum += val
                break
        result_sheet.write(row, 6, sum)
    except Exception:
        print('ERROR')
        continue

new_book.save('result.xls')

print('OK')
