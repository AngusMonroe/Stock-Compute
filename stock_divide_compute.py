import xlrd, xlwt
from xlutils.copy import copy
import datetime
import sys
import io

day_range = 20

def setup_io():
    sys.stdout = sys.__stdout__ = io.TextIOWrapper(sys.stdout.detach(), encoding='utf-8', line_buffering=True)
    sys.stderr = sys.__stderr__ = io.TextIOWrapper(sys.stderr.detach(), encoding='utf-8', line_buffering=True)
setup_io()

search_book1 = xlrd.open_workbook('创业板日度数据.xlsx')
search_sheet1 = search_book1.sheet_by_index(0)
rows1 = search_sheet1.nrows
cols1 = search_sheet1.ncols

search_book2 = xlrd.open_workbook('中小板日度数据.xlsx')
search_sheet2 = search_book2.sheet_by_index(0)
rows2 = search_sheet2.nrows
cols2 = search_sheet2.ncols

result_book = xlrd.open_workbook('查询结果.xlsx')
new_book = copy(result_book)
result_sheet = new_book.get_sheet(0)
refer_sheet = result_book.sheet_by_index(0)
result_rows = refer_sheet.nrows
result_cols = refer_sheet.ncols

def search(sheet, rows):
    sum = 0.0
    for r in range(rows):
        if r <= 1:
            continue
        r_value = sheet.row_values(r)
        # print(r_value)
        search_date = datetime.datetime.strptime(r_value[0],"%Y-%m-%d")
        if aim_date <= search_date:
            for i in range(r - day_range, r + day_range + 1):
                val = sheet.row_values(i)[1]
                # print(val)
                sum += float(val)
            break
    return sum

for row in range(result_rows):
    print(row)
    try:
        if row <= 1:
            continue
        # if row > 2:
        #     break
        row_value = refer_sheet.row_values(row)
        # print(row_value)
        aim_date = datetime.datetime.strptime(row_value[3],"%Y-%m-%d")  # 字符串转化为date形式
        if int(row_value[0].split('.')[0]) >= 300000:
            ans = search(search_sheet1, rows1)
        else:
            ans = search(search_sheet2, rows2)
        result_sheet.write(row, 10, ans)
    except Exception:
        print('ERROR')
        continue
new_book.save('result.xls')


print('OK')
