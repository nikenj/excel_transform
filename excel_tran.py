# -*- coding: <utf-8> -*- 
import xlwings as xw
import os
import arrow
import shutil
import sys

times = input("please input a need statistic date (format like 201801) :  ")
print(times)
base_path = 'E:/学习资料/use_excel/tables'

current_time = arrow.get(times,'YYYYMM')
current_month = current_time.month
current_year = current_time.year
last_time = current_time.shift(months=-1)
last_year = last_time.year
last_month = last_time.month
tow_month_ago = current_time.shift(months=-2)

current_time_str = "%s年%s月"%(current_year,current_month)
current_month_str = "%s月"%(current_month)
print(current_time_str,current_month_str)

last_time_str = "%s年%s月"%(last_year,last_month)
last_month_str = "%s月"%(last_month)
last_month_ago_str = "%s月份水电"%(last_month)
last_month_ago_time_str = last_time.format('YYYYMM')
tow_month_ago_str = "%s月份水电"%(tow_month_ago.month)
tow_month_ago_time_str = tow_month_ago.format('YYYYMM')
print(tow_month_ago_str,last_month_ago_str,last_month_ago_time_str)

def excel_transform(excel_path):
    old_excel_name = last_time_str+".xls"
    new_excel_name = current_time_str+".xls"
    in_put_excel_file = base_path + "/" + old_excel_name
    print(in_put_excel_file)
    out_put_excel_file = base_path + "/" + new_excel_name
    shutil.copyfile(in_put_excel_file,out_put_excel_file,)
    app2 = xw.App(visible=False,add_book=False)
    wb2 =app2.books.open(out_put_excel_file)
#     wb2 = xw.Book(out_put_excel_file)
    for sheet in wb2.sheets:
        sht = wb2.sheets[sheet]
        if not sht.name.startswith("兴华宇"):
            print("%s is doing"%sht.name)
            sht.range('C5:C30').options(transpose=True).value = sht.range('D5:D30').value
            for i in sht.range('B1:B50'):
                old_str = str(i.value)
                if last_time_str in old_str:
                    new_str_1 = old_str.replace(last_time_str,current_time_str)
                    new_str_3 = new_str_1.replace(tow_month_ago_str,last_month_ago_str)
                    i.options(transpose=True).value = new_str_3              
                elif last_month_str in old_str:
                    new_str_2 = old_str.replace(last_month_str,current_month_str)
                    i.options(transpose=True).value = new_str_2
        else:
            sht.name = sht.name.replace(tow_month_ago_time_str,last_month_ago_time_str)
            old_str = str(sht.range('A1').value)
            new_str_2 = old_str.replace(tow_month_ago_time_str,last_month_ago_time_str)
            sht.range('A1').options(transpose=True).value = new_str_2
            print(sht.name)
    wb2.save()
    wb2.close()
    app2.quit()

excel_transform(base_path)