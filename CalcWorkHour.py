import os
import sys
from os.path import join
import xlwings as xw
from decimal import Decimal, ROUND_HALF_UP
import pandas as pd
import yaml
import re

cur_path = os.getcwd()
g_row_num = 1
g_row_num_error = 2
g_data_num = 0
wb = 0


def is_file_locked(file_path):
    if os.path.exists(file_path):
        try:
            # 尝试以独占模式打开文件
            with open(file_path, 'a'):
                return False  # 文件未被锁定
        except IOError:
            return True  # 文件被锁定
    return False  # 文件不存在


def format_decimal(value, precision=2):
    return Decimal(value).quantize(Decimal('0.00'), rounding=ROUND_HALF_UP)


def excel():
    global g_row_num, wb
    global g_row_num_error
    global g_data_num
    new_path = join(cur_path, '工时统计.xlsx')

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    if not os.path.exists(new_path):
        wb = app.books.add()
    else:
        if is_file_locked(new_path):
            print("文件正在被其他程序占用，请关闭文件后重试。\n")
            print("按下Enter键后程序将关闭!")
            input()
            return

        wb = app.books.open(new_path)

    sheet = wb.sheets["研发中心日报"]
    sheet.range("A:A").api.NumberFormat = "@"
    sheet.range("B:B").api.NumberFormat = "@"
    sheet.range("G:G").api.NumberFormat = "@"
    rows = sheet.used_range.last_cell.row

    # 补齐数据
    for i in range(3, rows + 1):
        if sheet.range((i, 3)).value is None:
            if sheet.range((i, 8)).value is not None and sheet.range((i, 9)).value is not None:
                sheet.range((i, 3)).value = sheet.range((i - 1, 3)).value
                sheet.range((i, 7)).value = sheet.range((i - 1, 7)).value
            else:
                wb.close()
                app.quit()
                print("表格中存在无数据行! 错误行为%d \n" % i)
                print("按下Enter键后程序将关闭!")
                input()
                return

    # 获取所有sheet的名称列表
    sheet_names = [sheet.name for sheet in wb.sheets]

    # 要检查的sheet名称
    sheet_name_to_check1 = '月单位人工项目比例统计'
    sheet_name_to_check2 = '工时填写错误统计'

    # 判断sheet是否存在
    if sheet_name_to_check1 in sheet_names:
        # 如果存在，获取sheet并进行操作
        calc_sheet = wb.sheets[sheet_name_to_check1]
        calc_sheet.clear_contents()
    else:
        calc_sheet = wb.sheets.add("月单位人工项目比例统计")

    calc_sheet.range('A1:E1').value = ['姓名', '项目名称', '项目缩写 ', '工时比例', '工时处理']
    calc_sheet.range('A1:E1').api.Font.Bold = True
    # 设置单元格自动换行
    calc_sheet.range('B:B').api.WrapText = True
    var = calc_sheet.range('B2').column_width
    calc_sheet.range('B:B').column_width = 20 + var

    calc_sheet.range('C:C').api.WrapText = True
    var = calc_sheet.range('C2').column_width
    calc_sheet.range('C:C').column_width = 8 + var

    # 判断sheet是否存在
    if sheet_name_to_check2 in sheet_names:
        # 如果存在，获取sheet并进行操作
        error_sheet = wb.sheets[sheet_name_to_check2]
        error_sheet.clear_contents()
    else:
        error_sheet = wb.sheets.add("工时填写错误统计")

    error_sheet.range('A1').value = '填写人'
    error_sheet.range('B1').value = '项目'
    error_sheet.range('C1').value = '日期'
    error_sheet.range('D1').value = '工时比例'
    error_sheet.range('A1:D1').api.Font.Bold = True
    error_sheet.range('B1').api.HorizontalAlignment = -4108
    # 设置单元格自动换行
    error_sheet.range('B:B').api.WrapText = True
    var = error_sheet.range('B2').column_width
    error_sheet.range('B:B').column_width = 20 + var
    error_sheet.range('C:C').column_width = 5 + var

    error_sheet.autofit(axis='rows')

    data_dict = {}
    workhour_errors = []

    name_path = join(cur_path, 'names.yaml')
    with open(name_path, 'r', encoding='utf-8') as file:
        data = yaml.safe_load(file)

    target_names = data['names']

    for i in range(3, rows + 1):
        name = sheet.cells(i, 3).value
        date = sheet.cells(i, 7).value
        project = sheet.cells(i, 8).value
        work_hour = sheet.cells(i, 9).value

        if name in target_names:
            if (name, date) not in data_dict:
                data_dict[(name, date)] = {}
            if project not in data_dict[(name, date)]:
                data_dict[(name, date)][project] = 0
            data_dict[(name, date)][project] += float(work_hour)

    name_error = []

    for (name, date), projects in data_dict.items():
        total_work_hour = sum(projects.values())
        if total_work_hour <= 0.99 or total_work_hour >= 1.01:
            workhour_errors.append((name, date))
            name_error.append(name)

    for name, date in workhour_errors:
        for i in range(3, rows + 1):
            if sheet.cells(i, 3).value == name and sheet.cells(i, 7).value == date:
                cell = sheet.range(i, 9)
                cell.color = (255, 255, 0)
                error_sheet.range(g_row_num_error, 1).value = name
                error_sheet.range(g_row_num_error, 2).value = sheet.range(i, 8).value
                error_sheet.range(g_row_num_error, 3).value = date
                error_sheet.range(g_row_num_error, 4).value = sheet.range(i, 9).value
                g_row_num_error += 1

    len_hour_errors = len(workhour_errors)
    if len_hour_errors >= 1:
        wb.save(new_path)
        wb.close()
        app.quit()
        print("存在日报填错情况，请修改再运行程序!\n")
        print("按下Enter键后程序将关闭!")
        input()
        return

    set2 = set(name_error)
    set1 = set(target_names)

    name_list = list(set1 - set2)
    work_hour_dict = {}
    data_tmp_list = []
    prj_tmp_list = []

    for name in name_list:
        for i in range(3, rows + 1):
            if sheet.cells(i, 3).value == name:
                prj = sheet.cells(i, 8).value
                work_hour = sheet.cells(i, 9).value
                data_tmp = sheet.cells(i, 7).value
                data_tmp_list.append(data_tmp)
                prj_tmp_list.append(prj)

                if name not in work_hour_dict:
                    work_hour_dict[name] = {}
                if prj not in work_hour_dict[name]:
                    work_hour_dict[name][prj] = 0

                work_hour_dict[name][prj] += float(work_hour)

        g_data_num = len(list(set(data_tmp_list)))
        prj_list = list(set(prj_tmp_list))
        sum_tmp = 0
        sum_tmp_sum = 0
        tmp_i = 1
        prj_length = len(prj_list)

        work_hour_list = []
        row_list = []

        for prj_name in prj_list:
            calc_sheet.range(g_row_num + 1, 1).value = name
            calc_sheet.range(g_row_num + 1, 2).value = prj_name
            work_hour_tmp = round(work_hour_dict[name][prj_name] / float(g_data_num), 2)
            calc_sheet.range(g_row_num + 1, 4).value = work_hour_tmp
            calc_sheet.range(g_row_num + 1, 5).value = work_hour_tmp

            row_list.append(g_row_num + 1)
            work_hour_list.append(work_hour_tmp)

            pattern = r"[a-zA-Z0-9]"
            prj_string = ''
            for char in prj_name:
                if re.match(pattern, char) or char == '-':
                    prj_string = prj_string + char
                else:
                    break
            calc_sheet.range(g_row_num + 1, 3).value = prj_string

            sum_tmp += work_hour_tmp
            if prj_length >= 3:
                if tmp_i == prj_length - 1:
                    sum_tmp_sum = sum_tmp

            if tmp_i == prj_length:
                if sum_tmp != 1:
                    calc_sheet.range(g_row_num + 1, 4).value = 1 - sum_tmp_sum
                    tmp_data = 1 - sum_tmp_sum

                    if abs(calc_sheet.range(g_row_num + 1, 4).value) <= 0.001:
                        calc_sheet.range(g_row_num + 1, 4).value = 0

                    if tmp_data < 0:
                        max_value = max(work_hour_list)
                        max_index = work_hour_list.index(max_value)
                        calc_sheet.range(row_list[max_index], 4).value = max_value + tmp_data - 0.01
                        calc_sheet.range(g_row_num + 1, 4).value = 0.01
                        row_list.clear()
                        work_hour_list.clear()

            g_row_num += 1
            tmp_i += 1

        data_tmp_list.clear()
        work_hour_dict.clear()
        prj_tmp_list.clear()
        g_data_num = 0

    wb.save(new_path)
    wb.close()
    app.quit()
    print("工时统计完成\n")
    print("按下Enter键后程序将关闭。")
    input()


if __name__ == "__main__":
    excel()
    sys.exit()
