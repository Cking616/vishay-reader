# -*- coding: utf-8 -*-

"""
@version: 1.0
@license: Apache Licence
@author:  kht,cking616
@contact: cking616@mail.ustc.edu.cn
@software: PyCharm Community Edition
@file: __main__.py
@time: 2018/5/11 9:08
"""
import re
from collections import defaultdict
import os
import xlwt


xlt_header = ['lot id', 'count', 'receipe', 'starttime', 'endtime']


def generate_style():
    borders = xlwt.Borders()  # Create Borders
    borders.left = xlwt.Borders.THIN
    """
    # May be: NO_LINE, THIN, MEDIUM, DASHED, DOTTED, THICK, DOUBLE, HAIR, MEDIUM_DASHED, THIN_DASH_DOTTED, 
    MEDIUM_DASH_DOTTED, THIN_DASH_DOT_DOTTED, MEDIUM_DASH_DOT_DOTTED, SLANTED_MEDIUM_DASH_DOTTED, or 0x00 through 0x0D.
    """
    borders.right = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN
    borders.left_colour = 0x40
    borders.right_colour = 0x40
    borders.top_colour = 0x40
    borders.bottom_colour = 0x40

    alignment = xlwt.Alignment()  # Create Alignment
    alignment.horz = xlwt.Alignment.HORZ_CENTER
    """
    # May be: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER, HORZ_RIGHT, HORZ_FILLED, HORZ_JUSTIFIED, 
      HORZ_CENTER_ACROSS_SEL, HORZ_DISTRIBUTED
    """
    alignment.vert = xlwt.Alignment.VERT_CENTER
    # May be: VERT_TOP, VERT_CENTER, VERT_BOTTOM, VERT_JUSTIFIED, VERT_DISTRIBUTED

    font = xlwt.Font()
    font.name = 'SimSun'  # 指定“宋体”

    style = xlwt.XFStyle()  # Create Style
    style.borders = borders  # Add Borders to Style
    style.alignment = alignment  # Add Alignment to Style
    style.font = font
    return style


def analysis_original_file(filename):
    d = defaultdict()
    with open(filename, encoding='Big5') as f:
        for line in f:
            row = line.split(',')
            if row[0].startswith('LOT ID'):
                d['id'] = row[1]
            if row[0].startswith('RECEIPE'):
                d['rec'] = row[1] + row[2]
            if row[0].startswith('STARTTIME'):
                d['stime'] = row[1]
            if row[0].startswith('ENDTIME'):
                d['etime'] = row[1]
            if row[0].startswith('COUNT'):
                d['count'] = row[1]
    return d


def analysis_original_dir(target_year, target_month):
    xls_data = []
    log_dir = '.\Log'
    if target_month < 10:
        target_date = str(target_year) + '0' + str(target_month)
    else:
        target_date = str(target_year) + str(target_month)
    for (root, dirs, files) in os.walk(log_dir):
        for filename in files:
            if filename.endswith('.csv') and filename.find(target_date) != -1:
                d = analysis_original_file(os.path.join(root, filename))
                xls_data.append(d)
    return xls_data


def dat2xls(out_filename, xls_data):
    style = generate_style()

    workbook = xlwt.Workbook(encoding='utf-8')
    worksheet = workbook.add_sheet('sheet1')

    index = 0
    for data in xlt_header:
        worksheet.write(0, index, data, style)
        index = index + 1

    index = 1
    while len(xls_data) != 0:
        line = xls_data.pop(0)
        worksheet.write(index, 0, line['id'], style)
        worksheet.write(index, 1, line['count'], style)
        worksheet.write(index, 2, line['rec'], style)
        worksheet.write(index, 3, line['stime'], style)
        worksheet.write(index, 4, line['etime'], style)
        index = index + 1

    if not out_filename.endswith('.xls'):
        out_filename = out_filename + '.xls'
    workbook.save(out_filename)


def translate_process():
    print("用法：第一步，将需要转换的原始文件复制到Log文件夹下\n")
    if not os.path.exists('Log'):
        os.mkdir('Log')
    if not os.path.exists('excel'):
        os.mkdir('excel')
    input("完成后请按Enter继续\n")

    input_choice = int(input("请输入筛选的年份(范围2017到2050)\n"))
    while input_choice < 2017 or input_choice > 2050:
        input_choice = int(input("输入的数字不在范围内，请重新输入\n"))
    target_year = input_choice

    print("以下为循环选择，0退出")
    while input_choice != 0:
        input_choice = int(input("请输入筛选的月份(范围1到12),输入0退出\n"))
        if input_choice == 0:
            break
        while input_choice < 1 or input_choice > 12:
            input_choice = int(input("输入的数字不在范围内，请重新输入\n"))
        target_month = input_choice

        xls_filename = './excel/' + str(target_year) + '-' + str(target_month) + '.xls'

        txt = "在excel文件夹下生成" + xls_filename + '\n'
        print(txt)

        xls_data = analysis_original_dir(target_year, target_month)
        dat2xls(xls_filename, xls_data)


if __name__ == '__main__':
    translate_process()
    input("请按Enter键继续\n")
