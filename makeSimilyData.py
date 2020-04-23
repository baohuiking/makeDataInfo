# -*- coding:utf-8 -*-
import xlrd
import xlwt
from xlutils3.copy import copy


def set_style():
    style = xlwt.XFStyle()  # 初始化样式

    borders = xlwt.Borders()
    borders.left = 1
    borders.right = 1
    borders.top = 1
    borders.bottom = 1
    borders.bottom_colour = 0x3A
    style.borders = borders

    font = xlwt.Font()
    font.name = 'Times New Roman'
    style.font = font

    return style


if __name__ == '__main__':

    '''程序默认目的数据表在第一张'''
    '''程序直接覆盖目的数据表，务必进行备份以免出现错误不可回退'''

    rows = 66  # 源数据行数
    yuan = '装机.xlsx'  # 源
    mudi = 'satationinfo.xlsx'  # 目的

    start_row = 2  # 合并单元格的起始行数
    leap = 9  # 合并单元格的数目
    delta = 2  # 合并单元格的变化行数

    data = xlrd.open_workbook(yuan)
    source = data.sheets()[0]

    oldwb = xlrd.open_workbook(mudi, formatting_info=True)
    newwb = copy(oldwb)
    result = newwb.get_sheet(0)

    for i in range(0, rows):
        row = source.row_values(i)
        result.write_merge(start_row, start_row + leap, 0, 0, i + 1, set_style())
        for j in range(1, 5):
            result.write_merge(start_row, start_row + leap, j, j, row[j], set_style())
        start_row += leap + 2
        if start_row == 14:
            leap -= delta

    newwb.save('satationinfo.xlsx')