# -*- coding: utf-8 -*-
import xlrd
import xlwt
import os
from xlutils.copy import copy as xlcopy

def transpose_excel(file, sheet, start, stop, position, step, interval):
    """读取指定 excel 文件，在指定的 sheet 中，
    为指定的行，转换为指定的多个列。
    file: str excel 文件路径
    sheet: str excel 中读取的 sheet 名称
    start: array [row int, column int] 转置时读取的数据开始位置
    stop: array [row int, column int] 转置时读取的数据结束位置
    position: [row int, column int] 转置后，写入的位置
    step: int 转换后的单个列中的最大容量
    interval: int 每转换一列后，向右间隔指定的位置，进行下一列的写入
    """

    # 读取指定的 excel 文件
    excel = xlrd.open_workbook(file)

    # 切换至指定的 sheet 中
    sheet_obj = excel.sheet_by_name(sheet)

    if stop[1] <= start[1] or stop[0] != start[0]:
        raise TypeError("wrong position [{start}|{stop}]".format(
            start=start, stop=stop))

    # 数据获取
    data = sheet_obj.row_values(start[0], start[1], stop[1])

    # 文件拷贝
    excel_new = xlcopy(excel)
    sheet_new = excel_new.get_sheet(sheet)
    
    # 文件写入
    i = 0
    row, col = 0, 0
    for row_data in data:
        sheet_new.write(position[0] + row, position[1] + col, row_data)
        i = i + 1
        if i == step:
            col = col + interval + 1
            row = 0
            i = 0
            continue
        row = row + 1

    # 文件保存
    filename = os.path.basename(file)
    abs_path = os.path.dirname(file)
    excel_new.save(os.path.join(abs_path, "new_" + filename))

def conver_excel_position(ex_pos):
    """将 excel 的位置转换为二维坐标，
    input: str excel的位置，例如: A:1, C:4 等等
    """
    if type(ex_pos) is not str:
        raise TypeError("'ex_pos' should be string")
    col_, row = ex_pos.split(":")
    col = 0
    power = 1
    for i in range(len(col_)-1, -1, -1):
        ch = col_[i]
        col += (ord(ch)-ord('A')+1)*power
        power *= 26
    return int(row) - 1, col - 1

if __name__ == "__main__":
    # 转换测试
    transpose_excel(
        "./transpose.xlsx",   # 待转换的 excel 文件
        "Sheet1",             # excel 中的 sheet 名称
        conver_excel_position("B:1"),  # 待转换数据的起始位置
        conver_excel_position("K:1"),  # 待转换数据的停止位置
        conver_excel_position("B:5"),  # 转换后的数据起始位置
        3,                             # 转换后单列数据的最大容量
        2)                             # 转换后每列数据之间的间隔

