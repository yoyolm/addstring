# -*- coding: utf-8 -*-

from openpyxl import load_workbook

路径 = r"folder\数据.xlsx"

工作表 = load_workbook(路径).active

选择区域 = "A2:B43"
表格区域 = 工作表[选择区域]
数据列表 = []
for 单元格元组 in 表格区域:
    for 单元格对象 in 单元格元组:
        数据列表.append(单元格对象.value)
        # print(单元格对象.value, end=" ")
组合 = list(zip(数据列表[::2], 数据列表[1::2]))

字符化组合 = []
for k in 组合:
    # 循环转换为字符类型以及循环添加转换后的元素到新列表
    字符化组合.append("-".join('%s' % k for k in k))

输出 = ",".join(字符化组合)
# 输出到文本文档
with open(r"folder\数据.txt", "a+", encoding="utf-8") as k:
    k.write(输出 + "\n")
    k.close()
