

import time
import os
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font
from openpyxl.styles import Alignment
from openpyxl.styles import Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd
# pip install -i https://pypi.tuna.tsinghua.edu.cn/simple pandas
# pip install -i https://pypi.tuna.tsinghua.edu.cn/simple openpyxl
# pip install -i https://pypi.tuna.tsinghua.edu.cn/simple xlwt
path = '.'


def scan_file(path):
    filelist = list()
    all = os.listdir(path)
    for f in all:
        if 'MATS' in f and 'csv' in f:
            filelist.append(f)

    return filelist


def main():
    # recordNo = 'MATS7624726000'
    # infile = 'MATS3757157000.csv'
    # outfile = recordNo + '-进港舱单.xls'
    filelist = scan_file(path)
    print(filelist)

    # op_one_file(infile, outfile)
    for infile in filelist:
        recordNo = infile.rstrip('.csv')
        outfile = recordNo + '-进港舱单.xlsx'
        op_file(infile, outfile, recordNo)
    time.sleep(3)
    print('ok')


def op_file(infile, outfile, recordNo):
    df = pd.read_csv(infile, sep=",", encoding="gbk")
    # df = df[['工作号/入仓单号', '客户名','客服','品名','件数', '毛重', '体积', '报关', '排柜备注']]
    df = df[['工作号/入仓单号', '件数', '毛重', '体积', '报关', '排柜备注']]
    # print(df)
    dfrank = df.sort_values(by='报关', ascending=False)
    dfrank.insert(1, 'HS code', '')
    dfrank.insert(0, '关单号', '')
    dfrank.insert(0, '顺序', '')
    df = dfrank
    wb = Workbook()
    ws = wb.active

    ws.append(["提单号", recordNo])
    ws.append(["船名航次", ''])
    ws.append(["柜号", ''])
    ws.append(["封条号", ''])
    ws.append(["卸货港", ''])
    ws.append(["目的港", ''])

    head_row_count = 6
    content_row_count = 0

    for row in dataframe_to_rows(df, index=False, header=True):
        ws.append(row)
        content_row_count += 1

    headfont = Font(name="微软雅黑", size=10, bold=False,
                    italic=False, color="FF0000")
    titlefont = Font(name="微软雅黑", size=10   , bold=True,
                     italic=False, color="000000")
    contentfont = Font(name="微软雅黑", size=10, bold=False,
                       italic=False, color="000000")
    alignment = Alignment(horizontal="center", vertical="center")
    border = Border(left=Side(border_style='thin', color='000000'),
                    right=Side(border_style='thin', color='000000'),
                    top=Side(border_style='thin', color='000000'),
                    bottom=Side(border_style='thin', color='000000'))

    n_row = 0
    for row in ws.iter_rows():
        n_row += 1
        if n_row <= head_row_count:
            print(row)
            col = 0
            for cell in row:
                col+=1
                print(cell, cell.value, "headfont")
                cell.font = headfont
                cell.alignment = alignment
                cell.border = border
                if col>=2:
                    break
            continue
        for cell in row:
            print(cell, cell.value, "contentfont")
            cell.font = contentfont
            cell.alignment = alignment
            cell.border = border

    for i in range(1, df.shape[1]+1):
        cell = ws.cell(row=head_row_count + 1, column=i)
        print(cell.value)
        cell.font = titlefont

    ws.column_dimensions["I"].width = 50
    # 设置连续行行高：
    for r in range(head_row_count+1, head_row_count+content_row_count+1):  # 注意，行和列的序数，都是从1开始
        ws.row_dimensions[r].height = 22  #
    # 设置连续列列宽：
    for c in range(1, 8):  # 注意，列序数从1开始，但必须转变为A\B等字母
        w = get_column_letter(c)  # 把列序数转变为字母
        ws.column_dimensions[w].width = 16

    wb.save(outfile)


if __name__ == '__main__':
    main()
    # op_file('MATS4366005000.csv', 'out.xlsx', 'MATS4366005000')
