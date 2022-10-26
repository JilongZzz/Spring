from openpyxl import load_workbook, Workbook
from xlrd import open_workbook
import os
import shutil
import os
import win32com.client

SAVE_SHEET_NAME = ["买单", 'packing_list']

NEW_DIR = "国内报关文件"


def is_sheet_need_saved(sheetname):
    for s in SAVE_SHEET_NAME:
        if s in sheetname:
            return True
    return False


def rename_donefile(file):
    if 'done_' not in file:
        os.rename(file, 'done_'+file)


def delete_sheet(file):
    print(file)
    base_dir = os.path.dirname(os.path.abspath(__file__))  # 获取当前路径
    xlApp = win32com.client.Dispatch('Excel.Application')
    xlApp.Visible = False  # 显示excel界面
    xlApp.DisplayAlerts = False  # 是否关闭保存弹出框
    fullPath = os.path.join(base_dir, file)  # 得到完整的filepath
    wb = xlApp.Workbooks.Open(fullPath, ReadOnly=False)  # 打开对饮的excel文件
    sht_names = [sht.Name for sht in xlApp.Worksheets]
    print(sht_names)
    del_sheets = list()
    has_need_sheet = False
    for name in sht_names:
        if not is_sheet_need_saved(name):
            del_sheets.append(name)
        else:
            has_need_sheet = True
    print(del_sheets)
    have_Op = False
    if len(sht_names) > len(del_sheets):
        for name in del_sheets:
            #删除当前工作簿中相应的sheet页
            wb.Worksheets(name).Delete()
            have_Op = True
            #保存工作簿
    if have_Op or has_need_sheet:
        newpath = os.path.join(base_dir, NEW_DIR, file)
        wb.SaveAs(newpath, FileFormat=51)
        rename_donefile(file)
    xlApp.Application.Quit()


def main():
    print(__file__)
    if not os.path.exists(NEW_DIR):
        os.mkdir(NEW_DIR)

    filelist = os.listdir('.')
    print(filelist)
    for file in filelist:
        if 'done_' in file:
            continue
        if os.path.isdir(file):
            continue
        if file.endswith('.py') or file.endswith('.bat') or file.endswith('.csv'):
            continue
        if not file.endswith('买单.xls') and not file.endswith('买单.xlsx'):
            shutil.copy(file, NEW_DIR)
            rename_donefile(file)
            continue

        delete_sheet(file)


if __name__ == "__main__":

    main()
