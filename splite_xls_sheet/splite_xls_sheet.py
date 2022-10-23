from openpyxl import load_workbook, Workbook
from xlrd import open_workbook
from xlutils import copy
import os
import shutil
import os
import zipfile

SAVE_SHEET_NAME = ["买单",'packing_list']

NEW_DIR = "国内报关文件"

def is_sheet_need_saved(sheetname):
    for s in SAVE_SHEET_NAME:
        if s in sheetname:
            return True
    return False


def rename_donefile(file):
    if 'done_' not in file:
        os.rename(file, 'done_'+file)

def splite_xlsx(file):
    wb = 0
    try:
        wb = load_workbook(file)
    except Exception as e:
        print("openpyxl load fail "+ file)
        print(e)
        return -1
    sheetnames = wb.sheetnames
    print(sheetnames)

    for name in sheetnames:
        if not is_sheet_need_saved(name):
            ws = wb[name]
            wb.remove(ws)
    newfilepath = NEW_DIR + '/' + file
    try:
      wb.save(newfilepath)
    except Exception as e:
      print("splite_xlsx save fail "+ file)
      print(e)
      return
    rename_donefile(file)

def splite_xls(file):
    wb = open_workbook(file,formatting_info=True)
    sheetnames = wb.sheet_names()
    print(sheetnames)
    for sheet in wb.sheets():
        if not is_sheet_need_saved(sheet.name):
            continue

        wb = copy.copy(wb)
        print(wb._Workbook__worksheets)
        wb._Workbook__worksheets = [ worksheet for worksheet in wb._Workbook__worksheets if worksheet.name == sheet.name ]
        print(wb._Workbook__worksheets)
        newfilepath = NEW_DIR + '/' + file
        try:
            wb.save(newfilepath)
        except Exception as e:
            print("splite_xls save fail "+ file)
            print(e)
            return
        rename_donefile(file)

# img ---------------------------------------------------------------------------
# 判断是否是文件和判断文件是否存在
def isfile_exist(file_path):
  if not os.path.isfile(file_path):
    print("It's not a file or no such file exist ! %s" % file_path)
    return False
  else:
    return True
# 修改指定目录下的文件类型名，将excel后缀名修改为.zip


def change_file_name(file_path, new_type='.zip'):
  if not isfile_exist(file_path):
    return ''
  extend = os.path.splitext(file_path)[1]  # 获取文件拓展名
  if extend != '.xlsx' and extend != '.xls':
    print("It's not a excel file! %s" % file_path)
    return False
  file_name = os.path.basename(file_path)  # 获取文件名
  new_name = str(file_name.split('.')[0]) + new_type  # 新的文件名，命名为：xxx.zip
  dir_path = os.path.dirname(file_path)  # 获取文件所在目录
  new_path = os.path.join(dir_path, 'tmp')  # 新的文件路径
  new_path = os.path.join(new_path, new_name)  # 新的文件路径
  if os.path.exists('tmp'):
    shutil.rmtree('tmp')
  os.mkdir('tmp')
#   os.rename(file_path, new_path)  # 保存新文件，旧文件会替换掉
  shutil.copy(file_path, new_path)
  return new_path  # 返回新的文件路径，压缩包
# 解压文件


def unzip_file(zipfile_path):
  if not isfile_exist(zipfile_path):
    return False
  if os.path.splitext(zipfile_path)[1] != '.zip':
    print("It's not a zip file! %s" % zipfile_path)
    return False
  try:
    file_zip = zipfile.ZipFile(zipfile_path, 'r')
  except Exception as e:
    print("ZipFile Exception %s" % zipfile_path)
    print(e)
    return False
     
  file_name = os.path.basename(zipfile_path)  # 获取文件名
  zipdir = os.path.join(os.path.dirname(zipfile_path),
                        str(file_name.split('.')[0]))  # 获取文件所在目录
  file_zip.extractall(zipdir)
#   for files in file_zip.namelist():
#     file_zip.extract(files, zipfile_path) # 解压到指定文件目录
  # file_zip.extract(files, os.path.join(zipfile_path, files)) # 解压到指定文件目录
  file_zip.close()
  return True
# 读取解压后的文件夹，打印图片路径


def read_img(zipfile_path):
  if not isfile_exist(zipfile_path):
    return False
  dir_path = os.path.dirname(zipfile_path)  # 获取文件所在目录
  file_name = os.path.basename(zipfile_path)  # 获取文件名
  pic_dir = 'xl' + os.sep + 'media'  # excel变成压缩包后，再解压，图片在media目录
  pic_path = os.path.join(dir_path, str(file_name.split('.')[0]), pic_dir)
  if not isfile_exist(pic_path):
    return False
  file_list = os.listdir(pic_path)
  for file in file_list:
    filepath = os.path.join(pic_path, file)
    print(filepath)
    return True
  return False


def has_img(file):
  zip_file_path = change_file_name(file)
  if zip_file_path != '':
    if unzip_file(zip_file_path):
      return read_img(zip_file_path)
  return True
## #img end ---------------------------------------------------------------------
    
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
        if '买单.xls' not in file:
            newfile = NEW_DIR+'\\'+file
            rc= shutil.copy(file, NEW_DIR)
            rename_donefile(file)
            continue

        if file.endswith('xls'):
            xlsxfile = file+'x'
            os.rename(file, xlsxfile)
            file = xlsxfile

        # if has_img(file):
        #     continue
        if -1 == splite_xlsx(file):
            splite_xls(file)

if __name__ == "__main__":
    # has_img('2210007333-买单.xls')
    # has_img('2210007485-买单.xlsx')
    main()
    # if not os.path.isfile('tmp\\2210007485-买单\\xl\\media'):
    # if not os.path.isfile('tmp'):
    #     print("It's not a file or no such file exist !")
    # else:
    #    pass