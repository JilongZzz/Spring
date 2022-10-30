import os

DIR = "国内报关文件"


def  rename(file):
    if '-退税' not in file:
        return
    print(file)
    newfile = '退税-' + file.replace('-退税','')

    base_dir = os.path.dirname(os.path.abspath(__file__))  # 获取当前路径
    file = os.path.join(base_dir,DIR, file)  # 得到完整的filepath
    newfile = os.path.join(base_dir,DIR, newfile)  # 得到完整的filepath

    print(file)
    print(newfile)
    os.rename(file, newfile)


def main():
    print(__file__)
    filelist = list()
    try:
        filelist = os.listdir(DIR)
        print(filelist)
    except FileNotFoundError:
        print(DIR, "文件夹不存在")
        return
    for file in filelist:
        if os.path.isdir(file):
            continue
        rename(file)


if __name__ == "__main__":

    main()
