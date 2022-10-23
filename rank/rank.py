

from statistics import mode
import time
import os
import openpyxl
import pandas as pd
# pip install -i https://pypi.tuna.tsinghua.edu.cn/simple pandas
# pip install -i https://pypi.tuna.tsinghua.edu.cn/simple openpyxl
# pip install -i https://pypi.tuna.tsinghua.edu.cn/simple xlwt
path='.'

def get_head(recordNo):

    MAIN = ['提单号','船名航次','柜号','封条号','卸货港','目的港']
    cont = [recordNo, '','','','','']
   
    # 字典
    dict = {'MAIN': MAIN, '': cont}
        
    df = pd.DataFrame(dict)
    # print(df)
    return df


def op_one_file(infile, outfile,recordNo):
    df_head = get_head(recordNo)

    df = pd.read_csv(infile, sep=",", encoding="gbk")
    # 顺序	关单号	工作号/入仓单号	客户名	客服	品名	xx	HS code		件数	毛重	体积	报关	排柜备注

    # df = df[['工作号/入仓单号', '客户名','客服','品名','件数', '毛重', '体积', '报关', '排柜备注']]
    df = df[['工作号/入仓单号', '件数', '毛重', '体积', '报关', '排柜备注']]
    # print(df)
    dfrank = df.sort_values(by='报关', ascending=False)
    # dfrank.insert(3, '', '')
    dfrank.insert(1, 'HS code', '')
    # dfrank.insert(3, 'xx', '')
    # dfrank.insert(1, '品名', '')
    # dfrank.insert(1, '客服', '')
    # dfrank.insert(1, '客户名', '')
    dfrank.insert(0, '关单号', '')
    dfrank.insert(0, '顺序', '')


    with pd.ExcelWriter(outfile) as writer:
        # df_head.to_excel(writer, index=False, header=False)
        df_head.to_excel(writer, index=False)
        dfrank.to_excel(writer, index=False, startrow=7)



def scan_file(path):
    filelist=list()
    all=os.listdir(path)
    for f in all:
        if 'MATS' in f and 'csv' in f :
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
        op_one_file(infile,outfile,recordNo)
    time.sleep(3)
    print('ok')





if __name__ == '__main__':
    main()
# df1 = pd.DataFrame([["AAA", "BBB"]], columns=["Spam", "Egg"])  

# df2 = pd.DataFrame([["ABC", "XYZ"]], columns=["Foo", "Bar"])  

# with pd.ExcelWriter("path_to_file.xlsx") as writer:

#     df1.to_excel(writer, sheet_name="Sheet1")  

#     df2.to_excel(writer, startrow=5)  