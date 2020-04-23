import xlrd
import difflib
import xlwt
import pandas as pd


def string_similar(s1, s2):
    return difflib.SequenceMatcher(None, s1, s2).quick_ratio()


# # 打开Excel文件
# wb = xlrd.open_workbook('装机.xlsx')
# wb2 = xlrd.open_workbook('stationinfo.xlsx')
#
# # 定义表格
# workbook = xlwt.Workbook()
#
# # 获取第一个表
# sheet1 = wb.sheet_by_index(0)
# sheet2 = wb2.sheet_by_index(0)
# Col = sheet1.nrows  # 表的行数
# Col2 = sheet2.nrows

df = pd.read_excel("stationinfo.xlsx", usecols=[2], names=None)  # 读取项目名称和行业领域两列，并不要列名
df_li = df.values.tolist()
print(df_li)

df2 = pd.read_excel("装机.xlsx", usecols=[1], names=None)  # 读取项目名称和行业领域两列，并不要列名
df_li2 = df2.values.tolist()
print(df_li2)

df3 = pd.read_excel("装机.xlsx", usecols=[2], names=None)  # 读取项目名称和行业领域两列，并不要列名
df_li3 = df3.values.tolist()
print(df_li3)

for i in range(180):
    k=0 #记录相似度、初始0
    for j in range(350):
        max=difflib.SequenceMatcher(None, df_li[i][0], df_li2[j][0]).quick_ratio() #相似度
        if(max>k):
            k=max
        print(max)
        print(i,j)#第一张列表的第i行，第二张列表的第j行

