import xlwt
import xlrd
import openpyxl

# 打开Excel
wb = xlrd.open_workbook('workshop.xls')
workbook = xlwt.Workbook()

# 写数据操作
print("hello1")
workbook =openpyxl.load_workbook('workshop.xlsx')

table = workbook['Sheet1']


# 获取表
sheet1 = wb.sheet_by_index(5)
bumen1 = wb.sheet_by_index(0)
bumen2 = wb.sheet_by_index(1)
bumen3 = wb.sheet_by_index(2)
bumen4 = wb.sheet_by_index(3)
ypb = wb.sheet_by_index(4)

# 部门一数据
bumen1_k = bumen1.col_values(10)  # 部门一第k列的代号
bumen1_L = bumen1.col_values(11)  # 部门一第l列的装置名
bumen1_M = bumen1.col_values(12)  # 部门一第m列的编号
# 部门二数据
bumen2_k = bumen2.col_values(10)  # 部门二第k列的代号
bumen2_L = bumen2.col_values(11)  # 部门二第l列的装置名
bumen2_M = bumen2.col_values(12)  # 部门二第m列的编号
# 部门三数据
bumen3_k = bumen3.col_values(10)  # 部门三第k列的代号
bumen3_L = bumen3.col_values(11)  # 部门三第l列的装置名
bumen3_M = bumen3.col_values(12)  # 部门三第m列的编号
# 部门四数据
bumen4_k = bumen4.col_values(10)  # 部门四第k列的代号
bumen4_L = bumen4.col_values(11)  # 部门四第l列的装置名
bumen4_M = bumen4.col_values(12)  # 部门四第m列的编号
# 油品部数据
ypb_k = ypb.col_values(10)  # 油品部第k列的代号
ypb_L = ypb.col_values(11)  # 油品部第l列的装置名
ypb_M = ypb.col_values(12)  # 油品部第m列的编号


b = sheet1.col_values(1) # sheet1中的B列数据
# print(b)

for i in range(len(b)):  # i为sheet1里的每一行
    daihao = b[i].split()[-1]  # sheet1里B列的所有编号
    index = 'F%s' %(i+1)  # sheet1表中的F列
    index2 = 'G%s' %(i+1)  # sheet1表中的G列
    print(index, index2)
    for j in range(len(bumen1_k)):  # k为部门一中的k列每一行
        if bumen1_k[j] == daihao:
            table[index] = bumen1_L[j]
            table[index2] = bumen1_M[j]
            print(daihao, bumen1_L[j], bumen1_M[j])
    for k in range(len(bumen2_k)):  # k为部门一中的k列每一行
        if bumen2_k[k] == daihao:
            table[index] = bumen2_L[k]
            table[index2] = bumen2_M[k]
            print(daihao, bumen2_L[k], bumen2_M[k])
    for l in range(len(bumen3_k)):  # k为部门一中的k列每一行
        if bumen3_k[l] == daihao:
            table[index] = bumen3_L[l]
            table[index2] = bumen3_M[l]
            print(daihao, bumen3_L[l], bumen3_M[l])
    for m in range(len(bumen4_k)):  # k为部门一中的k列每一行
        if bumen4_k[m] == daihao:
            table[index] = bumen4_L[m]
            table[index2] = bumen4_M[m]
            print(daihao, bumen4_L[m], bumen4_M[m])
    for n in range(len(ypb_k)):  # k为部门一中的k列每一行
        if ypb_k[n] == daihao:
            table[index] = ypb_L[n]
            table[index2] = ypb_M[n]
            print(daihao, ypb_L[n], ypb_M[n])

# df = pd.read_excel("stationinfo.xlsx", usecols=[2], names=None)  # 读取项目名称和行业领域两列，并不要列名

workbook.save('workshop4.xlsx')

