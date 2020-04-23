
import xlrd
import xlwt

file = 'workshop.xls'
data = xlrd.open_workbook(file)
table = data.sheets()[0]
nrows = table.nrows
ncols = table.ncols

print(nrows, ncols)

workbook = xlwt.Workbook(encoding='utf-8')
news_sheet = workbook.add_sheet('news')
news_sheet.write(0, 0, 'Title')
news_sheet.write(0, 1, 'Content')
data = input('Input the Date,format(2019-03-19:)\n')

rank_list = []
for i in range(1, nrows):
	if table.row_values(i)[-1] == data:
		print(table.row_values(i)[-1])
		print(i)
		rank_list.append(i)
print(rank_list)

for i in range(len(rank_list)):
	news_sheet.write(i+1, 0, table.row_values(int(rank_list[i]))[0])
	news_sheet.write(i+1, 1, table.row_values(int(rank_list[i]))[1])

workbook.save('%s-网易新闻.xls' %(data))





def StringContian(strA, strB):
    for i in strB:              # 循环字符串A
        m = 0                   # 每次循环设置不等个数为0
        for j in strA:          # 循环字符串B
            if i != j:          # 如果字符串A中的第j个字符串和B中的第j个字符串不相等，就加1，不相等就加0
                m += 1
        if m >= len(strA):      # 如果不相等个数大于等于A字符串长度，就说明B的第i个字符在A字符串中没有，则绝对不包含
            return False
    return True                 # 表明B中的每个字符在A字符串中都出现了，返回Ture

strA = input().strip()          # 获取输入字符串A
strB = input().strip()          # 获取输入字符串B
print(StringContian(strA,strB)) # 打印结果