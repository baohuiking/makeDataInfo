import pandas as pd

path = 'workshop.xls'
data = pd.read_excel(path, None)  # 读取数据,设置None可以生成一个字典，字典中的key值即为sheet名字，此时不用使用DataFram，会报错
print(data.keys())  # 查看sheet的名字
for sh_name in data.keys():
    print('sheet_name的名字是：', sh_name)
    sh_data = pd.DataFrame(pd.read_excel(path, sh_name))  # 获得每一个sheet中的内容
    #print(sh_data)#尽量不要打开、，里面存储了所有表单的内容

data0 = pd.read_excel(path, sheet_name=[5], usecols=[2,3], names=None)#获取 总的

data1 = pd.read_excel(path, sheet_name=[0], usecols=[8,10,11,12], names=None) #第一个表单、第八、十列负责匹配、第十一、十二行输出
print(data1)

for i in range(len(data1)):
    if(data1[i] != ''):
        for j in range(len(data0)):


