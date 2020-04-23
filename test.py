import difflib
import pandas as pd

def string_similar(s1, s2):
    return difflib.SequenceMatcher(None, s1, s2).quick_ratio()

fb =open('匹配查询结果.txt', 'w', encoding='utf-8')

df = pd.read_excel("workshop.xlsx",sheet_name=[0], usecols=[11], names=None)  # 读取项目名称和行业领域两列，并不要列名
df_li = df.values.tolist()
print(df_li)
print(df)

df2 = pd.read_excel("workshop.xlsx",sheet_name=[6],usecols=[2,3], names=None)  # 读取项目名称和行业领域两列，并不要列名
df_li2 = df2.values.tolist()
print(df_li2)
print(df2.size)

for i in range(len(df_li)):
    lishi_max = 0  # 记录最大相似，初值为0
    k = 0  # 记录最大相似的行
    for j in range(len(df_li2)):
        max = string_similar(str(df_li[i][0]), str(df_li2[j][0]))
        if max>lishi_max:
            lishi_max = max
            k = j
        # print(df_li[i][0], '与', df_li2[j][0], '的相似度为', lishi_max, max)
        # print(df_li[i][0], '与', df_li2[j][0], '的相似度最为', max)
    #print( '第', i+2, '行的', df_li[i][0], '与', '第', k+2,'行的', df_li2[k][0], '装机量：', df_li2[k][1], '的相似度最高,相似度为', lishi_max)
    print(df_li[i][0],"-------",df_li2[k][0],"-------",lishi_max)
    print(df_li2[k][1],file=fb)
fb.close()
    # fb.write( df_li2[k][1])
    # print(df_li2[k][1])


# for df_index,df_value in enumerate(df_li):
#     linshi_max = 0 #记录每一层for循环后的最大相似度
#     linshi_df2_data = None
#     df2_name = ""
#     for df2_index, df2_value in enumerate(df_li2):
#         max  = string_similar(str(df_value[0]), str(df2_value[0])) #相似度
#         if max > linshi_max:
#             linshi_max = max
#             df2_name = str(df2_value[0])
#     print(df_value, "----", df2_name,"---相似度为：",linshi_max)
# print("正在执行第：", df_index, "个---最大相似度为：", linshi_max)
    #print("容量为:",df_li2[1]["与相似对应的第几个"])

 # for d1 in df_li:
 #     for d2 in df_li2:
 #         print(str(d1),d2)

# for i in range(180):
#     k=0 #记录相似度、初始0
#     for j in range(350):
#         print(df_li[i][0],df_li2[j][0])
        # max=string_similar(None, df_li[i][0], df_li2[j][0]) #相似度
        # if(max>k):
        #     k=max
        # print(max)
        # print(i,j)#第一张列表的第i行，第二张列表的第j行

