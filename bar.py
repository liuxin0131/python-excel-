import xlrd
import matplotlib.pyplot as plt
# 打开excel表格
data_excel = xlrd.open_workbook('2.xlsx')
table = data_excel.sheets()[0]
#划定区域
b_row1=9  #开始行
e_row1=14  #结束行
b_col1=11   #开始列
e_col1=17  #结束列
#读取标签
list_values1=[]
for i in range(b_row1, e_row1):
    col = table.col_values(i)[1]
    list_values1.append(col)
print(list_values1)
#提取单元格内容
for x in range(b_col1,e_col1):
    values = []
    col = table.row_values(x)
    col1=table.row_values(x)[1]
    # print(col1)
    for i in range(b_row1, e_row1):
        values.append(col[i])
    # print(values)
    plt.rcParams['font.sans-serif'] = ['SimHei']  # 正常显示中文标签
    labels = values
    sizes = values
    # explode = (0, 0.1, 0, 0)
    fig1, ax1 = plt.subplots()
    ax1.bar(list_values1, values)
    for x, y in enumerate(values):
        plt.text(x, y + 4, y, ha='center', va='bottom')
    plt.title(col1+"情况")
    plt.savefig("C:/Users/LLLLL/Desktop/作图任务/{}情况.png".format(col1))
    plt.show()





