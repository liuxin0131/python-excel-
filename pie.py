import xlrd
import matplotlib.pyplot as plt
# 打开excel表格
data_excel = xlrd.open_workbook('2.xlsx')
table = data_excel.sheets()[0]
#划定饼状图区域
b_row=11  #开始行
e_row=17  #结束行
b_col=9   #开始列
e_col=14  #结束列

#读取饼状图标签
list_values=[]
for i in range(b_row, e_row):
    row = table.row_values(i)[1]
    list_values.append(row)
# print(list_values)

#提取单元格内容
for x in range(b_col,e_col):
    values = []
    col = table.col_values(x)
    col1=table.col_values(x)[1]
    for i in range(b_row, e_row):
        values.append(col[i])
    # print(values)
    plt.rcParams['font.sans-serif'] = ['SimHei']  # 正常显示中文标签
    labels = values
    sizes = values
    # explode = (0, 0.1, 0, 0)
    fig1, ax1 = plt.subplots()
    ax1.pie(sizes, labels=labels, autopct='%1.1f%%', shadow=False, startangle=90)
    ax1.axis('equal')
    plt.legend(list_values,loc="upper left",bbox_to_anchor=(-0.168,0.95))
    plt.title(col1+"任务分配与经费占比情况")
    plt.savefig("C:/Users/LLLLL/Desktop/作图任务/{}任务分配与经费占比情况.png".format(col1))
    plt.show()





