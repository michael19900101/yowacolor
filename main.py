import tkinter.messagebox
import xml.dom.minidom
import os
import pandas as pd

# 输入输出文件的路径
inputFilePath = '颜色规范.xlsx'
outPutFileDir = "yowa_color_output/"

def mkdir(path):
    folder = os.path.exists(path)
    if not folder:  # 判断是否存在文件夹如果不存在则创建为文件夹
        os.makedirs(path)  # makedirs 创建文件时如果路径不存在会创建这个路径
        print
        "---  new folder...  ---"
        print
        "---  OK  ---"
    else:
        print
        "---  There is this folder!  ---"


# 在内存中创建一个空的文档
docDay = xml.dom.minidom.Document()
docNight = xml.dom.minidom.Document()
# 创建一个根节点resources对象
rootDay = docDay.createElement('resources')
rootNight = docNight.createElement('resources')
# 将根节点添加到文档对象中
docDay.appendChild(rootDay)
docNight.appendChild(rootNight)

# 读取excel中所有数据
try:
    f = pd.ExcelFile(inputFilePath)
    for sheet_name in f.sheet_names:
        print("----------【" + sheet_name + "】页遍历开始---------------")

        rootDay.appendChild(docDay.createComment("【" + sheet_name + "】页遍历开始"))
        rootNight.appendChild(docNight.createComment("【" + sheet_name + "】页遍历开始"))

        ori = pd.read_excel(inputFilePath, sheet_name=sheet_name, usecols=['KEY', '正常模式', '黑夜模式'])

        # 选取数据中需要的部分，先是行，后是列
        data = ori.iloc[0:, 0:5]

        if len(data.columns) == 0 :
            print("----------【" + sheet_name + "】页遍历结束---------------")
            rootDay.appendChild(docDay.createComment("【" + sheet_name + "】页遍历结束"))
            rootNight.appendChild(docNight.createComment("【" + sheet_name + "】页遍历结束"))
            continue

        # 给选取的数据列起个名字，方便后面使用
        data.columns = ["key", "day_value", "night_value"]

        # 遍历每一行
        for index, row in data.iterrows():
            if pd.isnull(row['day_value']):
                print("【" + str(row['key']) + "】[正常模式]的值为空")
            else:
                nodeDayColor = docDay.createElement('color')
                nodeDayColor.setAttribute('name', row['key'])
                nodeDayColor.appendChild(docDay.createTextNode(row['day_value'][0:9]))
                rootDay.appendChild(nodeDayColor)

            if pd.isnull(row['night_value']):
                print("【" + str(row['key']) + "】[黑夜模式]的值为空")
            else:
                nodeNightColor = docNight.createElement('color')
                nodeNightColor.setAttribute('name', row['key'])
                nodeNightColor.appendChild(docDay.createTextNode(row['night_value'][0:9]))
                rootNight.appendChild(nodeNightColor)

        print("----------【" + sheet_name + "】页遍历结束---------------")
        rootDay.appendChild(docDay.createComment("【" + sheet_name + "】页遍历结束"))
        rootNight.appendChild(docNight.createComment("【" + sheet_name + "】页遍历结束"))

    # 写入文件
    mkdir(outPutFileDir)
    fpDay = open(outPutFileDir + 'values.xml', 'w', encoding='utf-8')
    fpNight = open(outPutFileDir + 'values-night.xml', 'w', encoding='utf-8')
    docDay.writexml(fpDay, indent='', addindent='\t', newl='\n', encoding='utf-8')
    docNight.writexml(fpNight, indent='', addindent='\t', newl='\n', encoding='utf-8')
except FileNotFoundError as e:
    print("Excel file not found " + str(e))
    tkinter.messagebox.showinfo('提示', "找不到\"" + inputFilePath + "\"")
