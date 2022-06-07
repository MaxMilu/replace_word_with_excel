from time import sleep

from docxtpl import DocxTemplate
from openpyxl import load_workbook
import os

baseDir = "/Users/rem/_Work/output/处理3"  # 基础处理路径
dataFile = "/数据.xlsx"  # 整理好的数据文件
templateFile = "/模板.docx"  # 模板文件
outputFileSubfix = "一期多层大众新城彩生活物业起诉状.docx"  # 通用输出文件后缀


# 根据路径创建文件夹
def create_path(path):
    if not os.path.exists(path):
        os.mkdir(path)
    return


# 创建类别文件夹
def create_need_dirs():
    create_path(baseDir + "/电话和身份证号不全")
    create_path(baseDir + "/电话不全")
    create_path(baseDir + "/身份证号不全")
    create_path(baseDir + "/身份证号错误")
    create_path(baseDir + "/完成")
    return


# 校验row内容
def check_row_info(info_row):
    if info_row["telephone"] == 'None' and info_row["personID"] == 'None':
        return '电话和身份证号不全'
    elif info_row["telephone"] == 'None':
        return '电话不全'
    elif info_row["personID"] == 'None':
        return '身份证号不全'
    elif not len(info_row["personID"]) == 18:
        return '身份证号错误'
    else:
        return '完成'


wb = load_workbook(baseDir + dataFile)  # 根据路径打开excel文件
ws = wb['Sheet1']  # excel 的 sheet 页名称
contexts = []  # 新建数据数组
for row in range(2, ws.max_row + 1):  # 设定循环范围 第二行开始 到 excel最大行数
    # 值提取 并且转换为字符串去除两边空格
    name = str.strip(str(ws["A" + str(row)].value))
    telephone = str.strip(str(ws["B" + str(row)].value))
    personID = str.strip(str(ws["C" + str(row)].value))
    startDate = str.strip(str(ws["D" + str(row)].value))
    endDate = str.strip(str(ws["D" + str(row)].value))
    totalAmount = str.strip(str(ws["E" + str(row)].value))
    context = {"name": name, "telephone": telephone, "personID": personID, "startDate": startDate, "endDate": endDate,
               "totalAmount": totalAmount}
    contexts.append(context)

create_need_dirs()  # 创建输出文件夹
print("开始处理")
for i, context in enumerate(contexts):
    print("正在处理: {}".format(context["name"]))
    tpl = DocxTemplate(baseDir + templateFile)  # 打开模板文件
    tpl.render(context)  # 开始替换文字

    outputPath = check_row_info(context)  # 用于获取校验的分类

    personFileName = context["name"]
    if personFileName.find('/'):
        personFileName = personFileName
    personFileName = personFileName.replace("/", "_")

    # 如果 name 不为空 则输出
    if context["name"] != '' or context["name"] != 'None':
        tpl.save(baseDir + "/{}/{}{}".format(outputPath, personFileName, outputFileSubfix))
        sleep(1)  # 间隔1秒,为了方便按照创建时间排序
