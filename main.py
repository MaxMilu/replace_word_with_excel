from time import sleep

from docxtpl import DocxTemplate
from openpyxl import load_workbook
import os

baseDir = "C:/Users/Administrator/Desktop/新建文件夹/处理"
dataFile = "/数据.xlsx"
# dataFile = "/测试数据.xlsx"
templateFile = "/模板.docx"
outputFileSubfix = "一期多层大众新城彩生活物业起诉状.docx"


# 根据路径创建文件夹
def create_path(path):
    if not os.path.exists(path):
        os.mkdir(path)
    return


# 创建类别文件夹
def create_need_dirs():
    create_path(baseDir + "./电话和身份证号不全")
    create_path(baseDir + "./电话不全")
    create_path(baseDir + "./身份证号不全")
    create_path(baseDir + "./身份证号错误")
    create_path(baseDir + "./完成")
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


wb = load_workbook(baseDir + dataFile)
ws = wb['Sheet1']
contexts = []
for row in range(2, ws.max_row + 1):
    name = str(ws["A" + str(row)].value)
    telephone = str(ws["B" + str(row)].value)
    personID = str(ws["C" + str(row)].value)
    startDate = str(ws["D" + str(row)].value)
    totalAmount = str(ws["E" + str(row)].value)
    context = {"name": name, "telephone": telephone, "personID": personID, "startDate": startDate,
               "totalAmount": totalAmount}
    contexts.append(context)

create_need_dirs()

for i, context in enumerate(contexts):
    print("正在处理: {}".format(context["name"]))
    tpl = DocxTemplate(baseDir + templateFile)
    tpl.render(context)

    outputPath = check_row_info(context)

    personFileName = context["name"]
    if personFileName.find('/'):
        personFileName = personFileName
    personFileName = personFileName.replace("/", "_")

    # 如果 name 不为空 则输出
    if context["name"] != '' or context["name"] != 'None':
        tpl.save(baseDir + "/{}/{}{}".format(outputPath, personFileName, outputFileSubfix))
        sleep(1)
