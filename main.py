from time import sleep
from docxtpl import DocxTemplate
from openpyxl import load_workbook
import os

# 配置路径和文件名
BASE_DIR = "/Users/rem/_Work/output/处理3"
DATA_FILE = "数据.xlsx"
TEMPLATE_FILE = "模板.docx"
OUTPUT_FILE_SUFFIX = "输出模板.docx"


# 根据路径创建文件夹
def create_path(path):
    if not os.path.exists(path):
        os.makedirs(path)


# 创建类别文件夹
def create_need_dirs():
    categories = ["电话和身份证号不全", "电话不全", "身份证号不全", "身份证号错误", "完成"]
    for category in categories:
        create_path(os.path.join(BASE_DIR, category))


# 校验row内容
def check_row_info(info_row):
    if info_row["telephone"] == 'None' and info_row["personID"] == 'None':
        return '电话和身份证号不全'
    elif info_row["telephone"] == 'None':
        return '电话不全'
    elif info_row["personID"] == 'None':
        return '身份证号不全'
    elif len(info_row["personID"]) != 18:
        return '身份证号错误'
    else:
        return '完成'


# 读取Excel数据
def read_excel_data(file_path):
    wb = load_workbook(file_path)
    ws = wb.active  # 默认获取第一个sheet
    contexts = []
    for row in range(2, ws.max_row + 1):
        context = {
            "name": str(ws[f"A{row}"].value).strip(),
            "telephone": str(ws[f"B{row}"].value).strip(),
            "personID": str(ws[f"C{row}"].value).strip(),
            "startDate": str(ws[f"D{row}"].value).strip(),
            "endDate": str(ws[f"D{row}"].value).strip(),
            "totalAmount": str(ws[f"E{row}"].value).strip()
        }
        contexts.append(context)
    return contexts


# 生成Word文档
def generate_word_documents(contexts, template_path):
    for context in contexts:
        tpl = DocxTemplate(template_path)
        tpl.render(context)
        output_path = os.path.join(BASE_DIR, check_row_info(context))
        person_file_name = context["name"].replace("/", "_")
        if context["name"] and context["name"] != 'None':
            output_file = os.path.join(output_path, f"{person_file_name}{OUTPUT_FILE_SUFFIX}")
            tpl.save(output_file)
            print(f"生成文件: {output_file}")
            sleep(1)  # 间隔1秒,为了方便按照创建时间排序


def main():
    create_need_dirs()
    data_file_path = os.path.join(BASE_DIR, DATA_FILE)
    template_file_path = os.path.join(BASE_DIR, TEMPLATE_FILE)
    contexts = read_excel_data(data_file_path)
    generate_word_documents(contexts, template_file_path)
    print("处理完成")


if __name__ == "__main__":
    main()
