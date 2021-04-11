import os
import sys
import io
import re
import pdfminer.high_level # $ pip install pdfminer.six
from pdfminer.layout import LTTextContainer
from pdfminer.layout import LTTextLineVertical, LTTextBoxHorizontal, LTImage
from pdfminer.high_level import extract_pages
import pandas as pd


# Заменяет любой символ неподходящий под шаблон на \n
# (шаблон: все прописные буквы, числа, пробелы)
def find_text(text):
    return re.sub(r'[^А-ЯA-Z\d\s]', '\n', text)


def get_pdf_content(name):
    output = io.StringIO()
    laparams = pdfminer.layout.LAParams(detect_vertical=True)
    with open(name, 'rb') as file:
        pd = pdfminer.high_level.extract_text_to_fp(file, output, laparams=laparams)
    value = output.getvalue()
    return value


def get_array_values(content):
    result = []
    values = find_text(content).replace(' ', '').split('\n')
    for v in values:
        if v not in result and re.fullmatch(r'\d{1}\w{2}\d{2}[\w\d]*', v):
            result.append(v)
    return result


def import_to_excel(writer, name, arr):
    df = pd.DataFrame({'value': arr})
    df.to_excel(writer, sheet_name=name, index=False)
    return 0


if __name__ == '__main__':
    path = './PDF/'
    files = os.listdir(path)
    writer = pd.ExcelWriter('./result.xlsx', engine='xlsxwriter')
    for file in files:
        content = get_pdf_content(path + file)
        values = get_array_values(content)
        import_to_excel(writer, file[:-4], values)
        print(file[:-4] + ' ready')

    writer.save()


# strings = []
#
# for page_layout in extract_pages("./PDF/test.pdf"):
#
#     for element in page_layout:
#         if isinstance(element, LTTextBoxHorizontal):
#             print(element)
#             print('----')
#             strings.append(element.get_text())
            # print(element.get_text())
            # print('----')
# def get_answer(text):
#     if re.findall(r'[?]', text) and not re.findall(r'https://', text):
#         return text.rstrip()
#     return 0

