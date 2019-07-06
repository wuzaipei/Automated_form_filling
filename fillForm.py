# coding:utf-8

from docx import Document


import re,os

filePath = os.path.abspath("./") + "\\surfaceFile\\表扬信.docx"

document = Document(filePath)  #打开文件demo.docx
print(type(document))
for paragraph in document.paragraphs:
    paragraph.text = re.sub("论语","名言名句",paragraph.text)
    print(paragraph.text)  # 打印各段落内容文本
    if paragraph.text == "":
        paragraph.text = "光阴似箭，日月如梭"

document.save(filePath) #保存文档