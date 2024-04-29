#project by: codevuive
# coder: TuanLe

import os
import pdfplumber
from docx import Document
import win32com.client

#Lay danh sach file
def get_list_file(directory):
    file_names = []
    for root, dirs, files in os.walk(directory):
        for file_name in files:
            file_names.append(os.path.join(root, file_name))
    return file_names

#dao nguoc chuoi
def reverse_text(text):
    return text[::-1]

#cat tu dau den dau .
def split_extension(text):
    return text.split(".", 1)

# lay extension file
def get_extension(name):
    tmp = reverse_text(name)
    return reverse_text(split_extension(tmp)[0])

#tu extension xac dinh kieu file
def readfile_pdf(file_path):
    with pdfplumber.open(file_path) as pdf:
        text = ''
        for page in pdf.pages:
            text += page.extract_text()
    return text

def readfile_docx(file_path):
    doc = Document(file_path)
    text = ''
    for para in doc.paragraphs:
        text += para.text + '\n'
    
    return text

# pywin32 chi ho tro tren windows
def readfile_doc(file_path):
    word = win32com.client.Dispatch("Word.Application")
    doc = word.Documents.Open(file_path)
    text = ""
    for para in doc.Paragraphs:
        text += para.Range.Text
    doc.Close()
    word.Quit()
    return text

def readfile_txt(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        content = file.read()
    return content

#doc file tu dien
def readfile_dic():
    with open('dic.txt', 'r', encoding='utf-8') as file:
        dic = file.read()
    return dic.split('\n')


if __name__ == "__main__":
    fileloc = open('fileloc.txt', 'w', encoding='utf-8')
    dics = readfile_dic()
    directory = input("Directory folder: ")
    list_file = get_list_file(directory)
    for i in list_file:
        if "~$" not in i:
            if get_extension(i) == 'txt':
                txt = readfile_txt(i)
                for dic in dics:
                    if dic in txt:
                        fileloc.write(i)
                        fileloc.write('\n')
                txt = ''

            if get_extension(i) == 'pdf':
                pdf = readfile_pdf(i)
                for dic in dics:
                    if dic in pdf:
                        fileloc.write(i)
                        fileloc.write('\n')
                pdf = ''

            if get_extension(i) == 'docx':
                docx = readfile_docx(i)
                for dic in dics:
                    if dic in docx:
                        fileloc.write(i)
                        fileloc.write('\n')
                docx = ''

            if get_extension(i) == 'doc':
                doc = readfile_doc(i)
                for dic in dics:
                    if dic in doc:
                        fileloc.write(i)
                        fileloc.write('\n')
                doc = ''
    # print(list_file)
    # print(readfile_pdf(list_file[7]))
    # print(readfile_docx(list_file[2]))

    fileloc.close()
