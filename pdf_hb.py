# -*- coding: utf-8 -*-
"""
@author: hz
"""

import os
import sys
from PyPDF2 import PdfFileWriter,PdfFileReader
import time
import docx
time1 = time.time()

def getFileName(filePath, filetype):
    file_list = []
    for root, dirs, files in os.walk(filePath):
        for filePath in files:
            if filePath.endswith(filetype):
                #print(filePath)
                file_list.append(os.path.join(root, filePath))

    return file_list

def PdfToWord(filepath):
    """
    将文件夹里面的pdf文件转换成word 文件
    :param filepath: 文件夹名称
    :return:
    """
    pdf_list_names = getFileName(filepath,'.pdf')
    for pdf_name in pdf_list_names:
        pdfFileobj = open(pdf_name, 'rb')
        out_file = pdf_name.replace(".pdf", ".docx")
        pdfReader = PdfFileReader(pdfFileobj)
        pageObj = pdfReader.getPage(0)
        doc = docx.Document()
        doc.add_paragraph(pageObj.extractText())
        doc.save(out_file)
        pdfFileobj.close()

def MergePDF(filepath ,outfile):
    """
    将文件夹里面的pdf文件合并成一个文件
    :param filepath:
    :param outfile:
    :return:
    """
    output = PdfFileWriter()
    outputPages = 0
    pdf_fileName = getFileName(filepath, '.pdf')
    for each in pdf_fileName:
        input = PdfFileReader(open(each, 'rb'))

        if input.isEncrypted == True:
            input.decrypt('map')

        pageCount = input.getNumPages()
        outputPages += pageCount
        for iPage in range(0, pageCount):
            output.addPage(input.getPage(iPage))

    outputStream = open(outfile, 'wb')
    output.write(outputStream)
    outputStream.close()
    print('save:'+outfile +' finished!')

if __name__ == '__main__':
    file_dir = 'D:\\workdata\\20200530\\'
    out = r'D:\\workdata\\20200830.pdf'
    MergePDF(file_dir, out)
    #PdfToWord(file_dir)
    time2 = time.time()
    print('耗时：' + str(time2 - time1)+'s')