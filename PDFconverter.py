#!/usr/bin/env python3
# -*- coding: utf-8 -*-

' a converter module '

from PyPDF2 import PdfFileReader, PdfFileWriter, PdfFileMerger
from os import remove, listdir, mkdir
from os.path import join, isdir, split, splitext, basename, realpath, dirname
from reportlab.lib.pagesizes import A4, landscape, portrait
from reportlab.pdfgen import canvas
from win32com.client import constants, gencache
import os
import fitz

def pdfconvertpicture(pdf_path, imgDir, dstDir, pdfFn, dpi):
    if not isdir(dstDir):
        mkdir(dstDir)
    if not isdir(imgDir):
        mkdir(imgDir)
    #使用fitz.open()创建文档对象
    with fitz.open(pdf_path) as pdf:
        for pg in range(0, pdf.page_count):
            page = pdf[pg]
            #设置缩放和旋转系数,zoom_x, zoom_y取相同值，表示等比例缩放
            mat = fitz.Matrix(2, 2)
            #个人感觉png的效果更清晰一点。
            pm = page.get_pixmap(matrix=mat,dpi=dpi,alpha=False)
            page_num = pg + 1  # 页码从1开始
            pm.save(f'{imgDir}/{page_num}.png')  # 第1张图片名：1.png，以此类推
            print('PDF文件||{}||第{}页已转为图片'.format(pdfFn, page_num))

# 把图片合并为pdf文件
def merge_jpg2pdf(jpgpath, conDir, imgDir, pdfFn, dpi):
    # 要合并的图片
    jpg_files = [join(jpgpath, fn) for fn in listdir(jpgpath)
                    if fn.endswith('.png')]
    jpg_files.sort(key = lambda fn: int(splitext(basename(fn))[0]))
    result_pdf = PdfFileMerger()
    # 临时文件
    temp_pdf = ('{}\{}_dpi{}_temp.pdf'.format(imgDir, pdfFn[:-4], dpi))
    # 依次转pdf，再合并pdf
    for fn in jpg_files:
        # 转pdf，portrait纵向页面，landscape横向页面
        c = canvas.Canvas(temp_pdf, pagesize = portrait(A4))
        c.drawImage(fn, 0 , 0, *portrait(A4))       ## 注意：页面大小是A4
        c.save()
        # 合并
        with open(temp_pdf, 'rb') as fp:
            pdf_reader = PdfFileReader(fp)
            result_pdf.append(pdf_reader)
    # 保存结果
    if not isdir(conDir):
        mkdir(conDir)
    result_pdf.write(conDir + '\{}_dpi{}_converted.pdf'.format(pdfFn[:-4], dpi))
    print('PDF文件||{}||转换已成纯图像PDF文件：{}_dpi{}_converted.pdf'.format(pdfFn, pdfFn[:-4], dpi))
    result_pdf.close
    remove(temp_pdf)

def convert(filename, dpi):
    file = filename
    dstDir, pdfFn = split(file)
    script_path = os.path.realpath(__file__)
    script_dir = os.path.dirname(script_path)
    pdf_path = script_dir + '/' + pdfFn
    dstDir = str(script_dir) + '/page_img/'
    imgDir= dstDir + '/' + pdfFn[:-4] + '_dpi' + str(dpi) +'_img'
    conDir = str(script_dir) + '/converted'
    # 转图片
    pdfconvertpicture(pdf_path, imgDir, dstDir, pdfFn, dpi)
    # 图片合并成pdf
    merge_jpg2pdf(imgDir, conDir, imgDir, pdfFn, dpi)

if __name__ == '__main__':
    convert(filename, dpi)