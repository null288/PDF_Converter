#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
from os import listdir
import os
import time
from PyPDF2 import PdfFileReader, PdfFileWriter, PdfFileMerger
from os import remove, listdir, mkdir
from os.path import join, isdir, split, splitext, basename, realpath, dirname
from reportlab.lib.pagesizes import A4, landscape, portrait
from reportlab.pdfgen import canvas
from win32com.client import constants, gencache
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
    script_dir = os.path.realpath(os.path.dirname(sys.argv[0]))
    pdf_path = script_dir + '/' + pdfFn
    dstDir = str(script_dir) + '/page_img/'
    imgDir= dstDir + '/' + pdfFn[:-4] + '_dpi' + str(dpi) +'_img'
    conDir = str(script_dir) + '/converted'
    # 转图片
    pdfconvertpicture(pdf_path, imgDir, dstDir, pdfFn, dpi)
    # 图片合并成pdf
    merge_jpg2pdf(imgDir, conDir, imgDir, pdfFn, dpi)

script_dir = os.path.realpath(os.path.dirname(sys.argv[0]))
# 列出相同目录下的pdf文件
filelist = listdir(script_dir)
filelist1=[]

for i in range(0,len(filelist)):
    index = filelist[i].split(".")[-1]
    if index=='pdf':
        filelist1.append(filelist[i])

if __name__ == '__main__':

    # 使用方法
    print('纯图像PDF转换器',
        '\n使用方法：将convert.exe文件与需要转换的所有PDF放入同一文件夹中')

    # 取默认ddpi
    ddpi = 240
    print('默认dpi：{}'.format(ddpi))
    # 自定义dpi
    dpi = input('在闭区间[1,960]中选取一个合适的整数作为dpi输入并按Enter键确认：')

    # 判断dpi是否符合要求，不符合要求的设为默认
    if dpi.isdigit():   # 判断是否是纯数字
        dpi = int(dpi)
        if dpi > 960 or dpi <= 0:
            print('输入值不符合要求，以默认dpi运行')
            dpi = ddpi
    else:
        print('输入值不符合要求，以默认dpi运行')
        dpi = ddpi

    print('当前dpi设置：{}'.format(dpi))


    # 开始转换并计时
    print('开始转换！')
    start = time.perf_counter()

    for pdfname in filelist1:
        convert(pdfname, dpi)
        if pdfname == filelist1[-1]:
            print("{}已完成全部文件转换！".format("\n"))

    # 计时结束
    end = time.perf_counter()
    print("用时：", round(end-start), 'seconds')
    input("按Enter键退出。。。")


