#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import PDFconverter
import sys
from os import listdir
import os

# script_path表示当前文件的绝对路径
script_path = os.path.realpath(__file__)
# script_pdir表示当前文件上层文件夹的绝对路径
script_dir = os.path.dirname(script_path)
# 列出相同目录下的pdf文件
filelist = listdir(script_dir)
filelist1=[]

for i in range(0,len(filelist)):
    index = filelist[i].split(".")[-1]
    if index=='pdf':
        filelist1.append(filelist[i])

# 默认dpi为240
dpi = 240

if __name__ == '__main__':
    args = sys.argv
    # 取默认dpi
    if len(args)==1:
        print('当前dpi设置：{}'.format(dpi))
        for pdfname in filelist1:
            PDFconverter.convert(pdfname, dpi)
            if pdfname == filelist1[-1]:
                print("Finish！")
    # 自定义dpi
    elif len(args)==2:
        dpi = args[1]
        print('当前dpi设置：{}'.format(dpi))
        for pdfname in filelist1:
            PDFconverter.convert(pdfname, dpi)
            if pdfname == filelist1[-1]:
                print("Finish！")
    else:
        print('Too many arguments!ERROR!')