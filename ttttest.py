#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Jun 12 09:37:38 2018
利用百度api实现图片文本识别
@author: XnCSD
"""
# import pyHook
import ctypes
import pythoncom
import win32gui
from PIL import Image, ImageGrab
from win32api import GetSystemMetrics as gsm
import glob
from os import path
import os
from aip import AipOcr
import tkinter
import time
import cv2

def convertimg_old(picfile, outdir):
    '''调整图片大小，对于过大的图片进行压缩
    picfile:    图片路径
    outdir：    图片输出路径
    '''
    img = Image.open(picfile)
    width, height = img.size
    while(width*height > 4000000):  # 该数值压缩后的图片大约 两百多k
        width = width // 2
        height = height // 2
    new_img=img.resize((width, height),Image.BILINEAR)
    new_img.save(path.join(outdir,os.path.basename(picfile)))

def baiduOCR(picfile, outfile):
    """利用百度api识别文本，并保存提取的文字
    picfile:    图片文件名
    outfile:    输出文件
    """
    filename = path.basename(picfile)
    
    APP_ID = '16999492' # 刚才获取的 ID，下同
    API_KEY = 'nHwvwjvyAxRX3GR7Ww2NAkNd'
    SECRECT_KEY = '7G1cOkYFTZvBwGO25rGuachY2g5SHk49'
    client = AipOcr(APP_ID, API_KEY, SECRECT_KEY)
    
    i = open(picfile, 'rb')
    img = i.read()
    print("正在识别图片：\t" + filename)
    message = client.basicGeneral(img)   # 通用文字识别，每天 50 000 次免费
    #message = client.basicAccurate(img)   # 通用文字高精度识别，每天 800 次免费
    print("识别成功！")
    print("=================================================================================")
    i.close();
    
    with open(outfile, 'a+',encoding="utf-8") as fo:
        fo.writelines("+" * 60 + '\n')
        fo.writelines("识别图片：\t" + filename + "\n" * 2)
        fo.writelines("文本内容：\n")
        # 输出文本内容
        # for text in message.get('words_result'):
        #     fo.writelines(text.get('words') + '\n')
        for text in message.get("words_result"):
            print(text.get("words"))
            fo.writelines(text.get('words') + '\n')

        fo.writelines('\n'*2)
    print("============================================================")
    print()

if __name__ == "__main__":
    # capture_PIL()
    # global img
    # img = cv2.imread('./tttttemp.png')
    # cv2.namedWindow('image')
    # cv2.setMouseCallback('image', on_mouse)
    # cv2.imshow('image', img)
    # cv2.waitKey(0)

    outfile = 'export.txt'
    outdir = 'tmp'
    if path.exists(outfile):
        os.remove(outfile)
    if not path.exists(outdir):
        os.mkdir(outdir)
    print("压缩过大的图片...")
    # 首先对过大的图片进行压缩，以提高识别速度，将压缩的图片保存与临时文件夹中
    for picfile in glob.glob("picture/*"):
        convertimg_old(picfile, outdir)

    print("图片识别...")
    for picfile in glob.glob("tmp/*"):
        baiduOCR(picfile, outfile)
        os.remove(picfile)
    print('图片文本提取结束！文本输出结果位于 %s 文件中。' % outfile)
    os.removedirs(outdir)
