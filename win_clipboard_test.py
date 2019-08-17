# python读取剪切板内容

import win32con
from ctypes import Structure, sizeof, memmove, pointer, memset
from ctypes.wintypes import WORD, DWORD, LONG
import win32clipboard
import sys, time
import win32api
from tempfile import TemporaryFile
from PIL import Image, ImageGrab
import cv2
import glob
from os import path
import os
from aip import AipOcr
import tkinter
import time
import base64  
TEMP_FILE_NAME = "ay9a78wq_temp_u232dsjfb_temp.png"
OUTPUT_TXT = 'export.txt'

class BITMAPFILEHEADER(Structure):    # https://msdn.microsoft.com/en-us/library/windows/desktop/dd183374(v=vs.85).aspx
    _pack_   = 1                      # structure field byte alignment
    _fields_ = [
        ('bfType',      WORD),
        ('bfSize',      DWORD),
        ('bfReserved1', WORD),
        ('bfReserved2', WORD),
        ('bfOffBits',   DWORD),
        ]   
SIZEOF_BITMAPFILEHEADER = sizeof(BITMAPFILEHEADER)

class BITMAPINFOHEADER(Structure):          # https://msdn.microsoft.com/en-us/library/windows/desktop/dd183376(v=vs.85).aspx
    _pack_   = 1
    _fields_ = [
        ('biSize',          DWORD),
        ('biWidth',         LONG),
        ('biHeight',        LONG),
        ('biPLanes',        WORD),
        ('biBitCount',      WORD),
        ('biCompression',   DWORD),
        ('biSizeImage',     DWORD),
        ('biXpelsPerMeter', LONG),
        ('biYpelsPerMeter', LONG),
        ('biClrUsed',       DWORD),
        ('biClrImportant',  DWORD),
        ]
SIZEOF_BITMAPINFOHEADER = sizeof(BITMAPINFOHEADER)

def getClipboardImage():
    BI_BITFIELDS = 3
    win32clipboard.OpenClipboard()    # https://msdn.microsoft.com/zh-cn/library/windows/desktop/ff468802(v=vs.85).aspx
    try:
        if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_DIB):    # https://msdn.microsoft.com/en-us/library/windows/desktop/ms649013(v=vs.85).aspx
            data = win32clipboard.GetClipboardData(win32clipboard.CF_DIB)
        else:
            print('cliboard does not contain an image in DIB format.')
            sys.exit(1)
    finally:
        win32clipboard.CloseClipboard()
    BitMapInfoHeaderHandle = BITMAPINFOHEADER()
    memmove(pointer(BitMapInfoHeaderHandle), data, SIZEOF_BITMAPINFOHEADER)
    
    if BitMapInfoHeaderHandle.biCompression != BI_BITFIELDS:
        print('insupported compression type {}'.format(BitMapInfoHeaderHandle.biCompression))
        sys.exit(1)
        
    BitMapFileHeaderHandle = BITMAPFILEHEADER()
    memset(pointer(BitMapFileHeaderHandle), 0, SIZEOF_BITMAPFILEHEADER)
    BitMapFileHeaderHandle.bfType = ord('B') | (ord('M') << 8)
    BitMapFileHeaderHandle.bfSize = SIZEOF_BITMAPFILEHEADER + len(data)
    SIZEOF_COLORTABLE = 0
    BitMapFileHeaderHandle.bfOffBits = SIZEOF_BITMAPFILEHEADER + SIZEOF_BITMAPINFOHEADER + SIZEOF_COLORTABLE
    with open(TEMP_FILE_NAME, 'ab') as bmp_file:
        bmp_file.write(BitMapFileHeaderHandle)
        bmp_file.write(data)
    #保存为系统统一的png格式，使得后续可以识别
    im = Image.open(TEMP_FILE_NAME)
    im.save(TEMP_FILE_NAME)
    print('file created from clipboard image')  #完成图像创建

def baiduOCR():
    """利用百度api识别文本，并保存提取的文字
    """


    APP_ID = '16999492' # 刚才获取的 ID，下同
    API_KEY = 'nHwvwjvyAxRX3GR7Ww2NAkNd'
    SECRECT_KEY = '7G1cOkYFTZvBwGO25rGuachY2g5SHk49'
    client = AipOcr(APP_ID, API_KEY, SECRECT_KEY)
    
    i = open(TEMP_FILE_NAME, 'rb')
    img = i.read()
    print("正在识别截图")
    message = client.basicGeneral(img)   # 通用文字识别，每天 50 000 次免费
    #message = client.basicAccurate(img)   # 通用文字高精度识别，每天 800 次免费
    i.close();
    print("识别成功！")
    print(message)
    print("================s=================================================================")
    with open(OUTPUT_TXT, 'a+', encoding="utf-8") as fo:
        for text in message.get("words_result"):
            print(text.get("words"))
            fo.writelines(text.get('words') + '\n')
    print("=================================================================================")
    print("已经保存到export.txt")

def printScreen(filename):
    time.sleep(5)
    try:
        win32api.keybd_event(0x91, 0, 0, 0)    # https://msdn.microsoft.com/en-us/library/windows/desktop/dd375731(v=vs.85).aspx   0x91 --> win key
        win32api.keybd_event(0x2C, 0, 0, 0)    # 0x2C --> PRINT SCREEN key
        win32api.keybd_event(0x91, 0, win32con.KEYEVENTF_KEYUP, 0)    # https://msdn.microsoft.com/en-us/library/windows/desktop/ms646304(v=vs.85).aspx
        win32api.keybd_event(0x2C, 0, win32con.KEYEVENTF_KEYUP, 0)
    except:
        print('keyboard event does not successful.')
        sys.exit(1)
    
    BI_BITFIELDS = 3
    
    win32clipboard.OpenClipboard()    # https://msdn.microsoft.com/zh-cn/library/windows/desktop/ff468802(v=vs.85).aspx
    try:
        if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_DIB):    # https://msdn.microsoft.com/en-us/library/windows/desktop/ms649013(v=vs.85).aspx
            data = win32clipboard.GetClipboardData(win32clipboard.CF_DIB)
        else:
            print('cliboard does not contain an image in DIB format.')
            sys.exit(1)
    finally:
        win32clipboard.CloseClipboard()
        
    BitMapInfoHeaderHandle = BITMAPINFOHEADER()
    memmove(pointer(BitMapInfoHeaderHandle), data, SIZEOF_BITMAPINFOHEADER)
    
    if BitMapInfoHeaderHandle.biCompression != BI_BITFIELDS:
        print('insupported compression type {}'.format(BitMapInfoHeaderHandle.biCompression))
        sys.exit(1)
        
    BitMapFileHeaderHandle = BITMAPFILEHEADER()
    memset(pointer(BitMapFileHeaderHandle), 0, SIZEOF_BITMAPFILEHEADER)
    BitMapFileHeaderHandle.bfType = ord('B') | (ord('M') << 8)
    BitMapFileHeaderHandle.bfSize = SIZEOF_BITMAPFILEHEADER + len(data)
    SIZEOF_COLORTABLE = 0
    BitMapFileHeaderHandle.bfOffBits = SIZEOF_BITMAPFILEHEADER + SIZEOF_BITMAPINFOHEADER + SIZEOF_COLORTABLE
    
    with open(filename, 'wb') as bmp_file:
        bmp_file.write(BitMapFileHeaderHandle)
        bmp_file.write(data)
    print('file "{}" created from clipboard image'.format(filename))
    
def getClipboardText():
    w.OpenClipboard()
    d = w.GetClipboardData(win32con.CF_TEXT)
    w.CloseClipboard()
    return(d).decode('GBK')

def setClipboardText(aString):
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32con.CF_TEXT, aString)
    win32clipboard.CloseClipboard()

def init_temp_file():
    if path.exists(TEMP_FILE_NAME):
        os.remove(TEMP_FILE_NAME)
    if path.exists(OUTPUT_TXT):
        os.remove(OUTPUT_TXT)

if __name__ == "__main__":
    init_temp_file()
    getClipboardImage()
    baiduOCR()
    mystring = ""
    with open(OUTPUT_TXT, 'r', encoding="utf-8") as text:
        mystring += text.readline()
        mystring += '\n'
    #     for text in message.get("words_result"):
    #         print(text.get("words"))
    #         fo.writelines(text.get('words') + '\n')
    # all_the_text = open(OUTPUT_TXT,'r',encoding="utf-8").read()
    setClipboardText(mystring)
    