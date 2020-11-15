# -*- coding: utf-8 -*-
"""
Created on Sun Nov  8 22:14:50 2020

@author: buyibuyi
"""
from matplotlib import pyplot as plt
import win32gui,win32ui,win32con,win32api
#from PIL import Image
#import pytesseract
#import webbrowser
import os,time
import pyautogui as pag
#from aip import AipOcr
import cv2
#import json
#import numpy as np
from pymouse import PyMouse
from ctypes import *  # 获取屏幕上某个坐标的颜色
status=0
x1=0
x2=0
y1=0
y2=0
x_yinxiong=0
y_yinxiong=0
#下面两个坐标是用来确定是否回到每日答题页面的
x_top=0
y_top=0
def get_point():
    try:
        global x1,x2,y1,y2,x_yinxiong,y_yinxiong,x_top,y_top
        print('正在采集坐标1，请将鼠标移动到左上角')
        print(3)
        time.sleep(1)
        print(2)
        time.sleep(1)
        print(1)
        x1,y1=pag.position()
        print('采集成功，坐标为： ',(x1,y1))
        print('')
        
        
        print('正在采集坐标2，请将鼠标移动到右下角')
        print(3)
        time.sleep(1)
        print(2)
        time.sleep(1)
        print(1)
        x2,y2=pag.position()
        print('采集成功，坐标为： ',(x2,y2))
        print('')
        
        
        print('正在采集英雄篇坐标，请将鼠标移动到英雄篇上')
        print(3)
        time.sleep(1)
        print(2)
        time.sleep(1)
        print(1)
        x_yinxiong,y_yinxiong=pag.position()
        print('采集成功，坐标为： ',(x_yinxiong,y_yinxiong))
        print('')
        
        print('正在采集顶端标题坐标，请将鼠标移动到顶端红色标题上')
        print(3)
        time.sleep(1)
        print(2)
        time.sleep(1)
        print(1)
        x_top,y_top=pag.position()
        print('采集成功，坐标为： ',(x_top,y_top))
        print('')
        
        w=abs(x1-x2)
        h=abs(y1-y2)
        x=min(x1,x2)
        y=min(y1,y2)
        return (w,h,x,y)
    except KeyboardInterrupt:
        print('获取失败')
    
#获取截图
def window_capture(result,filename):
    #宽度w，高度h，左上角的坐标x，y
    w,h,x,y=result
    hwnd=0
    hwndDC=win32gui.GetWindowDC(hwnd)
    mfcDC=win32ui.CreateDCFromHandle(hwndDC)
    saveDC=mfcDC.CreateCompatibleDC()
    saveBitMap=win32ui.CreateBitmap()
    MonitorDev=win32api.EnumDisplayMonitors(None,None)
    saveBitMap.CreateCompatibleBitmap(mfcDC,w,h)
    saveDC.SelectObject(saveBitMap)
    saveDC.BitBlt((0,0),(w,h),mfcDC,(x,y),win32con.SRCCOPY)
    saveBitMap.SaveBitmapFile(saveDC,filename)
def get_color(x, y):
    gdi32 = windll.gdi32
    user32 = windll.user32
    hdc = user32.GetDC(None)  # 获取颜色值
    pixel = gdi32.GetPixel(hdc, x, y)  # 提取RGB值
    r = pixel & 0x0000ff
    g = (pixel & 0x00ff00) >> 8
    b = pixel >> 16
    return [r, g, b]
status='y'
start=time.time()
result=get_point()
window_capture(result,"jietu.jpg")
time.sleep(3)
img = cv2.imread("jietu.jpg")
#print (img)
#cv2.namedWindow("Image")
#cv2.imshow("Image", img)
m = PyMouse()
xhalf=(x1+x2)/2
yhalf=(y1+y2)/2
start_time=1
print("x1,y1:",(x1,y1),get_color(x1, y1))
print("x2,y2:",(x2,y2),get_color(x2, y2))
for j in range(0,1000):
    #滚到最上
    win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL,0,0,1)
    win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL,0,0,1)
    time.sleep(0.5)
    m.click(x_yinxiong,y_yinxiong,1,1)
    time.sleep(3)
    for i in range(start_time-1,100):
        color_point=get_color(x_top, y_top)
        if color_point[0]==176 and color_point[1]==0 and color_point[2]==12:
            print("已回到每日答题页面")
            break
        #默认从第一题开始
        start_time=1
        #先移动到最下，可以看到下一题的位置
        print("This is",j,"th time",i,"th question")
        pag.moveTo(xhalf, yhalf, 0.1)
        #pag.dragRel(0, -160, 2)
        win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL,0,0,-1)
        win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL,0,0,-1)
        time.sleep(1)
        #选择一个白色块点击
        x_tmp=int(xhalf-30)
        y_tmp=int(yhalf)
        win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL,0,0,-1)
        win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL,0,0,-1)
        time.sleep(0.05)
        #只选一个
        while 1:
                color_point=get_color(x_tmp, y_tmp)
                if (color_point[0]>250) and (color_point[1]>250) and (color_point[2]>250):
                    time.sleep(0.05)
                    print("White! break!")
                    break
                else:
                    y_tmp=y_tmp+10
                    time.sleep(0.05)
        m.click(x_tmp,y_tmp,1,1)
        #暂时不管多选的情况,如果要考虑多选用下面的代码
        '''
        if i<20:
            while 1:
                color_point=get_color(x_tmp, y_tmp)
                if (color_point[0]>250) and (color_point[1]>250) and (color_point[2]>250):
                    time.sleep(0.05)
                    print("White! break!")
                    break
                else:
                    y_tmp=y_tmp+10
                    time.sleep(0.05)
            m.click(x_tmp,y_tmp,1,1)
        else:
            while y_tmp<y2-30:
                color_point=get_color(x_tmp, y_tmp)
                if (color_point[0]>250) and (color_point[1]>250) and (color_point[2]>250):
                    time.sleep(0.05)
                    m.click(x_tmp,y_tmp,1,1)
                y_tmp=y_tmp+10
        '''
        
        #选择下一题
        x_tmp=int(xhalf)
        y_tmp=y2-30
        pag.moveTo(x_tmp, y_tmp, 0.1)
        while y_tmp>yhalf:
            color_point=get_color(x_tmp, y_tmp)
            #pag.moveTo(x_tmp, y_tmp, 0.1)
            if (color_point[0]>185) and (color_point[0]<210) and (color_point[1]>20) and (color_point[1]<35) and (color_point[2]>20) and (color_point[2]<40):
                time.sleep(0.05)
                print("red! click!")
                break
            else:
                time.sleep(0.05)
                y_tmp-=2
        m.click(x_tmp,y_tmp,1,1)
        time.sleep(3)
    #选择下一轮答题
    x_tmp=int(xhalf)
    y_tmp=y2-20
    pag.moveTo(x_tmp, yhalf, 0.1)
    win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL,0,0,-1)
    win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL,0,0,-1)
    while y_tmp>yhalf:
        color_point=get_color(x_tmp, y_tmp)
        #pag.moveTo(x_tmp, y_tmp, 0.1)
        if (color_point[0]>185) and (color_point[0]<210) and (color_point[1]>20) and (color_point[1]<35) and (color_point[2]>20) and (color_point[2]<40):
            time.sleep(0.05)
            print("red! click!")
            break
        else:
            time.sleep(0.05)
            y_tmp-=2
    m.click(x_tmp,y_tmp,1,1)
    time.sleep(3)


    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    