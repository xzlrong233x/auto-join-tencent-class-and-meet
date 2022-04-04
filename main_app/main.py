#coding: utf-8
#定时自动进入课堂，支持腾讯会议，腾讯课堂（作者上网课的平台就这些
#运行源文件需pywin32 pyautogui helium selenium openpyxl win10toast库
#author: xlr_233
#from: https://github.com/xzlrong233x/auto-join-tencent-class-and-meet/
#version: 1.1
from guimake import ControlGUI

if __name__ == '__main__':
    gui = ControlGUI()
    gui.ControlsSet()
