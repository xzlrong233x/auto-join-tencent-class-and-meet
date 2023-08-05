'''
重点优化模块
'''
import os
import threading
import time

import pyautogui
import pyperclip
import uiautomation as auto
from loadconfig import PlanConfig
from windowcontrol import *

class TxMeetingAutoControl(WindowControl):
    def __init__(self):
        self.MainHandel = self.GetWindowHwnd(winName='腾讯会议',index=-1)
        if self.MainHandel == 0:
            os.system('start wemeet://')
            self.__init__()

    def ClickJoin(self,meetingCode):
        '''将内容输入并加入'''
        try:pyautogui.moveTo(pyautogui.center(pyautogui.locateOnScreen('join_meeting_button.png')),duration=0.1)
        except:
            win32api.MessageBox(0,'呀,出错了,请检查后台是否有其他会议,关闭后在重试','呀',win32con.MB_RETRYCANCEL)
            return self.ClickJoin(meetingCode)
        win32gui.SetWindowPos(self.MainHandel, win32con.HWND_NOTOPMOST, 0,0,0,0,win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
        pyautogui.click()
        #join and type name
        time.sleep(3)
        pyautogui.typewrite(str(meetingCode),interval=0.05)
        if PlanConfig['MeetingJoinName'] is not None:
            pyautogui.moveTo(PlanConfig['MeetingNameClickX'],PlanConfig['MeetingNameClickY'],duration=0.3)
            pyautogui.doubleClick()
            pyautogui.hotkey('ctrl','a')
            pyperclip.copy(PlanConfig['MeetingJoinName'])
            pyautogui.hotkey('ctrl','v')
        pyautogui.press('enter')

    def Join(self,meetingCode):
        '''加入会议'''
        if not PlanConfig['UseProtocolJoinMeeting']:
            self.ShowWindow(self.MainHandel)
            win32gui.SetWindowPos(self.MainHandel, win32con.HWND_TOPMOST, 0,0,0,0,win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            self.ClickJoin(meetingCode)
        else:
            os.system('start wemeet://page/inmeeting?meeting_code=%s^&rs=17'%(str(meetingCode)))
        print("meeting join")

    def CloseMeeting(self,hwnd):
        '''关闭会议'''
        self.ShowWindow(hwnd,showWindowMode=win32con.SW_SHOWMAXIMIZED)
        time.sleep(0.2)
        win32gui.PostMessage(hwnd,win32con.WM_CLOSE,0,0)
        time.sleep(0.8)
        try:
            button = auto.ButtonControl(searchDepth=7, Name="离开会议")
            button.Click()
        except LookupError as e:
            os.system("taskkill /f /t /im wemeetapp.exe")
            time.sleep(5)
            if not PlanConfig['UseProtocolJoinMeeting']:
                os.system("start wemeet://")
                time.sleep(5)

    def waitMeetingEnd(self,meetingCode,frequency):
        '''若当前会议未结束则等待'''
        while True:
            if win32gui.FindWindow(None,'腾讯会议') == self.MainHandel:
                self.Join(meetingCode)
                return None
            elif frequency == 0:
                self.CloseMeeting(self.GetWindowHwnd(winName='腾讯会议'))
                self.Join(meetingCode)
                return None
            frequency -= 1
            time.sleep(1)

    def ToStartMeeting(self,meetingCode,stopWait=False):
        '''到时间准备点击'''
        if win32gui.FindWindow(None,'腾讯会议') != self.MainHandel and not stopWait:
            self.runWait = threading.Thread(target=self.waitMeetingEnd,args=(meetingCode,PlanConfig['MeetingCloseWaitTime']),name='WaitThread')
            self.runWait.start()
            return None
        elif win32gui.FindWindow(None,'腾讯会议') == self.MainHandel:
            self.Join(meetingCode)
            return None
#test
if __name__ =='__main__':
    test = TxMeetingAutoControl()
    pyautogui.alert(text='test',title='test',button='OK')
    test.ToStartMeeting(7874257923)
    time.sleep(5)
    test.ToStartMeeting(7874257923)