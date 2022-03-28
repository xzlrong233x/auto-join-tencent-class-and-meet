'''
重点优化模块
'''
import threading
import time

import pyautogui
import pyperclip
from loadconfig import PlanConfig
from windowcontrol import *

class TxMeetingAutoControl(WindowControl):
    def __init__(self):
        try: self.MainHandel = self.GetWindowHwnd(winName='腾讯会议',index=-1)
        except: self.MainHandel = 0
        if self.MainHandel == 0:
            pyautogui.alert(text='请打开腾讯会议',title='无法寻找到腾讯会议',button='OK')
            self.__init__()

    def ClickJoin(self,meetingCode):
        '''input info and join the meeting'''
        try:pyautogui.moveTo(pyautogui.center(pyautogui.locateOnScreen('join_meeting_button.png')),duration=0.1)
        except:
            win32api.MessageBox(0,'呀,出错了,请检查后台是否有其他会议,关闭后在重试','呀',win32con.MB_RETRYCANCEL)
            return self.ClickJoin(meetingCode)
        win32gui.SetWindowPos(self.MainHandel, win32con.HWND_NOTOPMOST, 0,0,0,0,win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
        pyautogui.click()
        #join and type name
        time.sleep(3)#等待前一个会议码出现
        pyautogui.typewrite(str(meetingCode),interval=0.05)
        if PlanConfig['MeetingJoinName'] is not None:
            pyautogui.moveTo(PlanConfig['MeetingNameClickX'],PlanConfig['MeetingNameClickY'],duration=0.3)
            pyautogui.doubleClick()
            pyautogui.hotkey('ctrl','a')
            pyperclip.copy(PlanConfig['MeetingJoinName'])
            pyautogui.hotkey('ctrl','v')
        pyautogui.press('enter')

    def Join(self,meetingCode):
        '''function to show window and join meeting'''
        self.ShowWindow(self.MainHandel)
        win32gui.SetWindowPos(self.MainHandel, win32con.HWND_TOPMOST, 0,0,0,0,win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
        self.ClickJoin(meetingCode)

    def CloseMeeting(self,hwnd):
        '''close meeting'''
        self.ShowWindow(hwnd,showWindowMode=win32con.SW_SHOWMAXIMIZED)
        time.sleep(0.2)
        win32gui.PostMessage(hwnd,win32con.WM_CLOSE,0,0)
        time.sleep(0.8)
        pyautogui.click(PlanConfig['CloseButtonClickX'],PlanConfig['CloseButtonClickY'])

    def waitMeetingEnd(self,meetingCode,frequency):
        '''wait meeting end or close meeting then run join'''
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

    def ToStartMeeting(self,meetingCode):
        '''check if the time is compliant'''
        if win32gui.FindWindow(None,'腾讯会议') != self.MainHandel:
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
    test.ToStartMeeting(12345674444)
