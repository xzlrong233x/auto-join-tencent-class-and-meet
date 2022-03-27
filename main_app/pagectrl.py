import json
import helium
import selenium
import time
import pyautogui
import win32api,win32con
from loadconfig import PlanConfig

class PageCtrl:
    '''open browser and to go to classes'''
    def __init__(self):
        self.dirver = helium.start_chrome()
        self.__tempVar = 0

    def login_ctrl(self):  #登录，尝试读取login_cookies.json中的cookies，如果失败会根据配置进行手动或自动登录
        __SaveCookiesFilePath = r'login_cookies.json'
        helium.go_to('ke.qq.com')
        try:
            with open(__SaveCookiesFilePath,'r+',encoding='utf-8') as f: GotCookies = json.loads(f.read())
            for cookie in GotCookies:
                self.dirver.add_cookie(cookie)
        except:
            helium.click(helium.Text('登录'))
            if PlanConfig['ClassAutoClickLogin']:
                time.sleep(5)
                pyautogui.click(PlanConfig['LoginClickX'],PlanConfig['LoginClickY'])
            elif win32api.MessageBox(0,'请确定你是否已登录\n是请按确定,否请按取消\n注:点击取消整个程序都会退出','确认',win32con.MB_OKCANCEL) == 2:
                self.exit()
                exit()
            time.sleep(10)
            DictCookies = self.dirver.get_cookies()
            SaveCookies = json.dumps(DictCookies)
            with open(__SaveCookiesFilePath,'a+',encoding='utf-8') as f:
                f.write(SaveCookies)
        helium.refresh()

    def go_url(self,gourl):   #该函数的等待机制未有过多测试（一天就一节腾讯课堂的课
        '''进入网站,如果未下课则等待一定时间,期间如下课则进入课堂,否则直接进入'''
        if self.__tempVar == 1:
            try:
                helium.wait_until(helium.Text('老师已经下课了！').exists,timeout_secs=PlanConfig['GourlTimeoutSec'])
                helium.go_to(gourl)
            except: helium.go_to(gourl)
        elif self.__tempVar == 0:
            helium.go_to(gourl)
            self.__tempVar = 1

    def check_boutton_and_click(self,sleepTime=2): #点击签到按钮
        while True:
            try:
                helium.click(helium.S('//*[@id="react-body"]/div[3]/div[2]/div/div[3]/span'))
            except:
                pass
            time.sleep(sleepTime)

    def exit(self):
        helium.kill_browser()