try: from win10toast import ToastNotifier
except: pass
import tkinter

from commonfunctions import *
from pagectrl import *
from wemeetingctrl import *


class ControlGUI:
    '''create a gui with many functions'''
    def __init__(self):
        '''def window and variable'''
        self.MainWindow = tkinter.Tk()
        self.MainWindow.title('控制界面')
        self.MainWindow.geometry('420x240')
        self.MainWindow.resizable(0,0)
        self.meetctrl = TxMeetingAutoControl()
        if PlanConfig['WebEnable']:
            self.pagectrl = PageCtrl()
            self.pagectrl.login_ctrl()
        self.ToastRemind = ToastNotifier()
        self.closerPlanStr = tkinter.StringVar()
        self.IfCursorMove,self.IfToastRemind,self.IfLoadFromXlsx = tkinter.IntVar(),tkinter.IntVar(),tkinter.IntVar()
        if PlanConfig['LoadXlsxEnable']:
            self.IfLoadFromXlsx.set(1)

    def exit(self):
        if PlanConfig['WebEnable']: self.pagectrl.exit()
        self.MainWindow.destroy()

    def CursorMove(self):
        '''move cursor every once in a while'''
        while True:
            if self.IfCursorMove.get() == 1:
                win32api.mouse_event(win32con.MOUSEEVENTF_MOVE, int(1), int(1), 0, 0)
            time.sleep(PlanConfig['CursorMoveWaitTime'])
    
    def GetcloserTime(self):
        '''package func to get closer time'''
        for i in self.timelist:
            if closerTimeCheck(i):
                return i
        if self.IfToastRemind.get() == 1: 
            try: self.ToastRemind.show_toast('没有任务了','无法检测到新的任务',icon_path=None,duration=5,threaded=True)
            except: pass
        return None

    def RedirectStart(self,retime=None,EnterCode=None):
        '''check time to start'''
        while True:
            if time.strftime('%H:%M') == retime:
                if type(EnterCode) == int:  #当进入码为int时定向到腾讯会议
                    if self.IfToastRemind.get() == 1: 
                        try: self.ToastRemind.show_toast('定向腾讯会议','请暂且不要操作电脑\n进入码:%s'%(EnterCode),duration=10,icon_path=None,threaded=True)
                        except: pass
                    #启动线程
                    meetingStart = threading.Thread(target=self.meetctrl.ToStartMeeting(meetingCode=EnterCode))
                    meetingStart.start()
                    self.startReireThread()
                    return None
                elif type(EnterCode) == str and EnterCode[:27] == 'https://ke.qq.com/webcourse': #当进入码为str且前27个字符为 https://ke.qq.com/webcourse 时定向腾讯课堂
                    if PlanConfig['WebEnable']:
                        if self.IfToastRemind.get() == 1: 
                            try: self.ToastRemind.show_toast('定向腾讯课堂','进入码:%s'%(EnterCode),duration=10,icon_path=None,threaded=True)
                            except: pass
                        #启动线程
                        pageStart = threading.Thread(target=self.pagectrl.go_url(EnterCode))
                        pageStart.start()
                    self.startReireThread()
                    return None
            time.sleep(1)

    def startReireThread(self):
        '''启动主要的定向线程'''
        self.updateVarible()
        if self.getTime is None:
            return None
        self.ReireThread = threading.Thread(target=self.RedirectStart,args=(self.getTime,self.plansDict[self.getTime]),name='ReireThread',daemon=True)
        self.ReireThread.start()

    def updateVarible(self):
        '''更新任务内容'''
        if self.IfLoadFromXlsx.get() == 1:
            self.plansDict = plans_get_from_xlsx()
            if self.plansDict is None: 
                self.plansDict = plans_info_get()
                self.IfLoadFromXlsx.set(0)
        else:
            self.plansDict = plans_info_get()
            
        self.timelist = list(self.plansDict.keys())
        self.timelist.sort()
        self.getTime = checkListContent(self.GetcloserTime(),self.timelist,self.plansDict)
        try:
            self.closerPlanStr.set('时间:%s\n进入码:%s'%(self.getTime,self.plansDict[self.getTime]))
        except:
            self.closerPlanStr.set('暂无任务')

    def SystemCheck(self):
        '''主要的系统检查'''
        while True:
            if time.strftime('%H:%M:%S') == '00:00:00':
                self.updateVarible()
                self.startReireThread()
            try:
                if self.ReireThread.is_alive(): self.StartbySelfButton.config(state='disabled')
                else: self.StartbySelfButton.config(state='normal')
            except: pass
            time.sleep(0.01)

    def startThreads(self):
        '''如函数名'''
        self.MoveThread = threading.Thread(target=self.CursorMove,name='MoveThread',daemon=True)
        self.OverCheck = threading.Thread(target=self.SystemCheck,name='SystemCheck',daemon=True)
        self.MoveThread.start()
        self.OverCheck.start()
        if PlanConfig['WebEnable']:
            self.AutoClick = threading.Thread(target=self.pagectrl.check_boutton_and_click,name='autoclick',daemon=True)
            self.AutoClick.start()

    def ShowAllPlans(self):
        '''输出能读取到的所有任务'''
        win = tkinter.Tk()
        win.title('全部任务')
        text = '全部任务:\n'
        for i in self.timelist:
            text = text + '时间:%s\n进入码:%s\n\n'%(i,self.plansDict[i])
        tkinter.Label(win,text=text).pack()
        win.mainloop()

    def DeBugInfo(self):
        '''显示详细信息包括线程状态,当前腾讯会议指向的句柄'''
        wind = tkinter.Tk()
        wind.title('详细信息')
        wind.geometry('300x200')
        ThreadStateText = '\n'
        try:ThreadStateText = ThreadStateText+'%s:%s\n'%(self.ReireThread.getName(),str(self.ReireThread.is_alive()))
        except:pass
        try:ThreadStateText = ThreadStateText+'%s:%s\n'%(self.AutoClick.getName(),str(self.AutoClick.is_alive()))
        except:pass
        ThreadStateText = ThreadStateText+'%s:%s\n'%(self.MoveThread.getName(),str(self.MoveThread.is_alive()))
        ThreadStateText = ThreadStateText+'%s:%s\n'%(self.OverCheck.getName(),str(self.OverCheck.is_alive()))
        ThreadStateText = ThreadStateText+'会议句柄:%s\n'%(str(self.meetctrl.MainHandel))
        tkinter.Label(wind,text=ThreadStateText,takefocus=True,justify='left').grid()
        wind.mainloop()

    def ControlsSet(self):
        #set Label
        tkinter.Label(self.MainWindow,textvariable=self.closerPlanStr,wraplength=160,takefocus=True).grid(row=0,column=1)
        #set button
        tkinter.Button(self.MainWindow,text='显示全部任务',command=self.ShowAllPlans,width=10).grid(row=1,column=0,padx=10,pady=8)
        tkinter.Button(self.MainWindow,text='刷新',command=self.updateVarible,width=10).grid(row=1,column=1,padx=10,pady=8)
        self.StartbySelfButton = tkinter.Button(self.MainWindow,text='手动启动',command=self.startReireThread,width=10)
        self.StartbySelfButton.grid(row=1,column=2,padx=10,pady=8)
        #set menu
        TopMenu = tkinter.Menu(self.MainWindow)

        SystemMenu = tkinter.Menu(TopMenu,tearoff=False)
        SystemMenu.add_command(label='刷新',command=self.updateVarible)
        SystemMenu.add_separator()
        SystemMenu.add_command(label='退出',command=self.exit)
        TopMenu.add_cascade(menu=SystemMenu,label='系统')

        SmallTools = tkinter.Menu(TopMenu,tearoff=False)
        SmallTools.add_checkbutton(variable=self.IfCursorMove,label='暂停锁屏')
        SmallTools.add_checkbutton(variable=self.IfToastRemind,label='开启通知')
        SmallTools.add_command(label='重载meetctrl',command=self.meetctrl.__init__)
        SmallTools.add_command(label='查看详细情况',command=self.DeBugInfo)
        TopMenu.add_cascade(menu=SmallTools,label='小功能')

        LoadModeChoose = tkinter.Menu(TopMenu,tearoff=False)
        LoadModeChoose.add_radiobutton(value=1,variable=self.IfLoadFromXlsx,label='xlsx加载',command=self.updateVarible)
        LoadModeChoose.add_radiobutton(value=0,variable=self.IfLoadFromXlsx,label='默认加载')
        TopMenu.add_cascade(menu=LoadModeChoose,label='加载模式')

        if PlanConfig['DeBug']:
            DeBugMenu = tkinter.Menu(TopMenu,tearoff=False)
            TopMenu.add_cascade(menu=DeBugMenu,label='DeBug')

        TopMenu.add_command(label='信息')
        self.MainWindow.config(menu=TopMenu)
        self.MainWindow.protocol("WM_DELETE_WINDOW",self.exit)
        #start thread controls
        self.startReireThread()
        self.startThreads()
        #loop
        self.MainWindow.mainloop()