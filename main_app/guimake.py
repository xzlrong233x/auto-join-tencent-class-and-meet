import tkinter
import tkinter.filedialog
from tkinter import ttk
import re

from util import *
from pagectrl import *
from wemeetingctrl import *
from loadconfig import PlanConfig, DEFAULT, VERSION

global FileManager
FileManager = FilesClass()

class ConfigWindow:
    def __init__(self) -> None:
        self.win = None
        self.dic = {}
        self.BOOL = ["False","True"]
        self.checked = False

    def BtnTextChange(self,event):
        b = self.win.focus_get()
        if (b["text"] == "True"): b["text"] = "False"
        elif (b["text"] == "False"): b["text"] = "True"

    def save(self):
        global PlanConfig
        dc = {}
        for k,i in self.dic.items():
            l = ""
            if (isinstance(i,ttk.Button)):
                l = bool(self.BOOL.index(i["text"]))
            elif (isinstance(i,ttk.Spinbox)):
                l = int(re.search("(\d+)",i.get()).group(1))
            elif (isinstance(i,ttk.Entry)):
                l = i.get()
                if (l == ""):
                    l = None
            else: l = 0
            dc[k] = l
        PlanConfig = dc
        print(PlanConfig)
        with open(r'plan_config.json','w+',encoding='utf-8') as f:
            f.write(json.dumps(PlanConfig, indent= 4))
        self.checked = True
        self.win.destroy()                  

    def startWindow(self):
        self.win = tkinter.Tk()
        self.win.title("设置")
        Grow = 0

        for k,i in PlanConfig.items():
            c1 = tkinter.Label(self.win,text=k)
            c1.grid(row=Grow,column=0)
            k1 = None
            if (isinstance(i,bool)):
                k1 = ttk.Button(self.win,text=self.BOOL[int(i)])
                k1.bind("<ButtonRelease-1>",self.BtnTextChange)
            elif (isinstance(i,int)):
                k1 = ttk.Spinbox(self.win,from_=0,to=9999999)
                k1.set(i)
            elif (isinstance(i,str) or i is None):
                k1 = ttk.Entry(self.win,width=10)
                if (i is None):i = ""
                k1.insert(0,i)
            else:
                k1 = ttk.Label(self.win,text="Unknown")
            k1.grid(row=Grow,column=2)
            self.dic[k] = k1
            Grow += 1
        ttk.Button(self.win,text="保存",command=self.save).grid(row=Grow+1,column=2)
        self.win.mainloop()


class FileWindow:
    def __init__(self):
        pass

    def change(self,event=None,index:str=None):
        EnterCodeRecord = getRecord("EnterCode")
        wi = tkinter.Tk()
        ttk.Label(wi,text="星期").grid(row=1,column=0)
        en1 = ttk.Combobox(wi,width=50,values=("Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"),state="readonly")
        en1.current(0)
        en1.grid(row=1,column=1)
        ttk.Label(wi,text="时间").grid(row=2,column=0)
        en2 = ttk.Combobox(wi,width=8,values=['0'*(2-len(str(i))) + str(i) for i in range(24)],state="readonly")
        en2.grid(row=2,column=1)
        en2_1 = ttk.Combobox(wi,width=8,values=['0'*(2-len(str(i))) + str(i) for i in range(60)],state="readonly")
        en2_1.grid(row=2,column=2)
        ttk.Label(wi,text="进入码").grid(row=3,column=0)
        en3 = ttk.Combobox(wi,width=50,values=EnterCodeRecord)
        en3.grid(row=3,column=1)
        self._fg = None
        self._te = ''
        if index is None or index not in FileManager.contect_dict:
            self._te = "文件名"
            self._fg = ttk.Entry(wi,width=50)
            en2.current(int(time.strftime("%H")))
            en2_1.current(int(time.strftime("%M")))
        else:
            self._te = "地址"
            self._fg = ttk.Label(wi,text=index,width=50)
            en1.current(["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"].index(FileManager.contect_dict[index]["week"]))
            en2.current(int(FileManager.contect_dict[index]["time"][0:2]))
            en2_1.current(int(FileManager.contect_dict[index]["time"][3:5]))
            en3.insert(0,FileManager.contect_dict[index]["entercode"])
        self._fg.grid(row=0,column=1)
        ttk.Label(wi,text=self._te).grid(row=0,column=0)
        def savetofile():
            global enter
            if str(en3.get()).isnumeric():
                enter = int(en3.get())
            else:
                enter = en3.get()

            if str(enter) not in EnterCodeRecord:
                writeRecord("EnterCode",str(enter))

            fileName = ""
            if index is None or index not in FileManager.contect_dict:
                if self._fg.get().isspace() or self._fg.get() == "":
                    fileName = f'{["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"].index(en1.get())+1}-{len(FileManager.getByWeek(en1.get()))+1}.json'
                else:
                    fileName = self._fg.get()

            pathh = loadfolder
            if index is None or index not in FileManager.contect_dict:
                if fileName.split('.')[-1] == 'json':
                    pathh = pathh + '\\' + fileName
                else:
                    pathh = pathh + '\\' + fileName + '.json'
            else:
                pathh = index

            FileManager.write(pathh,en1.get(),(f"{en2.get()}:{en2_1.get()}"),enter)
            FileManager.save()
            wi.destroy()

        ttk.Button(wi,text="保存",command=savetofile).grid(row=4,column=0)

        wi.mainloop()

    def filesearch(self,strr:str) -> dict:
        if "@" not in strr:
            return FileManager.getByEntercode(strr)
        else:
            re_str = strr.split("@")[-1]
            match re_str:
                case "time":
                    return FileManager.getByTime(strr)
                case "enter" | "code" | "entercode":
                    return FileManager.getByEntercode(strr)
                case "week":
                    return FileManager.getByWeek(strr)
                case _:
                    return FileManager.getByEntercode(strr)
            
    def update(self):
        while 1:
            try:
                if self.winen.get() == "":
                    if self.wincomb.get() != "All":
                        dict1 = FileManager.getByWeek(self.wincomb.get())
                        self.ListBL.set(list(dict1.keys()))
                    else:
                        dict1 = FileManager.contect_dict
                        self.ListBL.set(list(dict1.keys()))
                
                else:
                    self.ListBL.set(list(self.filesearch(self.winen.get())))
                FileManager.refuse()

            except:pass
            time.sleep(0.1)

    def load(self,event,file:str=None):
        if file is None:
            file = tkinter.filedialog.askopenfilename(parent=self.win)
        if file.split('.')[-1] == "json":
            with open(file,"r+",encoding='utf-8') as fi:
                con = json.loads(fi.read())
                if "week" in con and "time" in con and "entercode" in con:
                    FileManager.write(file,con["week"],con["time"],con["entercode"])
                else:
                    pyautogui.alert("文件缺少部分内容")

    def getcurrent(self,event):
        fil = None
        for i in event.widget.curselection():
            fil = event.widget.get(i)
        if "week" in FileManager.contect_dict[fil] and "time" in FileManager.contect_dict[fil] and "entercode" in FileManager.contect_dict[fil]:
            self.change(index=fil)
        else:
            pyautogui.alert("文件缺少部分内容")

    def window(self):
        self.win = tkinter.Tk()
        self.win.geometry('480x640')
        self.ListBL = tkinter.StringVar(self.win)
        self.ComBoV = tkinter.StringVar(self.win)
        
        ttk.Label(self.win,text="星期").place(relx=0.03,rely=0,relheight=0.05,relwidth=0.08)
        self.wincomb = ttk.Combobox(self.win,values=("All","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"),state="readonly")
        self.wincomb.current(0)
        self.wincomb.place(relx=0.11,rely=0,relheight=0.05,relwidth=0.86)
        ttk.Label(self.win,text="搜索").place(relx=0.03,rely=0.05,relwidth=0.08,relheight=0.05)
        self.winen = ttk.Entry(self.win)
        self.winen.place(in_=self.wincomb,relx=0,rely=1.05,relheight=1,relwidth=1)
        winbox = tkinter.Listbox(self.win,selectmode=tkinter.SINGLE,listvariable=self.ListBL)
        winbox.place(relx=0.03,rely=0.108,relheight=0.85,relwidth=0.94)

        AddMenu = tkinter.Menu(self.win)

        ADD = tkinter.Menu(AddMenu,tearoff=False)
        ADD.add_command(label="添加内容",command=self.change)
        ADD.add_command(label="加载",command=self.load)
        AddMenu.add_cascade(menu=ADD,label="文件")

        self.win.bind("<Control-o>",self.load)
        self.win.bind("<Control-e>",self.change)
        self.win.bind("<Control-O>",self.load)
        self.win.bind("<Control-e>",self.change)
        winbox.bind("<Double-Button-1>",self.getcurrent)

        upp = threading.Thread(target=self.update,daemon=True)
        upp.start()

        self.win.config(menu=AddMenu)

        self.win.mainloop()


class ControlGUI:
    def __init__(self):
        '''初始化'''
        self.MainWindow = tkinter.Tk()
        self.MainWindow.title('控制界面')
        self.MainWindow.geometry('420x240')
        self.MainWindow.resizable(0,0)
        self.ConfigWindow = ConfigWindow()
        self.meetctrl = TxMeetingAutoControl()
        self.pagectrl = None
        if PlanConfig['WebEnable']: 
            self.pagectrl = PageCtrl()
            self.pagectrl.login_ctrl() 
        self.closerPlanStr = tkinter.StringVar()
        self.IfCursorMove,self.IfToastRemind = tkinter.IntVar(),tkinter.IntVar()

    def exit(self):
        if PlanConfig['WebEnable']: self.pagectrl.exit()
        FileManager.save()
        self.MainWindow.destroy()

    def CursorMove(self):
        '''移动鼠标'''
        while True:
            try:
                if self.IfCursorMove.get() == 1:
                    win32api.mouse_event(win32con.MOUSEEVENTF_MOVE, int(1), int(1), 0, 0)
                time.sleep(PlanConfig['CursorMoveWaitTime'])
            except:pass
    
    def GetcloserTime(self):
        '''获取最近的时间'''
        for i in self.timelist:
            if closerTimeCheck(i):
                return i
        return None

    def RedirectStart(self,*arg):
        '''检查时间'''
        while True:
            #print(time.strftime('%H:%M'),self.getTime,time.strftime('%H:%M') == self.getTime,type(self.plansDict[self.getTime]),type(self.plansDict[self.getTime]) == int)
            if time.strftime('%H:%M') == self.getTime:
                if type(self.plansDict[self.getTime]) == int:  #当进入码为整形时定向到腾讯会议
                    print("meeting start")
                    #启动线程
                    meetingStart = threading.Thread(target=self.meetctrl.ToStartMeeting(meetingCode=self.plansDict[self.getTime],stopWait=checkListContent(self.getTime,self.plansDict)))
                    meetingStart.start()
                elif type(self.plansDict[self.getTime]) == str and self.plansDict[self.getTime][:27] == 'https://ke.qq.com/webcourse': #当进入码为str且前27个字符为 https://ke.qq.com/webcourse 时定向腾讯课堂
                    if PlanConfig['WebEnable']:
                        #启动线程
                        pageStart = threading.Thread(target=self.pagectrl.go_url(self.plansDict[self.getTime]))
                        pageStart.start()
                    else:
                        os.system('start %s'%(str(self.plansDict[self.getTime])))
                else: #上述未通过时尝试使用start
                    print("else start")
                    os.system('start %s'%(str(self.plansDict[self.getTime])))
            self.updateVarible()
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
        self.plansDict = plans_info_get()
            
        self.timelist = list(self.plansDict.keys())
        self.timelist.sort()
        self.getTime = self.GetcloserTime()
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
            if (self.ConfigWindow.checked):
                if PlanConfig['WebEnable']: 
                    self.pagectrl = PageCtrl()
                    self.pagectrl.login_ctrl()
                elif (self.pagectrl is not None):
                    self.pagectrl.exit()
                    self.pagectrl = None
                self.ConfigWindow.checked = False
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

    def Info(self):
        def start(*args):
            os.system("start https://github.com/xzlrong233x/auto-join-tencent-class-and-meet/")
        win = tkinter.Tk()
        win.title("关于")
        win.geometry("460x120")
        tkinter.Label(win,text="开源地址：").grid(row=1,column=1)
        label = tkinter.Label(win,text="https://github.com/xzlrong233x/auto-join-tencent-class-and-meet/",fg="blue",cursor="hand2")
        label.grid(row=1,column=2)
        label.bind("<Button-1>",start)
        tkinter.Label(win,text=f"版本：{VERSION}").grid(row=2,column=1)
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
        ttk.Label(self.MainWindow,textvariable=self.closerPlanStr,wraplength=160,takefocus=True).grid(row=0,column=1)
        #set button
        ttk.Button(self.MainWindow,text='显示全部任务',command=self.ShowAllPlans,width=10).grid(row=1,column=0,padx=10,pady=8)
        ttk.Button(self.MainWindow,text='刷新',command=self.updateVarible,width=10).grid(row=1,column=1,padx=10,pady=8)
        self.StartbySelfButton = ttk.Button(self.MainWindow,text='手动启动',command=self.startReireThread,width=10)
        self.StartbySelfButton.grid(row=1,column=2,padx=10,pady=8)
        #set menu
        TopMenu = tkinter.Menu(self.MainWindow)

        SystemMenu = tkinter.Menu(TopMenu,tearoff=False)
        SystemMenu.add_command(label='刷新',command=self.updateVarible)
        SystemMenu.add_command(label='设置',command=self.ConfigWindow.startWindow)
        SystemMenu.add_separator()
        SystemMenu.add_command(label='退出',command=self.exit)
        TopMenu.add_cascade(menu=SystemMenu,label='系统')

        def filestart():
            cli = FileWindow()
            oio = threading.Thread(target=cli.window,daemon=True)
            oio.start()
        SmallTools = tkinter.Menu(TopMenu,tearoff=False)
        SmallTools.add_checkbutton(variable=self.IfCursorMove,label='暂停锁屏')
        SmallTools.add_command(label="任务管理",command=filestart)
        SmallTools.add_command(label='重载meetctrl',command=self.meetctrl.__init__)
        TopMenu.add_cascade(menu=SmallTools,label='小功能')

        if PlanConfig['DeBug']:
            DeBugMenu = tkinter.Menu(TopMenu,tearoff=False)
            DeBugMenu.add_command(label='查看详细情况',command=self.DeBugInfo) #不会有人认真看吧
            TopMenu.add_cascade(menu=DeBugMenu,label='DeBug')

        TopMenu.add_command(label='关于',command=self.Info)
        self.MainWindow.config(menu=TopMenu)
        self.MainWindow.protocol("WM_DELETE_WINDOW",self.exit)
        #start thread controls
        self.startReireThread()
        self.startThreads()
        #loop
        self.MainWindow.mainloop()
