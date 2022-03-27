import win32api,win32con,pythoncom,win32com.client,win32gui

class WindowControl:
    '''some package functions to control windows'''
    def ShowWindow(self,hwnd,showWindowMode=win32con.SW_SHOW):
        '''show window'''
        try:
            pythoncom.CoInitialize()
            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd,win32con.SW_SHOWNORMAL)
            shell = win32com.client.Dispatch("WScript.Shell")
            shell.SendKeys('%')
            win32gui.ShowWindow(hwnd,showWindowMode)
            win32gui.SetForegroundWindow(hwnd)
            pythoncom.CoUninitialize
        except Exception as err:
            if win32api.MessageBox(0,'显示窗口失败,请重试\n错误:%s'%(err),'错误',win32con.MB_RETRYCANCEL) ==4:
                self.ShowWindow(hwnd)
        
    def GetWindowHwnd(self,winName :str=None,winClass :str=None,index=0):
        '''function to get window hwnd from all windows\n
        winName can help you find window by name from all windows\n
        winClass can help you find window by class from all windows'''
        hwndlist = []
        def hwndjudge(hwnd,nonuse):
            if win32gui.IsWindow(hwnd) and win32gui.IsWindowEnabled(hwnd) and win32gui.IsWindowVisible(hwnd):
                if win32gui.GetWindowText(hwnd) == winName or win32gui.GetClassName(hwnd) == winClass:
                    hwndlist.append(hwnd)
        win32gui.EnumWindows(hwndjudge,0)
        if index is None: return hwndlist
        try: return hwndlist[index]
        except: return 0
#test
if __name__ == '__main__':
    s = WindowControl()
    MainHandel = s.GetWindowHwnd(winName='腾讯会议')
    print(MainHandel)
    print(win32gui.FindWindow(None,'腾讯会议'))
    s.ShowWindow(MainHandel,showWindowMode=win32con.SW_SHOWMAXIMIZED)