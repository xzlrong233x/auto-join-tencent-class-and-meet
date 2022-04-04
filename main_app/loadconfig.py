import json,os
import win32api,win32con

'''
Load config file   
加载配置文件,如果没有将创建默认配置文件 
配置文件每个值的介绍:
DeBug: 还不知道用来做什么 
LoadXlsxEnable: 是否默认通过xlsx文件加载计划
WebEnable: 是否启用浏览器,如果为False将不会进行与腾讯课堂相关的操作
ClassAutoClickLogin: 是否在无法加载cookies时启用自动登录,目前只支持自动QQ登录
UseProtocolJoinMeeting: 是否使用腾讯会议的协议进入会议 原本的键盘鼠标操作不会消失可通过调整这个设置是否启用(懒得删.jpg)
LoginClickX: 以屏幕为参照物,自动登录点击的X坐标(默认值是在分辨率为1920x1080下)
LoginClickY: 自动登录点击的Y坐标
GourlTimeoutSec: 如果这节课未下课将等待的时间
MeetingJoinName: 您进入会议的名称
MeetingCloseWaitTime: 同GourlTimeoutSec
MeetingNameClickX: 由于识别名称较难所以将直接点击位置,此为X坐标
MeetingNameClickY: 点击处的Y坐标 
CloseButtonClickX: 会议退出按钮的X坐标(关闭前会议会最大化)
CloseButtonClickY: 会议退出按钮的Y坐标
CursorMoveWaitTime: 每次鼠标运动间隔的时间
'''


if not os.path.exists(r'plan_config.json'):
    PlanConfig = {'DeBug': False,
    'LoadXlsxEnable': False, 
    'WebEnable': True, 
    'ClassAutoClickLogin': True,
    'UseProtocolJoinMeeting': True,
    'LoginClickX': 700, 
    'LoginClickY': 600, 
    'GourlTimeoutSec': 360, 
    'MeetingJoinName': None, 
    'MeetingCloseWaitTime': 360, 
    'MeetingNameClickX': 820, 
    'MeetingNameClickY': 320, 
    'CloseButtonClickX': 1010,
    'CloseButtonClickY': 590,
    'CursorMoveWaitTime': 15}
    with open(r'plan_config.json','a+',encoding='utf-8') as f:
        f.write(json.dumps(PlanConfig))
    win32api.MessageBox(0,'加载配置文件出错,将使用默认配置\n如果你是第一次使用可以打开plan_config文件进行调试','注意',win32con.MB_OK)
with open(r'plan_config.json','r+',encoding='utf-8') as f:
    PlanConfig = json.loads(f.read())
#test
if __name__ == '__main__':
    print(PlanConfig)
