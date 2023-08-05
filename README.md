# auto-join-tencent-class-and-meet
计划性的进入腾讯课堂（网页）或腾讯会议（软件）  
第一次写不够成熟
# 1 准备
在运行源码前请确保安装了pywin32 pyautogui helium selenium openpyxl 等库  
当然，到时候我也许会去打包一下
# 2 启动
## 2.1 设置
若启动目录下没有plan_config.json 文件则会创建一个默认配置文件
默认配置中不会启用浏览器

DeBug: 查看一些没什么用的信息
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

这里推荐把ClassAutoClickLogin设为false，毕竟你可能是微信登录的  
当然loadconfig.py中的注释也有该内容  
## 2.2 腾讯课堂登录
在配置完成后会打开chrome，并在chrome中完成腾讯课堂的登录与cookies的录入  
如果您配置文件中的ClassAutoClickLogin为true的话将会点击(LoginClickX,LoginClickY)在屏幕中位置
注：cookies将会录入在目录下的login_cookies.json文件中，如果您发现登录信息过期的话请删除该文件

## 2.3 腾讯会议
程序1.1版采用腾讯会议自带的 wemeet:// 协议进入会议，之前的键鼠操作没有删除，可调设置改回（就是我不想删，反正不影响）
# 3 计划文件
## 3.1 JSON加载
JSON文件有示例不难理解  
名称不限，必须在plansfolder目录下
## 3.2 xlsx文件
首先，您需要在plansfolder创建xl.xlsx与xlsx_keys.json  
在xl.xlsx中：
1. A1格可以为任意数值
2. A列必须为时间（格式'xx:xx'）
3. 第一行需为星期且为英文（Monday）
4. 其他格可为空  
## 3.3 文件管理
于1.4更新，在小功能中点击任务管理
