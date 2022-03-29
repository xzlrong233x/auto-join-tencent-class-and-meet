# auto-join-tencent-class-and-meet
计划性的进入腾讯课堂（网页）或腾讯会议（软件）  
第一次写不够成熟
# 1 准备
在运行源码前请确保安装了pywin32 pyautogui helium selenium openpyxl win10toast（该库可有可无本人在测试发现显示不了消息不知是个人问题还是其他）等库  
当然，到时候我也许会去打包一下
# 2 第一次启动
## 2.1 配置
在运行后目录下会创建一个名为plan_config.json的文件，并且会弹出一个窗口，该文件为配置文件，在运行前可以对其进行修改，这里给出值的作用： 

DeBug: 还不知道用来做什么  
LoadXlsxEnable: 是否默认通过xlsx文件加载计划  
WebEnable: 是否启用浏览器,如果为False将不会进行与腾讯课堂相关的操作  
ClassAutoClickLogin: 是否在无法加载cookies时启用自动登录,目前只支持自动QQ登录  
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

在xlsx_keys.json中：  
xlsx文件中有效读取区中的值都在xlsx_keys.json中有所对应

# 4 想说的话
## 4.1 大实话
暂停锁屏实质是等待一段时间后将光标移动一个像素  
腾讯课堂集成了自动签到  
重载是给你应对腾讯会议突然崩了的情况  
开启通知可能不起作用（作者本人的弹不出来）  
定向线程会将int值看作为腾讯会议的会议码 将前27个字符为https://ke.qq.com/webcourse 的字符串作为腾讯课堂的课堂url  
刷新是刷给你看的不会改变已创建的定向线程  
会优化的，等我再咕一咕  
## 4.2
第一次写这类的程序，很多库都是现学现写的；看时期的程序，日常非刚需，上网课时也就有点需求，就写来乐呵一下  
说到底，我只是一个很闲(?)想摸鱼的学生罢了
