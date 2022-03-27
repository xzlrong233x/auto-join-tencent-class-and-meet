# auto-join-tencent-class-and-meet
计划性的进入腾讯课堂（网页）或腾讯会议（软件）
# 1. 准备
在运行源码前请确保安装了pywin32 pyautogui helium selenium openpyxl win10toast（该库可有可无本人在测试发现显示不了消息不知是个人问题还是其他）等库
当然，到时候我也许会去打包一下
# 2.1 配置
在运行后目录下会多出一个名为plan_config.json的文件，并且会弹出一个窗口，该文件为配置文件，在运行前可以对其进行修改，这里给出值的作用：
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
# 2.2 腾讯课堂登录
有点晚了到时再写
