import time
import pyperclip
import win32api
import win32con
import sys
import os

# 判断
while True:
    time_now = time.strftime("%H:%M", time.localtime())  # 获取当前时间
    if time_now == "07:40":  # 此处为消息发送的时间
        def open_app(app_dir):
            os.startfile(app_dir)


        if __name__ == "__main__":
            app_dir = r'C:\Program Files (x86)\Tencent\WeChat\WeChat.exe'  # 此处为微信的绝对路径
            open_app(app_dir)
            time.sleep(1)  # 电脑反应需要时间，使程序暂停一s时间来等待电脑反应
            # 使用crtl + F打开微信搜索框
            win32api.keybd_event(17, 0, 0, 0)
            win32api.keybd_event(70, 0, 0, 0)
            win32api.keybd_event(70, 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)
            time.sleep(1)

            pyperclip.copy('C楼B座留观楼')  # 此处存放微信好友的备注名
            pyperclip.paste()
            win32api.keybd_event(17, 0, 0, 0)  # ctrl键码
            win32api.keybd_event(86, 0, 0, 0)  # v的键码
            win32api.keybd_event(86, 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)
            time.sleep(1)
            win32api.keybd_event(13, 0, 0, 0)
            win32api.keybd_event(13, 0, win32con.KEYEVENTF_KEYUP, 0)
            time.sleep(1)
            pyperclip.copy('36.3')  # 此处存放要发送的字符串
            pyperclip.paste()
            win32api.keybd_event(17, 0, 0, 0)
            win32api.keybd_event(86, 0, 0, 0)
            win32api.keybd_event(86, 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)
            time.sleep(1)
            # 使用Enter发送内容
            #win32api.keybd_event(17, 0, 0, 0)
            win32api.keybd_event(13, 0, 0, 0)
            win32api.keybd_event(13, 0, win32con.KEYEVENTF_KEYUP, 0)
            #win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)

            # 使用Esc关闭后台
            win32api.keybd_event(27, 0, 0, 0)
            win32api.keybd_event(27, 0, win32con.KEYEVENTF_KEYUP, 0)

            sys.exit(0)  # 退出程序