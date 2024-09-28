import time
import threading
from pynput import keyboard
import psutil
from win10toast import ToastNotifier
import pandas as pd
import random
from datetime import datetime
import sounddevice as sd
import soundfile as sf
import pygetwindow as gw


# 初始化Windows通知
toaster = ToastNotifier()


##初始化表格导入

df_path = 'GoWork/Tixing.xlsx'
sheet = 'huayu'

df = pd.read_excel(df_path, sheet_name=sheet)


logph = 'GoWork/log.xlsx'
logs = 'Sheet1'


log = pd.read_excel(logph, sheet_name=logs)




##初始化计时器

#每5分钟弹窗一次的五分钟分析函数

fenxihanshu_chushihua = time.time()

##上一秒有没有按键盘的时间戳

shangyimiao_anjian = time.time()

##初始化五分钟内每秒钟一次记录的时间









# 记录键盘活动的时间戳
last_key_press_time = time.time()

# 监控键盘事件
def on_press(key):
    global last_key_press_time
    last_key_press_time = time.time()

def start_keyboard_listener():
    with keyboard.Listener(on_press=on_press) as listener:
        listener.join()

# 监控软件使用情况
def monitor_software_usage(logsheet):
    wufenzhongzongji = 0
    wufenzhonganjian = 0


    while True:
        # 获取当前活动的窗口标题
        active_window = gw.getActiveWindow().title
        print(f"Active Window: {active_window}")
        ##加一个5分检查点
        wufenzhongzongji = wufenzhongzongji + 1
        #判断是否是五分钟检查点：
        if wufenzhongzongji == 300:##300秒，就是五分钟
            dangqiandazilv = (wufenzhonganjian/wufenzhongzongji)*100
            toaster.show_toast(
                title=f"5分钟报告,打字时间{dangqiandazilv}%",
                msg=f"已经记录在表格中，可以进行查看”",
                icon_path="GoWork/icons/shiqin2.ico",  # 可选，通知图标的路径
                duration=10,  # 通知显示的持续时间（以秒为单位）
                threaded=True,  # 是否在单独的线程中显示通知（默认为True）
                #callback_on_click=on_click  # 当用户点击通知时调用的回调函数（可选）
                )
            data, fs = sf.read("GoWork/baogao.wav")
            sd.play(data, fs)
            sd.wait()  # 等待声音播放完毕
            current_datetime = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            logsheet = logsheet.append({'时间': current_datetime, '按键率': f"{dangqiandazilv}%"}, ignore_index=True)
            wufenzhongzongji = 0
            wufenzhonganjian = 0
            logsheet.to_excel("GoWork/log.xlsx", index=False)


        # 检查键盘活动
        current_time = time.time()
        ###开始记录上一秒中有没有
        if current_time - last_key_press_time <1:
            ##按键的一秒钟以内
            wufenzhonganjian +=1
        else:
            if current_time - last_key_press_time > 60:  # 如果一分钟内没有键盘活动
                ##创建一个提醒随机数
                random_integer = random.randint(0, 1)
                character = df.at[random_integer, '角色']
                doing = df.at[random_integer, '动作']
                talk = df.at[random_integer, '话语']
                ico_p = df.at[random_integer, '图片位置']
                toaster.show_toast(
                    title="同志同志，你在干什么呀~",
                    msg=f"{character}{doing}:“{talk}”",
                    icon_path=ico_p,  # 可选，通知图标的路径
                    duration=5,  # 通知显示的持续时间（以秒为单位）
                    threaded=True,  # 是否在单独的线程中显示通知（默认为True）
                    #callback_on_click=on_click  # 当用户点击通知时调用的回调函数（可选）
                )
                data, fs = sf.read("GoWork/gowork.wav")
                sd.play(data, fs)
                sd.wait()  # 等待声音播放完毕

        # 每半秒钟检查一次
        time.sleep(1)

# 启动键盘监控线程
keyboard_thread = threading.Thread(target=start_keyboard_listener)
keyboard_thread.daemon = True
keyboard_thread.start()

# 启动软件监控线程
software_thread = threading.Thread(target=monitor_software_usage(log))
software_thread.daemon = True
software_thread.start()

# 主线程保持运行
try:
    while True:
        time.sleep(1)
except KeyboardInterrupt:
    print("程序已停止")
