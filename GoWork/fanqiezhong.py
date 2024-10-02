import tkinter as tk
from tkinter import messagebox, simpledialog, font,ttk
import pandas as pd
import time
import threading
import tempfile
import os
import threading
from pynput import keyboard, mouse
import psutil
from win10toast import ToastNotifier
import random
from datetime import datetime
import sounddevice as sd
import soundfile as sf
import pygetwindow as gw
import subprocess
import sys



class PomodoroApp:
    def __init__(self, root):



        # 时间配置
        self.pomodoro_time = 1500  # 番茄钟工作时间:25分钟
        self.break_time = 300  # 番茄钟休息时间5分钟
        self.long_break_time = 300  # 完成一项任务后的休息时间5分钟
        self.changxiuxicishu = 0 ##初始化长休息次数
        self.changxiuxishijian = 1800 ##长休息时间:半小时
        self.baogaoshengcheng = 300 ##生成报告的时间：5分钟
        self.fadaishijian = 120 ##发呆两分钟自动警告

        self.root = root
        self.root.title("革命需要你！")
        ##设置图标
        self.root.geometry("300x465")
        self.root.configure(bg="#e0dad2")  # 你可以将 "lightblue" 替换为你喜欢的颜色


        icon_path = "GoWork/icons/shiqin2.ico"  # 替换为你的图标文件路径
        if os.path.exists(icon_path):
            self.root.iconbitmap(icon_path)


        #self.custom_font = font.Font(root,family="SLZLH J Thin", size=14)##拾陆字濑户
        self.custom_font = font.Font(root,family="Hangeuljaemin4.0", size=17)##韩国字体
        #self.custom_font = font.Font(root,family="XYGZJSB J", size=12)##新愚公装甲宋
        #self.custom_font = font.Font(root,family="汇文明朝体", size=15)
        #self.custom_font = font.Font(root,family="京華老宋体", size=17)
        #self.custom_font = font.Font(root,family="汀明体", size=16, weight="bold")
        #self.custom_font = font.Font(root,family="长城长宋体", size=17)
        #self.custom_font = font.Font(root,family="霞鹜铭心宋", size=17,weight="bold")






        # 文件路径
        self.file_path = "GoWork/tasks.xlsx"
        self.todo_df = pd.read_excel(self.file_path, sheet_name='今日待办')
        self.completed_df = pd.read_excel(self.file_path, sheet_name='今日完成')
        self.ico_p = 'GoWork/icons/shiqin2.ico'
        self.ico_p1 = 'GoWork/icons/111~1.ico'
        ##番茄钟开始的配音文件
        self.peyin_fanqiestart = 'GoWork/peiyin/3.wav'
        self.peyin_fanqieend = 'GoWork/peiyin/4.wav'

        # 初始化变量
        self.current_task = None
        self.pomodoro_count = 0
        self.pomodoro_count_task = 0
        self.flowers = 0
        self.interruptions = 0
        self.pause_time = 0
        self.pause = False  # 暂停状态
        self.remaining_time = 0  # 剩余时间
        self.state = "idle"  # 当前状态：idle, working, break, paused
        self.running = False  # 控制线程运行的标志


        # 创建界面组件
        self.create_widgets()
        # 加载任务
        self.load_task()

        # 初始化Windows通知
        self.toaster = ToastNotifier()

        ##初始化插入
        self.last_key_press_time = time.time()
        self.df_path = 'GoWork/Tixing.xlsx'
        self.sheet = 'huayu'
        self.df = pd.read_excel(self.df_path, sheet_name=self.sheet)
        self.logph = 'GoWork/log.xlsx'
        self.logs = 'Sheet1'
        self.log = pd.read_excel(self.logph, sheet_name=self.logs)
        self.bmd = ['Drama','BC项目组','心诗交响曲','留声机','bc_art']





    def create_widgets(self):
        # 创建当前任务标签
        self.task_label = tk.Label(self.root, text="当前任务: ", bg="#e0dad2", font=self.custom_font, fg="#4a372b")
        self.task_label.pack()

        # 创建任务内容标签
        self.task_content = tk.Label(self.root, text="", bg="#e0dad2", font=self.custom_font, fg="#4a372b")
        self.task_content.pack()

        # 创建进度标签
        self.progress_label = tk.Label(self.root, text="进度: ", bg="#e0dad2", font=self.custom_font, fg="#4a372b")
        self.progress_label.pack()

        # 创建进度条画布
        self.progress_bar = tk.Canvas(self.root, width=200, height=20, bg="#e0dad2",highlightbackground="#4a372b", highlightthickness=0.5)
        self.progress_bar.pack()

        # 创建剩余时间标签
        self.time_label = tk.Label(self.root, text="剩余时间: ", bg="#e0dad2", font=self.custom_font, fg="#4a372b")
        self.time_label.pack()

        button_frame = tk.Frame(self.root, bg="#e0dad2")
        button_frame.pack(pady=2)
        # 创建开始按钮
        self.start_button = tk.Button(button_frame, text="开始", bg="#e0dad2", command=self.start_pomodoro, font=self.custom_font, fg="#4a372b")
        self.start_button.grid(row=0, column=0, padx=5, pady=0)

        # 创建暂停按钮
        self.pause_button = tk.Button(button_frame, text="暂停", bg="#e0dad2", command=self.pause_pomodoro, font=self.custom_font, fg="#4a372b")
        self.pause_button.grid(row=0, column=1, padx=5, pady=0)

        # 创建继续按钮
        self.continue_button = tk.Button(button_frame, text="继续", bg="#e0dad2", command=self.continue_pomodoro, font=self.custom_font, fg="#4a372b")
        self.continue_button.grid(row=0, column=2, padx=5, pady=0)

        button_frame2 = tk.Frame(self.root, bg="#e0dad2")
        button_frame2.pack(pady=2)
        # 创建添加待办事项按钮
        self.add_task_button = tk.Button(button_frame2,text="添加待办", bg="#e0dad2", command=self.add_task, font=self.custom_font, fg="#4a372b")
        self.add_task_button.grid(row=0, column=0, padx=5, pady=0)
        # 创建显示待办事项按钮
        self.show_todo_button = tk.Button(button_frame2,text="显示待办", bg="#e0dad2", command=self.show_todo, font=self.custom_font, fg="#4a372b")
        self.show_todo_button.grid(row=0, column=1, padx=5, pady=0)

        # 创建事件完成按钮
        self.complete_button = tk.Button(self.root, text="事件完成", bg="#e0dad2", command=self.complete_task, font=self.custom_font, fg="#4a372b")
        self.complete_button.pack(pady=2)

        # 创建小红花标签
        self.flowers_label = tk.Label(self.root, text="小红花: 0", bg="#e0dad2", font=self.custom_font, fg="#4a372b")
        self.flowers_label.pack()

        # 创建番茄钟计数标签
        self.pomodoro_count_label = tk.Label(self.root, text="当前事件耗费番茄: 0", bg="#e0dad2", font=self.custom_font, fg="#4a372b")
        self.pomodoro_count_label.pack()

        # 创建打断次数标签
        self.interruptions_label = tk.Label(self.root, text="当前番茄被打断次数: 0", bg="#e0dad2", font=self.custom_font, fg="#4a372b")
        self.interruptions_label.pack()

        # 创建暂停时间标签
        self.pause_time_label = tk.Label(self.root, text="暂停时间: 0", bg="#e0dad2", font=self.custom_font, fg="#4a372b")
        self.pause_time_label.pack()

        # 创建长休息次数标签
        self.changxiuxi_label = tk.Label(self.root, text="差4次长休息", bg="#e0dad2", font=self.custom_font, fg="#4a372b")
        self.changxiuxi_label.pack()



    def load_task(self):
        # 如果今日待办sheet不为空
        if not self.todo_df.empty:
            # 找到优先度最小的任务
            min_priority = self.todo_df['优先度'].min()
            self.current_task = self.todo_df[self.todo_df['优先度'] == min_priority].iloc[0]
            # 更新任务内容标签
            self.task_content.config(text=self.current_task['任务'])

    def start_pomodoro(self):
        # 如果有当前任务
        if self.current_task is not None:
            if self.state == "idle":
                self.pause = False
                self.remaining_time = self.pomodoro_time  # 重置剩余时间
                self.state = "working"  # 设置状态为工作中
                self.running = True  # 设置线程运行标志为True
                self.toaster.show_toast(
                    title=f"已开始关于任务{self.current_task['任务']}的新的番茄钟！",
                    msg=f"林欣欣：番茄钟开始！同志，让我们为了最美好的前途战斗吧！",
                    icon_path=self.ico_p1,  # 可选，通知图标的路径
                    duration=10,  # 通知显示的持续时间（以秒为单位）
                    threaded=True,  # 是否在单独的线程中显示通知（默认为True）
                    #callback_on_click=on_click  # 当用户点击通知时调用的回调函数（可选）
                )
                data, fs = sf.read(self.peyin_fanqiestart)
                sd.play(data, fs)
                sd.wait()  # 等待声音播放完毕
                sd.stop() ##关闭声音
                self.pomodoro_thread = threading.Thread(target=self.run_pomodoro)
                self.pomodoro_thread.start()
                # 启动键盘监控线程
                if not hasattr(self, 'keyboard_thread') or not self.keyboard_thread.is_alive():
                    self.keyboard_thread = threading.Thread(target=self.start_keyboard_listener)
                    self.keyboard_thread.daemon = True
                    self.keyboard_thread.start()

                # 启动鼠标监控线程
                if not hasattr(self, 'mouse_thread') or not self.mouse_thread.is_alive():
                    self.mouse_thread = threading.Thread(target=self.start_mouse_listener)
                    self.mouse_thread.daemon = True
                    self.mouse_thread.start()

                # 启动软件监控线程
                if not hasattr(self, 'software_thread') or not self.software_thread.is_alive():
                    self.software_thread = threading.Thread(target=self.monitor_software_usage, args=(self.log, self.bmd))
                    self.software_thread.daemon = True
                    self.software_thread.start()

    def run_pomodoro(self):
        while self.running:
            if self.state == "working":
                self.working_pomodoro()
            elif self.state == "break":
                self.break_pomodoro()
            elif self.state == "paused":
                self.paused_pomodoro()
            time.sleep(1)

    def working_pomodoro(self):
        if not self.pause:
            self.update_progress(self.remaining_time, self.pomodoro_time)
            self.update_time(self.remaining_time)
            self.remaining_time -= 1
            if self.remaining_time <= 0:
                self.pomodoro_count += 1
                self.pomodoro_count_task += 1
                self.flowers += 1
                self.update_labels()
                self.state = "break"
                self.toaster.show_toast(
                    title=f"番茄钟结束",
                    msg=f"林欣欣：番茄钟结束~同志，可以放松一下啦~",
                    icon_path=self.ico_p1,  # 可选，通知图标的路径
                    duration=7,  # 通知显示的持续时间（以秒为单位）
                    threaded=True,  # 是否在单独的线程中显示通知（默认为True）
                    #callback_on_click=on_click  # 当用户点击通知时调用的回调函数（可选）
                )
                data, fs = sf.read(self.peyin_fanqieend)
                sd.play(data, fs)
                sd.wait()  # 等待声音播放完毕
                sd.stop() ##关闭声音
                self.changxiuxicishu = self.changxiuxicishu +1
                if self.changxiuxicishu != 4:
                    self.remaining_time = self.break_time
                    self.shifouchangxiuxi = 0
                else:
                    self.remaining_time = self.changxiuxishijian
                    self.changxiuxicishu = 0
                    self.shifouchangxiuxi = 1
                self.update_changxiuxi(self.changxiuxicishu)


    def break_pomodoro(self):
        if self.shifouchangxiuxi == 0:
            self.update_progress(self.remaining_time, self.break_time)
        else:
            self.update_progress(self.remaining_time, self.changxiuxishijian)
        self.update_time(self.remaining_time)
        self.remaining_time -= 1
        if self.remaining_time <= 0:
            self.toaster.show_toast(
                title=f"已开始关于任务{self.current_task['任务']}的新的番茄钟！",
                msg=f"林欣欣：番茄钟开始！同志，让我们为了最美好的前途战斗吧！",
                icon_path=self.ico_p1,  # 可选，通知图标的路径
                duration=10,  # 通知显示的持续时间（以秒为单位）
                threaded=True,  # 是否在单独的线程中显示通知（默认为True）
                #callback_on_click=on_click  # 当用户点击通知时调用的回调函数（可选）
            )
            data, fs = sf.read(self.peyin_fanqiestart)
            sd.play(data, fs)
            sd.wait()  # 等待声音播放完毕
            sd.stop() ##关闭声音
            self.load_task()
            self.state = "working"
            self.remaining_time = self.pomodoro_time

    def paused_pomodoro(self):
        self.pause_time += 1
        self.update_labels()
        if self.pause_time > 5 * 60:
            self.flowers -= 1
            self.pause_time = 0
            messagebox.showinfo("提示", "番茄钟重置，扣除一朵小红花")

    def pause_pomodoro(self):
        if self.state == "working":
            self.pause = True  # 暂停开关打开
            self.interruptions += 1
            self.state = "paused"  # 设置状态为暂停

    def continue_pomodoro(self):
        if self.pause == True:
            self.pause = False  # 暂停开关关闭
            self.state = "working"  # 设置状态为工作中


    def add_task(self):
        # 获取任务内容、优先度和预计番茄数量
        content = simpledialog.askstring("添加待办事项", "内容:")
        priority = simpledialog.askinteger("添加待办事项", "紧急程度:")
        pomodoros = simpledialog.askinteger("添加待办事项", "预计番茄数量:")

        if content and priority is not None and pomodoros is not None:
            # 创建新任务DataFrame
            new_task = pd.DataFrame({'任务': [content], '优先度': [priority], '预计消耗番茄数量': [pomodoros]})
            self.todo_df = pd.concat([self.todo_df, new_task], ignore_index=True)
            self.save_excel()

    def complete_task(self):
        if self.current_task is not None:
            ###创建通知
            self.toaster.show_toast(
                title=f"任务{self.current_task['任务']}已完成",
                msg=f"当前任务已完成，用了{self.pomodoro_count_task}个番茄钟",
                icon_path=self.ico_p,  # 可选，通知图标的路径
                duration=10,  # 通知显示的持续时间（以秒为单位）
                threaded=True,  # 是否在单独的线程中显示通知（默认为True）
                #callback_on_click=on_click  # 当用户点击通知时调用的回调函数（可选）
            )
            self.current_task['实际消耗番茄数量'] = self.pomodoro_count_task
            self.pomodoro_count_task = 0
            self.current_task['打断次数'] = self.interruptions
            self.interruptions = 0
            self.completed_df = pd.concat([self.completed_df, self.current_task.to_frame().T], ignore_index=True)
            self.todo_df = self.todo_df[self.todo_df['任务'] != self.current_task['任务']]
            self.save_excel()
            self.load_task()


            if self.current_task is not None:
                self.running = False  # 停止当前线程
                self.pomodoro_thread.join()  # 等待线程结束
                self.run_long_break()  # 开始长休息
            else:
                self.running = False  # 停止当前线程
                self.pomodoro_thread.join()  # 等待线程结束


    def run_long_break(self):
        self.state = "break"
        self.toaster.show_toast(
            title=f"番茄钟结束",
            msg=f"林欣欣：番茄钟结束~同志，可以放松一下啦~",
            icon_path=self.ico_p1,  # 可选，通知图标的路径
            duration=10,  # 通知显示的持续时间（以秒为单位）
            threaded=True,  # 是否在单独的线程中显示通知（默认为True）
            #callback_on_click=on_click  # 当用户点击通知时调用的回调函数（可选）
        )
        data, fs = sf.read(self.peyin_fanqieend)
        sd.play(data, fs)
        sd.wait()  # 等待声音播放完毕
        sd.stop() ##关闭声音
        self.remaining_time = self.long_break_time
        self.running = True
        self.pomodoro_thread = threading.Thread(target=self.run_pomodoro)
        self.pomodoro_thread.start()

    def save_excel(self):
        # 创建一个临时文件
        #temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        #temp_file_path = temp_file.name
        # 保存今日待办和今日完成sheet到临时文件
        with pd.ExcelWriter(self.file_path) as writer:
            self.todo_df.to_excel(writer, sheet_name='今日待办', index=False)
            self.completed_df.to_excel(writer, sheet_name='今日完成', index=False)

    def update_progress(self, remaining_time, total_time):
        # 计算进度并更新进度条
        progress = (total_time - remaining_time) / total_time * 200
        self.progress_bar.delete("all")
        self.progress_bar.create_rectangle(0, 0, progress, 20, fill="#4a372b")

    def update_time(self, remaining_time):
        # 计算剩余时间的分钟和秒数，并更新剩余时间标签
        minutes, seconds = divmod(remaining_time, 60)
        self.time_label.config(text=f"剩余时间: {minutes:02}:{seconds:02}")

    def update_changxiuxi(self, changxiuxitime):
        # 长休息时间
        changxiuxi_cha = 4 - changxiuxitime
        self.changxiuxi_label.config(text=f"差{changxiuxi_cha}次长休息")

    def update_labels(self):
        # 更新小红花、番茄钟计数、打断次数和暂停时间标签
        self.flowers_label.config(text=f"小红花: {self.flowers}")
        self.pomodoro_count_label.config(text=f"今日完成番茄: {self.pomodoro_count}")
        self.interruptions_label.config(text=f"当前番茄被打断次数: {self.interruptions}")
        self.pause_time_label.config(text=f"暂停时间: {self.pause_time // 60}:{self.pause_time % 60:02}")


    def start_keyboard_listener(self):
        with keyboard.Listener(on_press=self.on_press) as listener:
            listener.join()

    def start_mouse_listener(self):
        with mouse.Listener(on_click=self.on_click) as listener:
            listener.join()

    def on_press(self, key):
        self.last_key_press_time = time.time()

    def on_click(self, x, y, button, pressed):
        if pressed:
            self.last_key_press_time = time.time()

    def monitor_software_usage(self, logsheet, bmd):
        wufenzhongzongji = 0
        wufenzhonganjian = 0
        xiedaishijian = 0
        drama_time = 0
        wendang_time = 0
        ps_time = 0
        music_time = 0
        l2d_time = 0
        qitatime = 0

        while True:
            if self.state == "working":
                active_window = gw.getActiveWindow().title
                print(f"Active Window: {active_window}")
                dangqiangongzuochuangkou = '无用窗口'
                qitatime = qitatime + 1
                for item in bmd:
                    if item in active_window:
                        print(f"当前窗口包含: {item}")
                        dangqiangongzuochuangkou = '有用窗口'
                        chuangkoujishi = item
                        if item == 'Drama':
                            drama_time = drama_time + 1
                            qitatime = qitatime - 1
                        if item == 'BC项目组':
                            wendang_time = wendang_time + 1
                            qitatime = qitatime - 1
                        if item == '心诗交响曲':
                            ps_time = ps_time + 1
                            qitatime = qitatime - 1
                        if item == '留声机':
                            music_time = music_time + 1
                            qitatime = qitatime - 1
                        if item == 'bc_art':
                            l2d_time = l2d_time + 1
                            qitatime = qitatime - 1

                wufenzhongzongji = wufenzhongzongji + 1
                if wufenzhongzongji == self.baogaoshengcheng:
                    dangqiandazilv = (wufenzhonganjian/wufenzhongzongji)*100

                    current_datetime = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    drama_timelv = (drama_time/wufenzhongzongji)*100
                    wendang_timelv = (wendang_time/wufenzhongzongji)*100
                    ps_timelv = (ps_time/wufenzhongzongji)*100
                    music_timelv = (music_time/wufenzhongzongji)*100
                    l2d_timelv = (l2d_time/wufenzhongzongji)*100
                    qitatimelv = (qitatime/wufenzhongzongji)*100
                    logsheet = logsheet.append({'时间': current_datetime, '工作率': f"{dangqiandazilv}%",'其他占用时间': f"{qitatimelv}%",'live2d占用时间': f"{l2d_timelv}%",'音乐占用时间': f"{music_timelv}%",'ps占用时间': f"{ps_timelv}%",'文档占用时间': f"{wendang_timelv}%",'Drama占用时间':f"{drama_timelv}%"}, ignore_index=True)
                    wufenzhongzongji = 0
                    wufenzhonganjian = 0
                    drama_time = 0
                    wendang_time = 0
                    ps_time = 0
                    music_time = 0
                    l2d_time = 0
                    qitatime = 0
                    self.toaster.show_toast(
                        title=f"5分钟报告,有用窗口的工作时间占到了{dangqiandazilv}%，",
                        msg=f"已经记录在表格中，可以进行查看”",
                        icon_path="GoWork/icons/shiqin2.ico",
                        duration=300,
                        threaded=True,
                    )
                    data, fs = sf.read("GoWork/baogao.wav")
                    sd.play(data, fs)
                    sd.wait()
                    logsheet.to_excel("GoWork/log.xlsx", index=False)

                current_time = time.time()
                if dangqiangongzuochuangkou == '有用窗口':
                    if current_time - self.last_key_press_time < 1:
                        wufenzhonganjian += 1
                        xiedaishijian = 0
                    else:
                        xiedaishijian = xiedaishijian + 1
                else:
                    xiedaishijian = xiedaishijian + 1

                if xiedaishijian > self.fadaishijian:
                    xiedaishijian = xiedaishijian + 1
                    random_integer = random.randint(0, 1)
                    character = self.df.at[random_integer, '角色']
                    doing = self.df.at[random_integer, '动作']
                    talk = self.df.at[random_integer, '话语']
                    ico_p = self.df.at[random_integer, '图片位置']
                    peyin = self.df.at[random_integer, '配音']
                    self.toaster.show_toast(
                        title="同志同志，你在干什么呀~",
                        msg=f"{character}{doing}:“{talk}”",
                        icon_path=ico_p,
                        duration=5,
                        threaded=True,
                    )
                    data, fs = sf.read(peyin)
                    sd.play(data, fs)
                    sd.wait()

                time.sleep(1)
            else:
                time.sleep(1)
    def show_todo(self):
        # 创建新的小界面
        todo_window = tk.Toplevel(self.root)
        todo_window.title("待办事项")
        todo_window.geometry("1200x600")

        # 创建Treeview组件
        tree = ttk.Treeview(todo_window)
        tree.pack(fill=tk.BOTH, expand=True)

        # 设置表格列
        tree["columns"] = tuple(self.todo_df.columns)
        for col in self.todo_df.columns:
            tree.heading(col, text=col)
            tree.column(col, anchor="center")

        # 插入数据
        for index, row in self.todo_df.iterrows():
            tree.insert("", "end", values=tuple(row))

        # 自定义Treeview的样式
        style = ttk.Style()
        style.configure("Treeview",
                        background="#e0dad2",
                        foreground="#4a372b",
                        fieldbackground="#e0dad2",
                        font=self.custom_font)
        style.map('Treeview', background=[('selected', '#4a372b')])
        
        # 配置未被选择行的样式
        style.configure("Treeview.unselected",
                        background="#e0dad2",
                        foreground="#4a372b",
                        font=self.custom_font)
        tree.tag_configure('unselected', background="#e0dad2", foreground="#4a372b")

        # 配置表头的样式
        style.configure("Treeview.Heading",
                        background="#4a372b",
                        foreground="#4a372b",
                        font=self.custom_font)
        # 创建添加待办事项按钮
        add_task_button = tk.Button(todo_window, text="添加待办事项", bg="#e0dad2", command=self.add_task, font=self.custom_font, fg="#4a372b")
        add_task_button.pack()


if __name__ == "__main__":
    root = tk.Tk()
    app = PomodoroApp(root)
    root.mainloop()





