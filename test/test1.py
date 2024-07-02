import win32com.client
import pythoncom
import os
import pyautogui
from playsound import playsound
import datetime
import random
import time
from pywinauto import Application, Desktop
from PIL import ImageChops, ImageGrab
import psutil

# 播放随机音频文件
def play_random_audio(folder_path):
    audio_files = [f for f in os.listdir(folder_path) if f.endswith('.wav')]
    if audio_files:
        random_file = random.choice(audio_files)
        playsound(os.path.join(folder_path, random_file))

def is_red_exclamation_mark(image):
    # 检查图像中是否有特定的RGB颜色 (204, 74, 49)
    target_color = (204, 74, 49)
    for pixel in image.getdata():
        if pixel[:3] == target_color:
            return True
    return False

# 通过进程名得到进程号pid
def get_pid_by_name(process_name):
    for process in psutil.process_iter(['pid', 'name']):
        if process.info['name'] == process_name:
            return process.info['pid']
    return None

teams_pid = get_pid_by_name('ms-teams.exe')
if teams_pid:
    app = Application(backend="uia").connect(process=teams_pid)
    teams_window = app.window(title_re=".*Microsoft Teams.*")
    # 得到需要监控的菜单按钮
    chat_btn = teams_window.child_window(title="Chat", control_type="Button")
    activity_btn = teams_window.child_window(title="Activity", control_type="Button")
    teams_btn = teams_window.child_window(title="Teams", control_type="Button")
    calendar_btn = teams_window.child_window(title="Calendar", control_type="Button")
    buttons = [chat_btn, activity_btn, teams_btn, calendar_btn]

    while True:
        try:
            for btn in buttons:
                # 截图按钮区域
                btn_image = btn.capture_as_image()
                # 检查按钮是否有红色色块
                if is_red_exclamation_mark(btn_image):
                    print(f"New message in {btn.window_text()}!")
                    play_random_audio("../audio/new_teams_msg")
        except Exception as e:
            # 获取任务栏上的Teams按钮
            desktop = Desktop(backend="uia")
            teams_taskbar = desktop.window(class_name="Shell_TrayWnd")
            btn = teams_taskbar.child_window(title_re=".*Microsoft Teams.*", control_type="Button")
            btn_image = btn.capture_as_image()
            # 检查按钮是否有红色色块
            if is_red_exclamation_mark(btn_image):
                print(f"New message in {btn.window_text()}!")
                play_random_audio("../audio/new_teams_msg")
        time.sleep(5)
else:
    print("Teams not found.")

# 进入消息循环以便事件处理
pythoncom.PumpMessages()


