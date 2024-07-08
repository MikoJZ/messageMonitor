import win32com.client
import pythoncom
import pyautogui
import os
import sys
import random
import time
from pywinauto import Application, Desktop
from playsound import playsound
import datetime
import psutil
import subprocess

current_time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

def play_audio(file_path):
    playsound(file_path)

# 播放随机音频文件
def play_random_audio(folder_path):
    audio_files = [f for f in os.listdir(folder_path) if f.endswith('.wav')]
    if audio_files:
        random_file = random.choice(audio_files)
        playsound(os.path.join(folder_path, random_file))

# -------------------------------------outlook start--------------------------------------
class OutlookHandler:
    def OnNewMailEx(self, receivedItemsIDs):
        print(f"{current_time}: Received a new email！")
        # 模拟按下ESC键 唤醒计算机
        pyautogui.press('esc')
        # 播放语音
        play_random_audio("../audio/new_email")

outlook = win32com.client.DispatchWithEvents("Outlook.Application", OutlookHandler)
# -------------------------------------outlook end--------------------------------------

# -------------------------------------Teams start--------------------------------------
def is_red_exclamation_mark(image):
    return any(pixel[0] > 195 and pixel[1] < 100 and pixel[2] < 100 for pixel in image.getdata())

def restart_program():
    # 重启程序
    print("Restarting program...")
    subprocess.Popen([sys.executable] + sys.argv)

def get_pid_by_name(process_name):
    for process in psutil.process_iter(['pid', 'name']):
        if process.info['name'] == process_name:
            return process.info['pid']
    return None

teams_pid = get_pid_by_name('ms-teams.exe')
if teams_pid:
    app = Application(backend="uia").connect(process=teams_pid)
    teams_window = app.window(title_re=".*Microsoft Teams.*")
    buttons = [
        teams_window.child_window(title="Chat", control_type="Button"),
        teams_window.child_window(title="Activity", control_type="Button"),
        teams_window.child_window(title="Teams", control_type="Button"),
        teams_window.child_window(title="Calendar", control_type="Button")
    ]

    try:
        while True:
            try:
                for btn in buttons:
                    btn_image = btn.capture_as_image()
                    if is_red_exclamation_mark(btn_image):
                        print(f"{current_time}: New message in {btn.window_text()}!")
                        pyautogui.press('esc')
                        play_random_audio("../audio/new_teams_msg")
                        break
            except Exception as e:
                desktop = Desktop(backend="uia")
                teams_taskbar = desktop.window(class_name="Shell_TrayWnd")
                btn = teams_taskbar.child_window(title_re=".*Microsoft Teams.*", control_type="Button")
                btn_image = btn.capture_as_image()
                if is_red_exclamation_mark(btn_image):
                    print(f"{current_time}: New message in {btn.window_text()}!")
                    pyautogui.press('esc')
                    play_random_audio("../audio/new_teams_msg")
            time.sleep(10)
    except Exception as e:
        print(f"An error occurred in the Teams section: {e}")
        restart_program()
else:
    print("Teams not found.")
    restart_program()
# -------------------------------------Teams end----------------------------------------
pythoncom.PumpMessages()