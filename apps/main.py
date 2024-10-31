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
import claim_reminder  # 导入提醒模块
from utils import play_random_audio, restart_program  # 导入通用函数

print("Message Monitor is running...")

# -------------------------------------outlook start--------------------------------------
class OutlookHandler:
    def OnNewMailEx(self, receivedItemsIDs):
        current_time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
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

def get_pid_by_name(process_name):
    for process in psutil.process_iter(['pid', 'name']):
        if process.info['name'] == process_name:
            return process.info['pid']
    return None

teams_pid = get_pid_by_name('ms-teams.exe')
teams_flag = False
buttons = []
if teams_pid:
    app = Application(backend="uia").connect(process=teams_pid)
    teams_window = app.window(title_re=".*Microsoft Teams.*")
    buttons = [
        teams_window.child_window(title="Chat", control_type="Button"),
        teams_window.child_window(title="Activity", control_type="Button"),
        teams_window.child_window(title="Teams", control_type="Button"),
        teams_window.child_window(title="Calendar", control_type="Button")
    ]
    teams_flag = True
else:
    print("Teams not found.")
    restart_program()

def teams_handler(buttons):
    try:
        for btn in buttons:
            btn_image = btn.capture_as_image()
            if is_red_exclamation_mark(btn_image):
                current_time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
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
            current_time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            print(f"{current_time}: New message in {btn.window_text()}!")
            pyautogui.press('esc')
            play_random_audio("../audio/new_teams_msg")

# -------------------------------------Teams end----------------------------------------

try:
    claim_reminder.schedule_reminders()
    while True:
        if teams_flag:
            teams_handler(buttons)
        # 启动提醒调度
        claim_reminder.run_reminders()
        time.sleep(10)
except Exception as e:
    print(f"An error occurred in the Teams section: {e}")
    restart_program()

pythoncom.PumpMessages()