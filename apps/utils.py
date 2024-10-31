import os
import random
import subprocess
import sys
from playsound import playsound

# 播放随机音频文件
def play_random_audio(folder_path):
    audio_files = [f for f in os.listdir(folder_path) if f.endswith('.wav')]
    if audio_files:
        random_file = random.choice(audio_files)
        playsound(os.path.join(folder_path, random_file))

# 重启程序
def restart_program():
    print("Restarting program...")
    subprocess.Popen([sys.executable] + sys.argv)
    print("Restarted")