import time
import ctypes

# 等待1分钟（60秒）
time.sleep(60)

# 在1分钟后执行的代码
print("1 minute has passed!")
# SC_MONITORPOWER = 0xF170
# MONITOR_ON = -1
ctypes.windll.user32.SendMessageW(0xFFFF, 0x0112, 0xF170, -1)

