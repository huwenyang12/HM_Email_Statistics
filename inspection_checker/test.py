import pyautogui
import time
import keyboard  # 需先 pip install keyboard

time.sleep(5)

print("开始自动滚动，按Esc停止。")
while True:
    if keyboard.is_pressed('esc'):
        print("检测到Esc，停止滚动。")
        break
    pyautogui.scroll(-300)
    time.sleep(2)
