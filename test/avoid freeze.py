import ctypes
import pyautogui
import time
import sys
import os
import psutil
import logging
import glob
import datetime
import tkinter as tk
import win32gui
import win32con
from tkinter import filedialog

# 定义常量
ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001  # 阻止系统休眠
ES_DISPLAY_REQUIRED = 0x00000002  # 阻止锁屏/关闭显示器

# 启动防休眠（阻止休眠+锁屏）
ctypes.windll.kernel32.SetThreadExecutionState(
    ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED
)

import pyautogui

import time

def activate_rapt_window():
    try:
        # 修正点1：精确匹配标题+类名（示例类名需替换为实际值）
        hwnd = win32gui.FindWindow("RAPT_Class", "RAPT")  
        if not hwnd:
            logging.warning("窗口未找到，尝试模糊匹配标题...")
            hwnd = find_window_by_substring("RAPT")  # 需实现模糊查找函数
        
        if hwnd:
            # 修正点2：显式显示窗口（应对隐藏状态）
            win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            
            # 修正点3：绕过系统焦点限制
            win32gui.BringWindowToTop(hwnd)
            win32gui.SetFocus(hwnd)
            logging.info("窗口激活成功")
            return True
    except Exception as e:
        logging.error(f"激活失败: {e}")
    return False

# 辅助函数：通过标题关键词模糊查找句柄
def find_window_by_substring(substring):
    hwnd_list = []
    def callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd) and substring in win32gui.GetWindowText(hwnd):
            hwnd_list.append(hwnd)
        return True
    win32gui.EnumWindows(callback, None)
    return hwnd_list[0] if hwnd_list else None


activate_rapt_window()



def prevent_lock_screen():


    while True:


        pyautogui.moveRel(1, 0)  # 鼠标向右移动1像素


        time.sleep(10)  # 每10秒钟移动一次


        pyautogui.moveRel(-1, 0)  # 鼠标向左移动1像素


        time.sleep(10)
