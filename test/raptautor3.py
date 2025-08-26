import pyautogui

import time
import sys
import os
import logging

def get_resource_path(relative_path):
    """ 动态获取资源路径 """
    try:
        base_path = sys._MEIPASS  # PyInstaller创建的临时目录
    except AttributeError:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


def wait_and_click_save(image_path, timeout=400, interval=10, max_retries=40):
    retry_count = 0
    while retry_count < max_retries:
        start_time = time.time()
        found = False
        
        # 单次等待循环
        while time.time() - start_time < timeout:
            try:
                # 尝试定位图片（启用 OpenCV 加速，设置置信度阈值）
                location = pyautogui.locateOnScreen(
                    image_path, 
                    grayscale=True, 
                    confidence=0.8  # 需安装 OpenCV 才能使用此参数 [11](@ref)
                )
                if location:
                    center = pyautogui.center(location)
                    pyautogui.click(center)
                    print(f"成功点击 {image_path} 位置: {center}")
                    found = True
                    break
                else:
                    time.sleep(interval)  # 未找到则等待
            except pyautogui.ImageNotFoundException:
                time.sleep(interval)  # 捕获异常后继续等待 [10,11](@ref)
        
        if found:
            return True
        else:
            retry_count += 1
            print(f"第 {retry_count} 次重试，等待 {image_path} 出现...")
    
    print(f"错误: 在 {max_retries} 次重试后仍未找到 {image_path}")
    return False

def autoprint():
    try:

        
        # 关键操作步骤（添加日志标记）
        pyautogui.moveTo(10, 10)
        pyautogui.leftClick(101, 0)
        time.sleep(5)
        logging.info("完成初始点击")
        
        try:
            pyautogui.hotkey('alt','R')
            
        except Exception as e:
            return  # 终止执行
        
        time.sleep(5)

        pyautogui.hotkey('enter')

        time.sleep(15)

        pyautogui.hotkey('ctrl','p')
        time.sleep(15)

        pyautogui.hotkey('enter')

        pyautogui.sleep(10)

        
        save_img = get_resource_path('Save.PNG')

        wait_and_click_save(save_img, timeout=400, max_retries=40)
            
        # 其他操作...
        
    except Exception as e:
        logging.exception("程序崩溃")
        pyautogui.alert(f"错误: {str(e)}")

if __name__ == "__main__":
    autoprint()