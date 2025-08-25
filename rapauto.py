import pyautogui  
import cv2  
import pytesseract  
import time

def wait_and_click_save(image_path, timeout=400, interval=10, max_retries=40):
    """
    等待指定图片出现并点击，支持超时和重试机制
    :param image_path: 图片路径（如 "SAVE.PNG"）
    :param timeout: 单次等待超时时间（秒）
    :param interval: 每次检测间隔（秒）
    :param max_retries: 最大重试次数
    """
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
    #screenshot = pyautogui.screenshot()
    #screenshot.save("current_screen.png")
    

    
    


    pyautogui.moveTo(10,10)
    pyautogui.leftClick(101,0)

    time.sleep(5)

    location_report = pyautogui.locateOnScreen('report.png')
    if location_report is not None:
        report_pos = pyautogui.center(location_report)
        print(f"found the button at:{location_report}")
        print(report_pos)
    else:
        print("No button")
    
    
    pyautogui.leftClick(report_pos)
    time.sleep(5)


    location_CR = pyautogui.locateOnScreen('CreateReport.png')
    if location_CR is not None:
        CR_pos = pyautogui.center(location_CR)
        print(f"found the button at:{location_CR}")
    else:
        print("No button")
    pyautogui.leftClick(CR_pos)
    time.sleep(15)

    pyautogui.hotkey('ctrl','p')
    time.sleep(15)

    OK_loc = pyautogui.locateCenterOnScreen('OK.png')
    pyautogui.leftClick(OK_loc)

    pyautogui.sleep(10)

    wait_and_click_save("Save.PNG", timeout=400, max_retries=40)


autoprint()