import pyautogui
import cv2
import pytesseract
import time
import sys
import os
import logging
from PIL import Image
import numpy as np

# 配置日志
logging.basicConfig(filename='error.log', level=logging.INFO, format='%(asctime)s - %(levelname)s: %(message)s')

def get_resource_path(relative_path):
    """ 动态获取资源路径 """
    try:
        base_path = sys._MEIPASS  # PyInstaller创建的临时目录
    except AttributeError:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

def preprocess_image(image_path):
    img = Image.open(image_path)
    img = img.convert("L")  # 转灰度图
    img_np = np.array(img)
    # 锐化处理（增强边缘）
    kernel = np.array([[0, -1, 0], [-1, 5, -1], [0, -1, 0]])
    sharpened = cv2.filter2D(img_np, -1, kernel)
    # 降噪
    denoised = cv2.fastNlMeansDenoising(sharpened, None, 10, 7, 21)
    return denoised

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
        # 获取资源路径
        report_img = get_resource_path('report.png')
        cr_img = get_resource_path('CreateReport.png')
        ok_img = get_resource_path('OK.png')
        save_img = get_resource_path('Save.PNG')
        
        # 关键操作步骤（添加日志标记）
        pyautogui.moveTo(10, 10)
        pyautogui.leftClick(101, 0)
        time.sleep(5)
        logging.info("完成初始点击")
        
        # 查找report.png
        try:
            #report_img = get_resource_path('report.png')
            processed_report = preprocess_image(report_img)  # 预处理图片
            
            # 使用优化后的参数匹配
            location_report = pyautogui.locateOnScreen(
                processed_report,
                confidence=0.6,
                grayscale=True,
            )
            if not location_report:
                raise Exception("report.png未找到（请检查屏幕缩放或分辨率）")
            
        except Exception as e:
            error_screenshot = get_resource_path("error_screen.png")
            pyautogui.screenshot(error_screenshot)
            logging.error(f"报告按钮识别失败 | 错误: {str(e)} | 截图: {error_screenshot}")
            pyautogui.alert(f"报告按钮识别失败，截图已保存到程序目录")
            return  # 终止执行
        
        report_pos = pyautogui.center(location_report)
        pyautogui.leftClick(report_pos)
        time.sleep(5)


        location_CR = pyautogui.locateOnScreen(cr_img)
        if location_CR is not None:
            CR_pos = pyautogui.center(location_CR)
            print(f"found the button at:{location_CR}")
        else:
            print("No button")
        pyautogui.leftClick(CR_pos)
        time.sleep(15)

        pyautogui.hotkey('ctrl','p')
        time.sleep(15)

        OK_loc = pyautogui.locateCenterOnScreen(ok_img)
        pyautogui.leftClick(OK_loc)

        pyautogui.sleep(10)

        wait_and_click_save(save_img, timeout=400, max_retries=40)
            
        # 其他操作...
        
    except Exception as e:
        logging.exception("程序崩溃")
        pyautogui.alert(f"错误: {str(e)}")

if __name__ == "__main__":
    autoprint()