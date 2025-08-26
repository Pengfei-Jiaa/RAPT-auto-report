import pyautogui
import time
import sys
import os
import logging
import glob
import datetime
import tkinter as tk
from tkinter import filedialog

# 配置日志系统
def setup_logging(output_dir):
    log_file = os.path.join(output_dir, f"print_log_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
    logging.basicConfig(
        filename=log_file,
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    logging.getLogger().addHandler(console_handler)

# 获取资源路径（支持PyInstaller打包）
def get_resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

# 新增功能1：处理"无法获取更新"对话框 [6,9](@ref)
def handle_update_exception():
    try:
        # 尝试定位并点击OK按钮
        ok_img = get_resource_path('OK.PNG')  # 需准备OK按钮截图
        if wait_and_click(ok_img, timeout=10, max_retries=3):
            logging.info("成功关闭更新异常对话框")
            return True
        # 若图像识别失败，改用模拟回车键 [7](@ref)
        pyautogui.press('enter')
        logging.info("通过回车键关闭更新对话框")
        return True
    except Exception as e:
        logging.error(f"处理更新对话框失败: {str(e)}")
        return False

# 新增功能2：处理关闭时的保存提示 [9](@ref)
def handle_save_prompt():
    try:
        # 定位并点击"No"按钮
        no_img = get_resource_path('No.PNG')  # 需准备No按钮截图
        if wait_and_click(no_img, timeout=10, max_retries=3):
            logging.info("已跳过文件保存")
            return True
        # 备用方案：键盘导航至"No"按钮 [8](@ref)
        pyautogui.press('left')  # 移动焦点至"No"
        pyautogui.press('enter') # 确认选择
        return True
    except Exception as e:
        logging.error(f"处理保存提示失败: {str(e)}")
        return False

# 通用图像点击函数（增强版）[1,9](@ref)
def wait_and_click(image_path, timeout=40, interval=2, max_retries=5):
    for _ in range(max_retries):
        try:
            location = pyautogui.locateOnScreen(image_path, confidence=0.85, grayscale=True)
            if location:
                center = pyautogui.center(location)
                pyautogui.click(center)
                logging.info(f"点击成功: {image_path}")
                return True
            time.sleep(interval)
        except pyautogui.ImageNotFoundException:
            time.sleep(interval)
    logging.warning(f"未找到图像: {image_path}")
    return False

# 单个文件处理流程（新增异常监控）
def autoprint(file_path):
    try:
        logging.info(f"开始处理: {os.path.basename(file_path)}")
        # 原有操作流程
        pyautogui.hotkey('alt', 'R')
        time.sleep(5)
        pyautogui.press('enter')
        time.sleep(15)
        pyautogui.hotkey('ctrl', 'p')
        time.sleep(15)
        pyautogui.press('enter')
        time.sleep(10)
        
        # 新增：保存对话框监控
        save_img = get_resource_path('Save.PNG')
        if not wait_and_click(save_img):
            return False

        logging.info(f"文件打印成功")
        return True
        
    except Exception as e:
        logging.exception(f"处理异常: {str(e)}")
        return False

# 批量处理流程（核心增强）
def batch_process_rpf_files(folder_path):
    rpf_files = glob.glob(os.path.join(folder_path, '*.rpf'))
    results = []
    for file_path in rpf_files:
        file_name = os.path.basename(file_path)
        try:
            os.startfile(file_path)
            time.sleep(10)
            
            # 新增：持续监控异常对话框 [3](@ref)
            start_time = time.time()
            while time.time() - start_time < 600:  # 10分钟超时
                if handle_update_exception():  # 阻塞直到对话框消失
                    time.sleep(2)
                else:
                    break
            
            success = autoprint(file_path)
        except Exception as e:
            success = False
            logging.error(f"文件打开失败: {file_name} - {str(e)}")
        
        # 新增：关闭文件并处理保存提示
        try:
            pyautogui.hotkey('alt', 'f4')
            time.sleep(3)
            handle_save_prompt()  # 关键新增步骤 [9](@ref)
        except Exception as e:
            logging.warning(f"关闭异常: {file_name} - {str(e)}")
        
        status = "成功" if success else "失败"
        results.append((file_name, status))
    return results

# 主函数
def main():
    # 创建Tkinter根窗口（不显示）
    root = tk.Tk()
    root.withdraw()
    
    # 让用户选择文件夹
    folder_path = filedialog.askdirectory(title="选择包含rpf文件的文件夹")
    if not folder_path:
        logging.error("未选择文件夹，程序退出")
        return
    
    # 设置日志
    setup_logging(folder_path)
    logging.info(f"开始处理文件夹: {folder_path}")
    
    # 批量处理文件
    results = batch_process_rpf_files(folder_path)
    
    # 生成报告
    if results:
        report_file = generate_print_report(results, folder_path)
        logging.info(f"处理完成！报告已保存至: {report_file}")
        
        # 在资源管理器中显示报告文件
        try:
            os.startfile(report_file)
        except:
            logging.info(f"请手动查看报告: {report_file}")

if __name__ == "__main__":
    main()






    