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
    """动态获取资源路径"""
    try:
        base_path = sys._MEIPASS  # PyInstaller创建的临时目录
    except AttributeError:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

# 等待并点击保存按钮
def wait_and_click_save(image_path, timeout=400, interval=10, max_retries=40):
    retry_count = 0
    while retry_count < max_retries:
        start_time = time.time()
        found = False
        
        # 单次等待循环
        while time.time() - start_time < timeout:
            try:
                # 尝试定位图片（启用OpenCV加速）
                location = pyautogui.locateOnScreen(
                    image_path, 
                    grayscale=True, 
                    confidence=0.8
                )
                if location:
                    center = pyautogui.center(location)
                    pyautogui.click(center)
                    logging.info(f"成功点击 {image_path} 位置: {center}")
                    found = True
                    break
                else:
                    time.sleep(interval)  # 未找到则等待
            except pyautogui.ImageNotFoundException:
                time.sleep(interval)  # 捕获异常后继续等待
        
        if found:
            return True
        else:
            retry_count += 1
            logging.warning(f"第 {retry_count} 次重试，等待 {image_path} 出现...")
    
    logging.error(f"错误: 在 {max_retries} 次重试后仍未找到 {image_path}")
    return False

# 单个文件的打印流程
def autoprint(file_path):
    try:
        logging.info(f"开始处理文件: {os.path.basename(file_path)}")
        
        # 关键操作步骤
        pyautogui.moveTo(10, 10)
        pyautogui.leftClick(101, 0)
        time.sleep(5)
        
        # 使用快捷键组合
        pyautogui.hotkey('alt', 'R')
        time.sleep(5)

        pyautogui.hotkey('enter')
        time.sleep(15)

        pyautogui.hotkey('ctrl', 'p')
        time.sleep(15)

        pyautogui.hotkey('enter')
        time.sleep(10)

        # 等待并点击保存按钮
        save_img = get_resource_path('Save.PNG')
        if not wait_and_click_save(save_img, timeout=400, max_retries=40):
            return False
            
        logging.info(f"文件 {os.path.basename(file_path)} 打印成功")
        return True
        
    except Exception as e:
        logging.exception(f"处理文件 {os.path.basename(file_path)} 时发生错误")
        return False

# 批量处理rpf文件
def batch_process_rpf_files(folder_path):
    # 获取所有rpf文件
    rpf_files = glob.glob(os.path.join(folder_path, '*.rpf'))
    if not rpf_files:
        logging.error("未找到任何.rpf文件")
        return []
    
    results = []
    
    for file_path in rpf_files:
        file_name = os.path.basename(file_path)
        
        # 双击打开文件
        try:
            os.startfile(file_path)
            time.sleep(10)  # 等待文件打开
        except Exception as e:
            logging.error(f"无法打开文件 {file_name}: {str(e)}")
            results.append((file_name, "失败", f"打开文件失败: {str(e)}"))
            continue
        
        # 执行打印操作
        start_time = time.time()
        success = autoprint(file_path)
        elapsed_time = time.time() - start_time
        
        # 关闭当前文件
        try:
            pyautogui.hotkey('alt', 'f4')
            time.sleep(5)  # 等待关闭
        except Exception as e:
            logging.warning(f"关闭文件 {file_name} 时遇到问题: {str(e)}")
        
        # 记录结果
        status = "成功" if success else "失败"
        results.append((file_name, status, f"处理时间: {elapsed_time:.2f}秒"))
    
    return results

# 生成打印报告
def generate_print_report(results, output_dir):
    report_file = os.path.join(output_dir, f"print_report_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
    
    with open(report_file, 'w', encoding='utf-8') as f:
        f.write("RPT文件打印报告\n")
        f.write("=" * 50 + "\n")
        f.write(f"生成时间: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"处理文件总数: {len(results)}\n\n")
        
        success_count = sum(1 for r in results if r[1] == "成功")
        failure_count = len(results) - success_count
        
        f.write(f"打印成功: {success_count} 个文件\n")
        f.write(f"打印失败: {failure_count} 个文件\n")
        f.write("=" * 50 + "\n\n")
        
        f.write("文件处理详情:\n")
        f.write("-" * 50 + "\n")
        for file_name, status, remark in results:
            f.write(f"文件名: {file_name}\n")
            f.write(f"状态: {status}\n")
            f.write(f"备注: {remark}\n")
            f.write("-" * 50 + "\n")
    
    return report_file

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