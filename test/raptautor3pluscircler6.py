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

def close_edge_process():
    """关闭所有Edge进程并返回关闭的进程数量"""
    closed_count = 0
    for proc in psutil.process_iter():
        try:
            # 检查进程名和命令行参数[4](@ref)
            if "msedge" in proc.name().lower() or "MicrosoftEdge" in proc.name():
                # 检查命令行参数确认是PDF查看器[7](@ref)
                cmdline = " ".join(proc.cmdline()).lower()
                if "--type=utility" in cmdline and "pdf" in cmdline:
                    proc.kill()
                    logging.info(f"已关闭Edge PDF查看器进程: PID={proc.pid}")
                    closed_count += 1
        except (psutil.NoSuchProcess, psutil.AccessDenied) as e:
            # 处理权限问题[6](@ref)
            logging.warning(f"访问进程{proc.pid}被拒绝: {str(e)}")
        except Exception as e:
            logging.error(f"关闭Edge进程时出错: {str(e)}")
    return closed_count

def activate_rapt_window():
    """激活RAPT主窗口"""
    try:
        hwnd = win32gui.FindWindow(None, "RAPT")
        if hwnd:
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd)
            logging.info("已激活RAPT窗口")
            return True
    except Exception as e:
        logging.error(f"激活RAPT窗口失败: {str(e)}")
    return False

def close_edge_and_activate_rapt():
    """关闭Edge并激活RAPT窗口的组合操作"""
    # 先关闭Edge进程
    closed_count = close_edge_process()
    
    # 如果关闭了Edge进程，等待更长时间
    if closed_count > 0:
        time.sleep(3)  # 等待Edge完全关闭
    
    # 尝试激活RAPT窗口
    if not activate_rapt_window():
        # 备用方案：通过Alt+Tab切换窗口
        pyautogui.hotkey('alt', 'tab')
        time.sleep(1)

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


# 获取资源路径
def get_resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

# 增强的图像查找函数（带超时和重试限制）
def wait_and_click(image_path, timeout=10, interval=1, max_retries=3, click=True):
    """
    查找图像并点击（如果找到）
    
    参数:
    image_path - 图像文件路径
    timeout - 每次尝试的最大等待时间
    interval - 检查间隔
    max_retries - 最大重试次数
    click - 是否执行点击操作
    
    返回:
    找到返回True，否则返回False
    """
    retry_count = 0
    while retry_count < max_retries:
        try:
            # 尝试定位图像（使用OpenCV加速）
            location = pyautogui.locateOnScreen(
                image_path, 
                grayscale=True,
                confidence=0.75  # 降低置信度要求以提高兼容性
            )
            if location:
                if click:
                    center = pyautogui.center(location)
                    pyautogui.click(center)
                    logging.info(f"成功点击图像: {image_path}")
                else:
                    logging.info(f"成功找到图像: {image_path}")
                return True
        except pyautogui.ImageNotFoundException:
            pass  # 忽略未找到图像的异常[3](@ref)
        
        # 记录重试信息
        if retry_count < max_retries - 1:
            logging.debug(f"未找到图像: {image_path}，{interval}秒后重试...")
            time.sleep(interval)
        retry_count += 1
    
    logging.info(f"未找到图像: {image_path} (达到最大重试次数)")
    return False

# 智能处理更新异常对话框
def handle_update_exception():
    """
    处理'Cannot get update'对话框
    返回True如果找到并处理了对话框，否则False
    """
    ok_img = get_resource_path('OK.PNG')
    
    # 尝试通过图像识别关闭对话框
    if wait_and_click(ok_img, timeout=5, max_retries=2):
        logging.info("通过图像识别关闭更新对话框")
        return True
    
    # 备用方案：尝试按回车键
    logging.info("尝试通过回车键关闭对话框")
    pyautogui.press('enter')
    
    # 再次检查对话框是否消失
    if not wait_and_click(ok_img, timeout=2, max_retries=1, click=False):
        logging.info("更新对话框已关闭")
        return True
    
    logging.warning("无法确认更新对话框状态")
    return False

# 处理关闭时的保存提示
def handle_save_prompt():
    """
    处理关闭文件时的保存提示
    返回True如果找到并处理了提示，否则False
    """
    no_img = get_resource_path('No.PNG')
    
    # 尝试通过图像识别点击"No"
    if wait_and_click(no_img, timeout=25, interval=2, max_retries=5):
        logging.info("已跳过文件保存")
        return True
    
    # 备用方案：键盘导航
    logging.info("通过键盘选择不保存")
    pyautogui.press('right')  # 移动焦点至"No"
    time.sleep(0.5)
    pyautogui.press('enter')  # 确认选择
    return True

# 单个文件处理流程（增加状态控制）
def autoprint(file_path):
    """
    处理单个RPT文件的打印流程
    
    返回:
    success - 是否成功完成打印
    """
    
    try:
        file_name = os.path.basename(file_path)
        logging.info(f"开始处理文件: {file_name}")
        
        # 状态标记：是否已开始打印操作
        print_started = False
        
        # 关键操作步骤
        pyautogui.moveTo(10, 10)
        pyautogui.leftClick(101, 0)
        time.sleep(5)
        
        # 检查并处理更新对话框（仅尝试一次）
        #handle_update_exception()
    
        # 使用快捷键组合
        pyautogui.hotkey('alt', 'R')
        time.sleep(5)
        
        # 在关键操作点之间加入对话框检查
        pyautogui.press('enter')
        time.sleep(15)
        '''
        # 检查对话框（最多尝试2次）
        for _ in range(2):
            if not wait_and_click(get_resource_path('OK.PNG'), timeout=3, max_retries=1, click=False):
                break
            handle_update_exception()
        '''
        pyautogui.hotkey('ctrl', 'p')
        time.sleep(10)
        '''
        # 检查对话框（最多尝试2次）
        for _ in range(2):
            if not wait_and_click(get_resource_path('OK.PNG'), timeout=3, max_retries=1, click=False):
                break
            handle_update_exception()
        '''
        pyautogui.press('enter')
        time.sleep(10)
        
        # 标记已开始打印
        print_started = True
        
        # 等待并点击保存按钮
        save_img = get_resource_path('Save.PNG')
        if wait_and_click(save_img, timeout=4000, interval=5, max_retries=100):
            logging.info(f"文件 {file_name} 打印成功")

            time.sleep(5)
            
            close_edge_and_activate_rapt()

            return True
        
        logging.warning(f"文件 {file_name} 打印失败：未找到保存按钮")
        return False
        
        

    except Exception as e:
        logging.exception(f"处理文件时发生错误: {str(e)}")
        return False



    finally:
        # 确保文件关闭（无论成功与否）
        try:
            if print_started:
                # 确保焦点在RAPT窗口上
                close_edge_and_activate_rapt()
                time.sleep(2)
                #关闭当前RATP文件
                pyautogui.hotkey('alt', 'f4')
                time.sleep(2)
                #处理保存提示
                handle_save_prompt()
        except Exception as e:
            logging.warning(f"关闭文件时出错: {str(e)}")

# 批量处理RPT文件
def batch_process_rpf_files(folder_path):
    # 获取所有rpf文件
    rpf_files = glob.glob(os.path.join(folder_path, '*.rpf'))
    if not rpf_files:
        logging.error("未找到任何.rpf文件")
        return []
    
    results = []
    
    for file_path in rpf_files:
        file_name = os.path.basename(file_path)
        
        try:
            # 打开文件
            os.startfile(file_path)
            logging.info(f"已打开文件: {file_name}")
            time.sleep(10)  # 等待文件加载
            
            # 处理更新对话框（最多尝试2次）
            '''
            for _ in range(2):
                if not wait_and_click(get_resource_path('OK.PNG'), timeout=5, max_retries=1, cl
                i
                ck=False):
                    break
                handle_update_exception()
                time.sleep(2)
            '''
            # 执行打印操作
            success = autoprint(file_path)
            status = "成功" if success else "失败"
            results.append((file_name, status, ""))
            
        except Exception as e:
            logging.error(f"无法处理文件 {file_name}: {str(e)}")
            results.append((file_name, "失败", f"错误: {str(e)}"))
        finally:
            # 确保关闭文件
            try:
                pyautogui.hotkey('alt', 'f4')
                time.sleep(2)
                handle_save_prompt()
                time.sleep(3)
            except:
                pass
    
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
            if remark:
                f.write(f"备注: {remark}\n")
            f.write("-" * 50 + "\n")
    
    return report_file

# 主函数
def main():
    # 创建Tkinter根窗口
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
    # 设置pyautogui安全参数
    pyautogui.FAILSAFE = True  # 启用安全特性[6](@ref)
    pyautogui.PAUSE = 0.5      # 设置操作间隔[8](@ref)
    
    main()