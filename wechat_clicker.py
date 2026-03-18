import pyautogui
import time
import logging
from PIL import Image
import io
import numpy as np
import subprocess
import win32gui
import win32con
import os
import sys

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def capture_screen_region(x, y, width, height):
    """
    捕获屏幕指定区域
    
    参数:
    x, y: 区域左上角坐标
    width, height: 区域宽高
    
    返回:
    PIL.Image对象
    """
    screenshot = pyautogui.screenshot(region=(x, y, width, height))
    return screenshot

def is_wechat_running():
    """
    检查微信进程是否正在运行
    
    返回:
    bool - 如果微信进程正在运行则返回True，否则返回False
    """
    try:
        # 检查常见的微信进程名
        wechat_process_names = ['Weixin.exe', 'WeChatApp.exe']
        for process_name in wechat_process_names:
            result = subprocess.run(['tasklist', '/FI', f'IMAGENAME eq {process_name}'], 
                                  capture_output=True, text=True)
            if process_name in result.stdout:
                logging.info(f"发现微信进程: {process_name}")
                return True
        return False
    except Exception as e:
        logging.error(f"检查微信进程时出错: {e}")
        return False

def bring_wechat_to_foreground():
    """
    将微信窗口置顶显示，增强版
    
    返回:
    bool - 如果成功置顶则返回True，否则返回False
    """
    try:
        # 微信窗口的标题和类名模式
        wechat_window_titles = ['微信', 'WeChat', 'Weixin']
        wechat_window_classes = ['WeChatMainWndForPC', 'WeChat']
        found_window = False
        
        def window_callback(hwnd, ctx):
            """窗口枚举回调函数"""
            nonlocal found_window
            try:
                # 获取窗口标题和类名
                title = win32gui.GetWindowText(hwnd)
                class_name = win32gui.GetClassName(hwnd)
                
                # 检查是否匹配微信窗口标题或类名
                is_wechat_window = False
                
                # 检查标题
                for wechat_title in wechat_window_titles:
                    if wechat_title in title:
                        is_wechat_window = True
                        break
                
                # 如果标题不匹配，检查类名
                if not is_wechat_window:
                    for wechat_class in wechat_window_classes:
                        if wechat_class in class_name:
                            is_wechat_window = True
                            break
                
                if is_wechat_window:
                    logging.info(f"找到微信窗口: 标题='{title}', 类名='{class_name}', 窗口句柄={hwnd}")
                    
                    # 检查窗口状态，如果是最小化则先恢复
                    window_rect = win32gui.GetWindowPlacement(hwnd)
                    if window_rect[1] == win32con.SW_SHOWMINIMIZED:
                        logging.info("微信窗口处于最小化状态，尝试恢复...")
                        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                        # time.sleep(0.5)  # 给窗口时间恢复
                    
                    # 尝试更可靠的窗口前置方法
                    try:
                        # 多次尝试设置为前台
                        for _ in range(3):
                            # 先激活窗口
                            try:
                                win32gui.SetActiveWindow(hwnd)
                            except:
                                pass
                            
                            # 再设置为前台
                            try:
                                result = win32gui.SetForegroundWindow(hwnd)
                                if result:
                                    break
                            except:
                                pass
                            # time.sleep(0.1)
                        
                        # 确保窗口可见
                        win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
                        
                        # 使用pyautogui点击窗口区域确保焦点
                        rect = win32gui.GetWindowRect(hwnd)
                        if rect[2] > rect[0] and rect[3] > rect[1]:
                            center_x = (rect[0] + rect[2]) // 2
                            center_y = (rect[1] + rect[3]) // 2
                            # 点击窗口的标题栏区域，更可靠
                            title_bar_y = rect[1] + 30  # 假设标题栏高度约30像素
                            pyautogui.click(center_x, title_bar_y)
                            # time.sleep(0.2)
                        
                        logging.info("微信窗口已成功置顶并激活")
                        found_window = True
                        return False  # 停止枚举
                    except Exception as e:
                        logging.warning(f"设置窗口为前台时出错: {e}")
                        # 回退方案：使用pyautogui直接点击窗口区域
                        rect = win32gui.GetWindowRect(hwnd)
                        if rect[2] > rect[0] and rect[3] > rect[1]:
                            center_x = (rect[0] + rect[2]) // 2
                            center_y = (rect[1] + rect[3]) // 2
                            logging.info(f"尝试通过点击激活窗口: ({center_x}, {center_y})")
                            pyautogui.click(center_x, center_y)
                            # time.sleep(0.5)
                            found_window = True
                            return False
            except Exception as e:
                logging.warning(f"处理窗口时出错: {e}")
            return True
        
        # 枚举所有顶级窗口
        win32gui.EnumWindows(window_callback, None)
        
        # 如果没找到，尝试通过窗口标题直接查找
        if not found_window:
            logging.info("尝试通过窗口标题直接查找微信窗口...")
            try:
                # 直接查找包含微信关键词的窗口
                def find_wechat_window(hwnd, ctx):
                    nonlocal found_window
                    try:
                        title = win32gui.GetWindowText(hwnd)
                        class_name = win32gui.GetClassName(hwnd)
                        
                        # 检查是否是可见窗口且标题包含微信关键词
                        if win32gui.IsWindowVisible(hwnd):
                            if any(keyword in title for keyword in ['微信', 'WeChat', 'Weixin']):
                                logging.info(f"找到微信窗口: 标题='{title}', 类名='{class_name}', 窗口句柄={hwnd}")
                                
                                # 恢复窗口并置顶
                                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                                win32gui.SetForegroundWindow(hwnd)
                                found_window = True
                                return False
                    except:
                        pass
                    return True
                
                win32gui.EnumWindows(find_wechat_window, None)
            except Exception as e:
                logging.warning(f"通过窗口标题查找微信窗口时出错: {e}")
        
        # 额外的回退方案：直接查找并点击任务栏中的微信图标
        if not found_window:
            logging.info("尝试通过任务栏点击微信...")
            try:
                # 获取任务栏高度（通常在底部）
                screen_width, screen_height = pyautogui.size()
                taskbar_height = 50  # 估计的任务栏高度
                
                # 扫描任务栏区域寻找可能的微信图标
                for x in range(0, screen_width, 20):
                    screenshot = pyautogui.screenshot(region=(x, screen_height - taskbar_height, 40, taskbar_height))
                    # 简单检查是否有绿色像素（可能的微信图标特征）
                    np_img = np.array(screenshot)
                    green_pixels = np.sum((np_img[:, :, 1] > 150) & (np_img[:, :, 0] < np_img[:, :, 1] * 0.8) & (np_img[:, :, 2] < np_img[:, :, 1] * 0.8))
                    if green_pixels > 50:  # 如果找到足够的绿色像素
                        pyautogui.click(x + 20, screen_height - taskbar_height // 2)
                        # time.sleep(1)
                        found_window = True
                        logging.info("通过任务栏点击了可能的微信图标")
                        break
            except Exception as e:
                logging.warning(f"尝试通过任务栏点击微信时出错: {e}")
        
        if found_window:
            logging.info("微信窗口已成功置顶")
            return True
        else:
            logging.warning("未找到微信窗口")
            return False
    except Exception as e:
        logging.error(f"置顶微信窗口时出错: {e}")
        return False

def is_green_background_with_text(image, text_keyword="进入微信", threshold=0.8):
    """
    检查图像是否为绿底白字的"进入微信"按钮
    
    参数:
    image: PIL.Image - 要分析的图像
    text_keyword: str - 关键词
    threshold: float - 绿色阈值
    
    返回:
    bool - 是否匹配
    """
    try:
        # 转换为RGB模式
        img_rgb = image.convert('RGB')
        # 获取图像数据
        np_img = np.array(img_rgb)
        
        # 提取RGB通道
        r, g, b = np_img[:, :, 0], np_img[:, :, 1], np_img[:, :, 2]
        
        # 统计总像素数
        total_pixels = np_img.shape[0] * np_img.shape[1]
        
        # 定义绿色背景：G值明显高于R和B，且G值较大
        green_area = (g > 150) & (r < g * 0.8) & (b < g * 0.8)
        
        # 计算绿色区域占比
        green_ratio = np.sum(green_area) / total_pixels
        
        # 绿色区域占比需要超过阈值
        if green_ratio < threshold:
            return False
        
        # 检查是否有足够的白色像素（文字部分）
        white_pixels = (r > 200) & (g > 200) & (b > 200)
        white_ratio = np.sum(white_pixels) / total_pixels
        
        # 白色像素占比应该在合理范围内（文字占比）
        if not (0.05 < white_ratio < 0.5):
            return False
        
        return True
    except Exception as e:
        logging.error(f"图像分析出错: {e}")
        return False

def find_and_click_green_wechat_button(timeout=30):
    """
    查找并点击绿底白字的"进入微信"按钮
    
    参数:
    timeout: int - 超时时间（秒）
    
    返回:
    bool - 是否成功点击
    """
    start_time = time.time()
    screen_width, screen_height = pyautogui.size()
    
    logging.info("开始查找绿底白字的'进入微信'按钮...")
    
    while time.time() - start_time < timeout:
        try:
            # 获取整个屏幕截图
            full_screenshot = pyautogui.screenshot()
            
            # 扫描步长，平衡速度和精度
            scan_step = 20
            
            # 假设按钮大小至少为40x20像素
            min_button_width, min_button_height = 80, 40
            
            # 遍历屏幕查找匹配区域
            for x in range(0, screen_width - min_button_width, scan_step):
                for y in range(0, screen_height - min_button_height, scan_step):
                    # 提取一个区域进行分析
                    region = full_screenshot.crop((x, y, x + min_button_width, y + min_button_height))
                    
                    # 检查是否为绿底白字的按钮
                    if is_green_background_with_text(region):
                        # 计算中心点
                        center_x = x + min_button_width // 2
                        center_y = y + min_button_height // 2
                        
                        logging.info(f"找到'进入微信'按钮，位置: ({center_x}, {center_y})")
                        
                        # 移动到目标位置并点击
                        pyautogui.moveTo(center_x, center_y, duration=0.2)
                        pyautogui.click()
                        logging.info("已点击'进入微信'按钮")
                        return True
            
            logging.info("未找到匹配区域，2秒后重试...")
            # time.sleep(2)
            
        except Exception as e:
            logging.error(f"查找过程中出错: {e}")
            # time.sleep(2)
    
    logging.warning(f"超时：{timeout}秒内未找到'进入微信'按钮")
    return False

def main():
    """
    主函数：检测微信进程，将窗口置顶，然后查找并点击绿底白字的'进入微信'按钮
    """
    print("微信进入按钮自动点击器启动")
    print("检测微信进程并确保窗口在前台...")
    print("正在寻找绿底白字的'进入微信'按钮...")
    print("按 Ctrl+C 可以随时退出")
    
    try:
        # 检查微信进程是否正在运行
        if is_wechat_running():
            print("✓ 微信进程正在运行")
            print("尝试将微信窗口置顶...")
            
            # 尝试多次将微信窗口置顶，提高成功率
            max_attempts = 3
            attempts = 0
            success_foreground = False
            
            while attempts < max_attempts and not success_foreground:
                attempts += 1
                print(f"窗口置顶尝试 {attempts}/{max_attempts}...")
                if bring_wechat_to_foreground():
                    success_foreground = True
                    print("✓ 微信窗口已成功置顶")
                    # 窗口置顶后，可以尝试点击屏幕上一个合理的位置以确保焦点转移
                    screen_width, screen_height = pyautogui.size()
                    pyautogui.click(screen_width // 2, screen_height // 2)
                    # time.sleep(0.5)
                    break
                else:
                    print("窗口置顶失败，1秒后重试...")
                    # time.sleep(1)
            
            if not success_foreground:
                print("⚠ 无法将微信窗口置顶，尝试其他方法...")
                # 尝试使用Alt+Tab切换窗口作为备用方案
                print("尝试使用Alt+Tab切换窗口...")
                pyautogui.keyDown('alt')
                for _ in range(5):  # 切换几个窗口
                    pyautogui.press('tab')
                    # time.sleep(0.2)
                pyautogui.keyUp('alt')
                # time.sleep(1)
            
            # 再次尝试置顶，确保窗口在前台
            if not success_foreground:
                bring_wechat_to_foreground()
            
            # 短暂等待，确保窗口有足够时间显示在前台
            print("等待窗口稳定显示...")
            # time.sleep(2)
        else:
            print("⚠ 未检测到微信进程运行")
        
        # 查找并点击按钮
        print("开始查找并点击'进入微信'按钮...")
        success = find_and_click_green_wechat_button(timeout=60)
        
        if success:
            print("🎉 成功点击'进入微信'按钮！")
            # 点击后再次检查并置顶微信窗口，确保操作后窗口可见
            print("再次确认微信窗口在前台...")
            if is_wechat_running():
                if bring_wechat_to_foreground():
                    print("✓ 操作完成，微信窗口保持在前台")
        else:
            print("❌ 未找到'进入微信'按钮，请确保按钮在屏幕上可见")
    
    except KeyboardInterrupt:
        print("\n程序已被用户中断")
    except ImportError as e:
        print(f"缺少必要的库: {e}")
        print("请安装所需库: pip install pyautogui Pillow numpy pywin32")
    except Exception as e:
        print(f"程序异常: {e}")
    finally:
        print("程序已退出")

if __name__ == "__main__":
    main()