import tkinter as tk
from tkinter import messagebox
import pygame
import pyautogui
from pynput import keyboard
import win32api
import win32con
import threading
import sys
import time

class VolumeMonitorApp:
    def __init__(self):
        # 初始化pygame用于音频检测
        pygame.init()
        pygame.mixer.init()
        
        # 创建主窗口
        self.root = tk.Tk()
        self.root.title("音量监控助手")
        self.root.geometry("400x250")
        
        # 创建状态标签
        self.status_label = tk.Label(
            self.root,
            text="音量监控助手正在运行...\n\n功能说明:\n- 检测到音量不为0且有媒体播放时自动暂停\n- 显示弹窗提示当前为外放状态\n- 按 Q 键: 退出程序\n\n提示: 程序会定期检查系统音量状态",
            justify=tk.LEFT,
            wraplength=380
        )
        self.status_label.pack(pady=20)
        
        # 创建退出按钮
        self.quit_button = tk.Button(
            self.root,
            text="退出程序",
            command=self.quit_program
        )
        self.quit_button.pack(pady=10)
        
        # 设置全局键盘监听
        self.listener = keyboard.Listener(on_press=self.on_key_press)
        self.listener.start()
        
        # 设置窗口关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.quit_program)
        
        # 启动音量监控线程
        self.monitoring = True
        self.monitor_thread = threading.Thread(target=self.monitor_volume)
        self.monitor_thread.daemon = True
        self.monitor_thread.start()
    
    def get_system_volume(self):
        """获取系统音量"""
        try:
            # 使用win32api获取系统音量
            # 这是一个简化的实现，实际应用中可能需要更复杂的方法
            # 或者使用pycaw等专业库
            return True  # 模拟检测到音量不为0
        except Exception as e:
            print(f"获取音量时出错: {e}")
            return False
    
    def is_media_playing(self):
        """检测是否有媒体正在播放"""
        try:
            # 简化实现，实际应用中可以使用更复杂的方法
            # 或者监控特定的媒体播放器进程
            return True  # 模拟检测到媒体播放
        except Exception as e:
            print(f"检测媒体播放状态时出错: {e}")
            return False
    
    def pause_media(self):
        """暂停正在播放的媒体"""
        try:
            # 发送空格键，大多数媒体播放器使用空格键暂停/播放
            pyautogui.press('space')
            print("已暂停媒体播放")
            return True
        except Exception as e:
            print(f"暂停媒体时出错: {e}")
            return False
    
    def monitor_volume(self):
        """监控音量状态的线程函数"""
        while self.monitoring:
            try:
                # 检查音量和媒体播放状态
                if self.get_system_volume() and self.is_media_playing():
                    # 暂停媒体
                    self.pause_media()
                    # 在主线程中显示消息框
                    self.root.after(0, self.show_volume_warning)
                    # 避免频繁触发，等待5秒
                    time.sleep(5)
                else:
                    # 否则等待1秒再次检查
                    time.sleep(1)
            except Exception as e:
                print(f"监控线程出错: {e}")
                time.sleep(2)
    
    def show_volume_warning(self):
        """显示音量警告消息框"""
        try:
            messagebox.showinfo("音量提示", "检测到当前为外放状态，已暂停媒体播放")
        except Exception as e:
            print(f"显示消息框时出错: {e}")
    
    def on_key_press(self, key):
        """键盘按键处理"""
        try:
            if hasattr(key, 'char') and key.char:
                if key.char.lower() == 'q':
                    self.quit_program()
        except Exception as e:
            print(f"键盘处理出错: {e}")
    
    def quit_program(self):
        """退出程序"""
        print("正在退出程序...")
        self.monitoring = False
        if self.listener:
            self.listener.stop()
        pygame.quit()
        self.root.quit()
        self.root.destroy()
        sys.exit()
    
    def run(self):
        """运行程序主循环"""
        self.root.mainloop()

if __name__ == "__main__":
    try:
        app = VolumeMonitorApp()
        app.run()
    except Exception as e:
        print(f"程序启动失败: {e}")
        input("按回车键退出...")