import tkinter as tk
from pynput import keyboard
import pyautogui
import sys
import threading
import time
import win32gui
import win32process
import psutil

class DocumentSaverApp:
    def __init__(self):
        # 创建主窗口
        self.root = tk.Tk()
        self.root.title("文档自动保存助手")
        self.root.geometry("420x300")
        
        # 创建状态标签
        self.status_label = tk.Label(
            self.root,
            text="文档自动保存助手正在运行...\n\n可用快捷键:\n- 按 A 键: 保存当前活动文档\n- 按 Q 键: 退出程序\n\n功能说明:\n- 自动识别当前活动窗口的类型\n- 针对不同类型的应用程序使用相应的保存快捷键\n- 支持大多数常见的办公软件和编辑器\n- 操作结果会实时显示在状态栏",
            justify=tk.LEFT,
            wraplength=400
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
        
        # 记录最后一次操作时间，避免频繁操作
        self.last_save_time = 0
        self.cooldown_time = 0.5  # 冷却时间，单位秒
        
        # 应用程序特定的保存快捷键映射
        self.save_hotkeys = {
            # 通用保存快捷键
            'default': ['ctrl', 's'],
            # 可以根据需要添加特定应用程序的保存快捷键
        }
        
        # 已知的文档编辑应用程序列表
        self.document_apps = [
            'word', 'excel', 'powerpnt', 'outlook',  # Microsoft Office
            'notepad++', 'sublime_text', 'vscode', 'atom',  # 文本编辑器
            'photoshop', 'illustrator', 'indesign',  # Adobe Creative Suite
            'pycharm', 'intellij', 'eclipse', 'visualstudio',  # IDEs
            'chrome', 'firefox', 'edge', 'opera',  # 浏览器
        ]
    
    def is_cooldown_over(self):
        """检查操作冷却时间是否已过"""
        current_time = time.time()
        if current_time - self.last_save_time < self.cooldown_time:
            return False
        self.last_save_time = current_time
        return True
    
    def get_active_window_info(self):
        """获取当前活动窗口的信息"""
        try:
            # 获取当前活动窗口的句柄
            hwnd = win32gui.GetForegroundWindow()
            # 获取窗口标题
            window_title = win32gui.GetWindowText(hwnd)
            # 获取进程ID
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            # 获取进程名称
            process = psutil.Process(pid)
            process_name = process.name().lower()
            return window_title, process_name
        except Exception as e:
            print(f"获取活动窗口信息时出错: {e}")
            return "未知窗口", "unknown"
    
    def is_document_application(self, process_name):
        """检查是否为文档类应用程序"""
        for app in self.document_apps:
            if app in process_name:
                return True
        return False
    
    def save_current_document(self):
        """保存当前活动文档"""
        if not self.is_cooldown_over():
            return
        
        try:
            # 获取当前活动窗口信息
            window_title, process_name = self.get_active_window_info()
            
            # 检查是否为文档类应用程序
            is_document = self.is_document_application(process_name)
            
            # 选择适当的保存快捷键
            if process_name in self.save_hotkeys:
                hotkey = self.save_hotkeys[process_name]
            else:
                hotkey = self.save_hotkeys['default']  # 使用默认的Ctrl+S
            
            # 发送保存快捷键
            pyautogui.hotkey(*hotkey)
            
            # 构建状态消息
            if is_document:
                status_msg = f"已保存文档: {window_title}"
            else:
                status_msg = f"已尝试保存: {process_name}"
            
            print(status_msg)
            # 更新状态标签
            self.root.after(0, lambda: self.update_status(status_msg))
            
        except Exception as e:
            error_msg = f"保存文档时出错: {e}"
            print(error_msg)
            self.root.after(0, lambda: self.update_status(error_msg))
    
    def update_status(self, status_msg):
        """更新状态标签，显示最近操作"""
        try:
            original_text = self.status_label.cget("text")
            lines = original_text.split('\n')
            # 保留前几行说明文字，更新最后一行状态
            if len(lines) > 1:
                updated_text = '\n'.join(lines[:-1]) + f"\n最近操作: {status_msg}"
            else:
                updated_text = original_text + f"\n最近操作: {status_msg}"
            self.status_label.config(text=updated_text)
        except Exception as e:
            print(f"更新状态标签时出错: {e}")
    
    def on_key_press(self, key):
        """键盘按键处理"""
        try:
            if hasattr(key, 'char') and key.char:
                if key.char.lower() == 'a':
                    self.save_current_document()
                elif key.char.lower() == 'q':
                    self.quit_program()
        except Exception as e:
            print(f"键盘处理出错: {e}")
    
    def quit_program(self):
        """退出程序"""
        print("正在退出程序...")
        if self.listener:
            self.listener.stop()
        self.root.quit()
        self.root.destroy()
        sys.exit()
    
    def run(self):
        """运行程序主循环"""
        self.root.mainloop()

if __name__ == "__main__":
    try:
        app = DocumentSaverApp()
        app.run()
    except Exception as e:
        print(f"程序启动失败: {e}")
        input("按回车键退出...")