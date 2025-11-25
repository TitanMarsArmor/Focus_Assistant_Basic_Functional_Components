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
            if hwnd == 0:
                return "无活动窗口", "unknown"
            
            # 获取窗口标题
            window_title = win32gui.GetWindowText(hwnd)
            if not window_title:
                window_title = "无标题窗口"
                
            # 获取进程ID
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            # 获取进程名称
            try:
                process = psutil.Process(pid)
                process_name = process.name().lower()
            except:
                process_name = "unknown_process"
                
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
            # 调试信息 - 开始保存操作
            print("\n===== 开始保存文档操作 =====")
            
            # 获取当前活动窗口信息
            window_title, process_name = self.get_active_window_info()
            
            # 调试信息 - 显示窗口信息
            print(f"检测到的窗口标题: '{window_title}'")
            print(f"检测到的进程名称: '{process_name}'")
            
            # 避免保存本程序自身
            if "pythonw.exe" in process_name or "python.exe" in process_name:
                status_msg = "当前是Python程序窗口，无需保存"
                print(status_msg)
                self.root.after(0, lambda: self.update_status(status_msg))
                return
            
            # 检查是否为文档类应用程序
            is_document = self.is_document_application(process_name)
            print(f"是否为文档应用: {is_document}")
            
            # 选择适当的保存快捷键
            if process_name in self.save_hotkeys:
                hotkey = self.save_hotkeys[process_name]
            else:
                hotkey = self.save_hotkeys['default']  # 使用默认的Ctrl+S
            print(f"使用的保存快捷键: {hotkey}")
            
            # 确保焦点在当前窗口
            hwnd = win32gui.GetForegroundWindow()
            if hwnd:
                win32gui.SetForegroundWindow(hwnd)
                print("已确保焦点在当前窗口")
            
            # 添加一个小延迟，确保在正确的窗口中执行
            time.sleep(0.2)
            
            # 发送保存快捷键
            print("正在发送保存快捷键...")
            pyautogui.hotkey(*hotkey)
            print("保存快捷键已发送")
            
            # 添加一个小延迟，确保保存操作完成
            time.sleep(0.3)
            
            # 更精确地检测保存是否成功 - 可以根据窗口标题变化来判断
            new_window_title = win32gui.GetWindowText(win32gui.GetForegroundWindow())
            if new_window_title != window_title:
                print(f"窗口标题变化: '{window_title}' -> '{new_window_title}'")
            
            # 构建状态消息 - 重点显示窗口标题而非进程名
            if window_title and window_title != "无标题窗口" and window_title != "无活动窗口":
                # 移除可能的路径信息，只显示文件名
                if '\\' in window_title:
                    window_title = window_title.split('\\')[-1]
                status_msg = f"已保存文档: {window_title}"
            else:
                status_msg = f"已尝试保存当前窗口: {process_name}"
            
            print(status_msg)
            print("===== 保存文档操作完成 =====")
            
            # 更新状态标签
            self.root.after(0, lambda: self.update_status(status_msg))
            
        except Exception as e:
            error_msg = f"保存文档时出错: {str(e)}"
            print(error_msg)
            print("===== 保存文档操作失败 =====")
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
                char = key.char.lower()
                if char == 'a':
                    self.save_current_document()
                elif char == 'q':
                    # 获取当前活动窗口的标题
                    current_hwnd = win32gui.GetForegroundWindow()
                    current_title = win32gui.GetWindowText(current_hwnd)
                    
                    # 获取程序窗口的标题
                    app_title = self.root.title()
                    
                    # 调试信息
                    print(f"当前活动窗口标题: '{current_title}'")
                    print(f"程序窗口标题: '{app_title}'")
                    
                    # 检查当前活动窗口是否为程序窗口（通过标题匹配）
                    if app_title in current_title or current_title in app_title:
                        print("检测到程序窗口是当前活动窗口，执行退出操作")
                        self.quit_program()
                    else:
                        print("程序窗口不是当前活动窗口，不执行退出操作")
        except Exception as e:
            print(f"键盘处理出错: {str(e)}")
    
    def quit_program(self):
        """安全退出程序，确保所有资源正确释放"""
        print("正在退出程序...")
        
        # 停止键盘监听器
        if hasattr(self, 'listener') and self.listener:
            try:
                self.listener.stop()
                print("键盘监听器已停止")
            except Exception as e:
                print(f"停止键盘监听器时出错: {e}")
        
        # 安全关闭Tkinter窗口
        try:
            # 先quit，再destroy，确保事件循环停止
            if hasattr(self, 'root') and self.root:
                self.root.quit()
                print("Tkinter主循环已停止")
                # 短暂延迟，确保quit生效
                time.sleep(0.1)
                self.root.destroy()
                print("Tkinter窗口已销毁")
        except Exception as e:
            print(f"关闭Tkinter窗口时出错: {e}")
        
        # 最后退出Python进程
        print("程序退出完成")
        sys.exit(0)
    
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