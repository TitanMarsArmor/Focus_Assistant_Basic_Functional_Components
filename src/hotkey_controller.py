import tkinter as tk
from pynput import keyboard
import pyautogui
import sys
import threading
import time

class HotkeyControllerApp:
    def __init__(self):
        # 创建主窗口
        self.root = tk.Tk()
        self.root.title("快捷键控制器")
        self.root.geometry("400x300")
        
        # 创建状态标签
        self.status_label = tk.Label(
            self.root,
            text="快捷键控制器正在运行...\n\n可用快捷键:\n- 按 A 键: 暂停/播放视频\n- 按 B 键: 切换到上一个窗口 (Alt+Tab)\n- 按 Q 键: 退出程序\n\n功能说明:\n- 无论焦点在哪个窗口，都可以使用这些全局快捷键\n- 程序运行在后台，随时响应按键操作\n- 所有操作都会在控制台输出日志",
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
        
        # 记录最后一次操作时间，避免频繁操作
        self.last_operation_time = {}
        self.cooldown_time = 0.5  # 冷却时间，单位秒
    
    def is_cooldown_over(self, operation_name):
        """检查操作冷却时间是否已过"""
        current_time = time.time()
        if operation_name in self.last_operation_time:
            if current_time - self.last_operation_time[operation_name] < self.cooldown_time:
                return False
        self.last_operation_time[operation_name] = current_time
        return True
    
    def pause_video(self):
        """暂停或播放视频"""
        if not self.is_cooldown_over('pause_video'):
            return
        
        try:
            # 发送空格键，大多数视频播放器使用空格键暂停/播放
            pyautogui.press('space')
            print("已发送暂停/播放命令")
            # 更新状态标签
            self.root.after(0, lambda: self.update_status("已暂停/播放视频"))
        except Exception as e:
            error_msg = f"暂停视频时出错: {e}"
            print(error_msg)
            self.root.after(0, lambda: self.update_status(error_msg))
    
    def switch_window(self):
        """切换到上一个窗口 (Alt+Tab)"""
        if not self.is_cooldown_over('switch_window'):
            return
        
        try:
            # 发送Alt+Tab组合键，切换到上一个窗口
            pyautogui.hotkey('alt', 'tab')
            print("已切换到上一个窗口")
            # 更新状态标签
            self.root.after(0, lambda: self.update_status("已切换到上一个窗口"))
        except Exception as e:
            error_msg = f"切换窗口时出错: {e}"
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
                char = key.char.lower()
                if char == 'a':
                    self.pause_video()
                elif char == 'b':
                    self.switch_window()
                elif char == 'q':
                    self.quit_program()
        except Exception as e:
            print(f"键盘处理出错: {e}")
    
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
        app = HotkeyControllerApp()
        app.run()
    except Exception as e:
        print(f"程序启动失败: {e}")
        input("按回车键退出...")