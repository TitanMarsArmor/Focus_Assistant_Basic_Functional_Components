import tkinter as tk
from tkinter import messagebox, ttk
import subprocess
import sys
import os
import threading
import time

class SmartAssistantApp:
    def __init__(self):
        # 创建主窗口
        self.root = tk.Tk()
        self.root.title("智能专注助手")
        self.root.geometry("500x400")
        self.root.resizable(True, True)
        
        # 设置样式
        self.style = ttk.Style()
        self.style.configure("TButton", font=('Arial', 10))
        self.style.configure("TLabel", font=('Arial', 10))
        self.style.configure("Header.TLabel", font=('Arial', 12, 'bold'))
        
        # 创建标题标签
        header_label = ttk.Label(
            self.root,
            text="智能专注助手 - 任务2基础功能部件整合",
            style="Header.TLabel"
        )
        header_label.pack(pady=10)
        
        # 创建功能说明文本框
        info_text = """
        欢迎使用智能专注助手！
        
        本程序整合了任务2的三个基础功能部件：
        1. 音量监控 - 检测到音量不为0且有媒体播放时自动暂停并提示
        2. 快捷键控制 - 使用A键暂停视频，B键切换窗口
        3. 文档保存 - 使用A键快速保存当前活动文档
        
        请选择要启动的功能，或点击"启动全部功能"来运行所有组件。
        """
        self.info_label = ttk.Label(
            self.root,
            text=info_text.strip(),
            justify=tk.LEFT,
            wraplength=480
        )
        self.info_label.pack(pady=10, padx=20, fill=tk.X)
        
        # 创建功能选择框架
        self.functions_frame = ttk.LabelFrame(self.root, text="功能选择")
        self.functions_frame.pack(pady=10, padx=20, fill=tk.X)
        
        # 创建功能复选框变量
        self.volume_var = tk.BooleanVar(value=False)
        self.hotkey_var = tk.BooleanVar(value=False)
        self.saver_var = tk.BooleanVar(value=False)
        
        # 创建功能复选框
        volume_check = ttk.Checkbutton(
            self.functions_frame,
            text="音量监控助手",
            variable=self.volume_var
        )
        volume_check.pack(anchor=tk.W, pady=5, padx=10)
        
        hotkey_check = ttk.Checkbutton(
            self.functions_frame,
            text="快捷键控制器",
            variable=self.hotkey_var
        )
        hotkey_check.pack(anchor=tk.W, pady=5, padx=10)
        
        saver_check = ttk.Checkbutton(
            self.functions_frame,
            text="文档自动保存助手",
            variable=self.saver_var
        )
        saver_check.pack(anchor=tk.W, pady=5, padx=10)
        
        # 创建按钮框架
        self.buttons_frame = ttk.Frame(self.root)
        self.buttons_frame.pack(pady=15)
        
        # 创建控制按钮
        self.start_all_button = ttk.Button(
            self.buttons_frame,
            text="启动全部功能",
            command=self.start_all_functions
        )
        self.start_all_button.pack(side=tk.LEFT, padx=5)
        
        self.start_selected_button = ttk.Button(
            self.buttons_frame,
            text="启动选中功能",
            command=self.start_selected_functions
        )
        self.start_selected_button.pack(side=tk.LEFT, padx=5)
        
        self.stop_all_button = ttk.Button(
            self.buttons_frame,
            text="停止所有功能",
            command=self.stop_all_functions,
            state=tk.DISABLED
        )
        self.stop_all_button.pack(side=tk.LEFT, padx=5)
        
        # 创建退出按钮
        self.quit_button = ttk.Button(
            self.root,
            text="退出程序",
            command=self.quit_program
        )
        self.quit_button.pack(pady=10)
        
        # 创建状态标签
        self.status_var = tk.StringVar(value="就绪 - 请选择要启动的功能")
        self.status_label = ttk.Label(
            self.root,
            textvariable=self.status_var,
            relief=tk.SUNKEN,
            anchor=tk.W
        )
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X, pady=5, padx=5)
        
        # 存储子进程
        self.processes = []
        
        # 获取当前脚本所在目录
        self.script_dir = os.path.dirname(os.path.abspath(__file__))
        
        # 功能脚本映射
        self.function_scripts = {
            'volume': 'volume_monitor.py',
            'hotkey': 'hotkey_controller.py',
            'saver': 'document_saver.py'
        }
    
    def update_status(self, status):
        """更新状态栏文本"""
        self.status_var.set(f"状态: {status}")
        self.root.update_idletasks()
    
    def start_script(self, script_name):
        """启动指定的Python脚本"""
        try:
            script_path = os.path.join(self.script_dir, script_name)
            if not os.path.exists(script_path):
                raise FileNotFoundError(f"脚本文件不存在: {script_path}")
            
            # 启动子进程
            process = subprocess.Popen(
                [sys.executable, script_path],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True
            )
            
            # 添加到进程列表
            self.processes.append(process)
            
            # 更新状态
            script_base = os.path.basename(script_path)
            self.update_status(f"已启动: {script_base}")
            print(f"已启动脚本: {script_base}")
            
            return True
        except Exception as e:
            error_msg = f"启动脚本时出错: {e}"
            print(error_msg)
            messagebox.showerror("启动失败", error_msg)
            return False
    
    def start_all_functions(self):
        """启动所有功能"""
        # 先停止所有正在运行的进程
        self.stop_all_processes()
        
        # 设置所有复选框为选中
        self.volume_var.set(True)
        self.hotkey_var.set(True)
        self.saver_var.set(True)
        
        # 启动所有脚本
        success_count = 0
        for script in self.function_scripts.values():
            if self.start_script(script):
                success_count += 1
                # 添加短暂延迟，避免同时启动多个窗口导致的问题
                time.sleep(0.5)
        
        # 更新按钮状态
        self.update_button_states()
        
        if success_count > 0:
            messagebox.showinfo("启动成功", f"已成功启动 {success_count} 个功能组件")
    
    def start_selected_functions(self):
        """启动选中的功能"""
        # 启动选中的脚本
        success_count = 0
        
        if self.volume_var.get():
            if self.start_script(self.function_scripts['volume']):
                success_count += 1
                time.sleep(0.5)
        
        if self.hotkey_var.get():
            if self.start_script(self.function_scripts['hotkey']):
                success_count += 1
                time.sleep(0.5)
        
        if self.saver_var.get():
            if self.start_script(self.function_scripts['saver']):
                success_count += 1
        
        # 更新按钮状态
        self.update_button_states()
        
        if success_count > 0:
            messagebox.showinfo("启动成功", f"已成功启动 {success_count} 个选中的功能组件")
        else:
            messagebox.showwarning("未启动", "请至少选择一个功能")
    
    def stop_all_processes(self):
        """停止所有子进程，确保强制终止所有子进程"""
        import signal
        import psutil
        
        # 创建进程列表的副本，避免在迭代时修改列表
        processes_to_stop = list(self.processes)
        
        for process in processes_to_stop:
            try:
                # 首先尝试优雅终止
                process.terminate()
                
                # 等待进程结束，增加超时时间到3秒
                try:
                    process.wait(timeout=3)
                    print(f"进程 {process.pid} 已终止")
                except subprocess.TimeoutExpired:
                    # 如果超时，尝试更强制的终止方式
                    try:
                        # 在Windows上，kill()实际上就是terminate()，所以需要特殊处理
                        # 获取进程对象
                        proc = psutil.Process(process.pid)
                        
                        # 强制终止进程树，确保所有子进程也被终止
                        for child in proc.children(recursive=True):
                            try:
                                child.kill()
                            except Exception as e:
                                print(f"终止子进程 {child.pid} 时出错: {e}")
                        
                        # 终止主进程
                        proc.kill()
                        print(f"进程 {process.pid} 已强制终止")
                    except psutil.NoSuchProcess:
                        print(f"进程 {process.pid} 已不存在")
                    except Exception as e:
                        print(f"强制终止进程 {process.pid} 时出错: {e}")
            except Exception as e:
                print(f"停止进程时出错: {e}")
            finally:
                # 确保从列表中移除该进程
                if process in self.processes:
                    self.processes.remove(process)
        
        # 清空进程列表（双重保障）
        self.processes.clear()
        self.update_status("所有功能已停止")
    
    def stop_all_functions(self):
        """停止所有功能"""
        self.stop_all_processes()
        self.update_button_states()
        messagebox.showinfo("已停止", "所有功能组件已停止")
    
    def update_button_states(self):
        """更新按钮状态"""
        has_processes = len(self.processes) > 0
        self.stop_all_button.config(state=tk.NORMAL if has_processes else tk.DISABLED)
    
    def quit_program(self):
        """退出程序"""
        # 询问用户是否确定退出
        if messagebox.askyesno("确认退出", "确定要退出智能专注助手吗？\n所有正在运行的功能将被停止。"):
            # 停止所有进程
            self.stop_all_processes()
            # 退出主程序
            self.root.quit()
            self.root.destroy()
            sys.exit()
    
    def run(self):
        """运行程序主循环"""
        # 设置窗口关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.quit_program)
        
        # 启动主循环
        self.root.mainloop()

if __name__ == "__main__":
    try:
        # 检查是否安装了必要的库
        required_libraries = ['tkinter', 'pynput', 'pyautogui', 'psutil', 'pygame', 'win32gui', 'win32process']
        missing_libraries = []
        
        for lib in required_libraries:
            try:
                if lib == 'tkinter':
                    import tkinter as tk
                else:
                    __import__(lib)
            except ImportError:
                missing_libraries.append(lib)
        
        if missing_libraries:
            msg = f"缺少以下必要的Python库:\n{', '.join(missing_libraries)}\n\n请运行以下命令安装:\npip install {' '.join([lib.replace('win32gui', 'pywin32').replace('win32process', 'pywin32') for lib in missing_libraries])}"
            print(msg)
            messagebox.showerror("缺少依赖", msg)
        else:
            # 启动应用程序
            app = SmartAssistantApp()
            app.run()
    except Exception as e:
        print(f"程序启动失败: {e}")
        input("按回车键退出...")