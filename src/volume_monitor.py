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
import win32com.client
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

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
            text="音量监控助手正在运行...\n\n功能说明:\n- 检测到音量不为0且有媒体播放时自动设为静音\n- 显示弹窗提示当前为外放状态\n- 按 Q 键: 退出程序\n\n提示: 程序会定期检查系统音量状态",
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
        # 添加静音设置禁用时间，初始为0表示未禁用
        self.mute_disabled_until = 0
        self.monitor_thread = threading.Thread(target=self.monitor_volume)
        self.monitor_thread.daemon = True
        self.monitor_thread.start()
    
    def get_system_volume(self):
        """获取系统音量，检测是否音量不为0且非静音模式"""
        try:
            # 优先尝试使用pycaw获取音量，但使用更兼容的方式
            try:
                # 导入必要的模块
                from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
                from ctypes import cast, POINTER
                from comtypes import CLSCTX_ALL
                
                # 获取所有活动的音频端点
                devices = AudioUtilities.GetSpeakers()
                
                # 尝试获取音量接口
                try:
                    # 使用正确的pycaw方法获取音量
                    volume = devices.EndpointVolume
                    
                    # 获取音量级别（0.0到1.0）
                    current_volume = volume.GetMasterVolumeLevelScalar()
                    
                    # 检查是否静音
                    is_muted = volume.GetMute()
                    
                    # 如果音量大于0且非静音模式，则返回True
                    has_sound = current_volume > 0.01 and not is_muted
                    print(f"音量检测 (pycaw): 级别={current_volume:.2f}, 静音={is_muted}, 有声音={has_sound}")
                    return has_sound
                except Exception as activation_error:
                    print(f"pycaw激活失败: {activation_error}")
                    
                    # 尝试获取所有音频会话并检查是否有活动的声音
                    sessions = AudioUtilities.GetAllSessions()
                    active_sessions = 0
                    
                    # 定义媒体播放器进程列表（为了避免重复定义，直接在这里列出常见的）
                    media_players = ['vlc.exe', 'potplayermini64.exe', 'potplayer64.exe', 'mpc-hc.exe', 'mpc-hc64.exe',
                                    'wmplayer.exe', 'mplayer2.exe', 'kodi.exe', 'plexmediaplayer.exe', 'mxplayer.exe',
                                    'spotify.exe', 'itunes.exe', 'qqmusic.exe', 'neteasemusic.exe', 'kuwo.exe',
                                    'foobar2000.exe', 'winamp.exe', 'groove.exe', 'musicbee.exe', 'aimp.exe',
                                    'chrome.exe', 'firefox.exe', 'edge.exe', 'opera.exe', 'brave.exe', 'safari.exe']
                    
                    for session in sessions:
                        if session.State == 1:  # 活动状态
                            volume = session.SimpleAudioVolume
                            if volume and not volume.GetMute() and volume.GetMasterVolume() > 0.0:
                                active_sessions += 1
                                # 只对媒体播放器进程输出日志
                                if session.Process and session.Process.name().lower() in media_players:
                                    print(f"检测到活动音频会话: {session.Process.name()}")
                    
                    has_active_sound = active_sessions > 0
                    # 不输出会话数量，避免暴露不必要的信息
                    print(f"音频会话检测: 有声音={has_active_sound}")
                    return has_active_sound
            except Exception as pycaw_error:
                print(f"pycaw方法失败: {pycaw_error}")
                
                # 备用方案：使用PowerShell命令获取系统音量
                try:
                    import subprocess
                    import json
                    
                    # 使用PowerShell获取系统音量和静音状态
                    ps_command = "(Get-SoundVolume).MasterVolumeLevelPercentage"  # 注意：这可能需要安装额外模块
                    result = subprocess.run(["powershell", "-Command", ps_command], 
                                          capture_output=True, text=True, timeout=2)
                    
                    if result.returncode == 0 and result.stdout.strip():
                        volume_percent = float(result.stdout.strip())
                        has_sound = volume_percent > 1.0  # 音量大于1%认为有声音
                        print(f"PowerShell音量检测: 级别={volume_percent:.2f}%, 有声音={has_sound}")
                        return has_sound
                except Exception as ps_error:
                    print(f"PowerShell方法失败: {ps_error}")
                    
                    # 备用方案：使用win32api（但不作为主要方法，因为它可能不准确）
                    try:
                        from ctypes import windll, byref, c_uint
                        winmm = windll.winmm
                        
                        # 设置音量变量
                        volume = c_uint(0)
                        
                        # 获取当前音量
                        result = winmm.waveOutGetVolume(0, byref(volume))
                        
                        if result == 0:  # 成功获取音量
                            # 分离左右声道音量
                            left_volume = volume.value & 0xFFFF
                            right_volume = (volume.value >> 16) & 0xFFFF
                            
                            # 计算平均音量（0-1范围）
                            average_volume = ((left_volume + right_volume) / 2) / 65535.0
                            
                            # 重要：不要仅依赖这个值，添加额外的检查逻辑
                            # 由于winmm可能不准确，我们将阈值提高到20%，避免误判
                            has_sound = average_volume > 0.20
                            print(f"winmm音量检测: 级别={average_volume:.2f}, 有声音={has_sound}")
                            return has_sound
                        else:
                            print(f"winmm音量获取失败，错误代码: {result}")
                    except Exception as winmm_error:
                        print(f"winmm方法失败: {winmm_error}")
            
            # 如果所有方法都失败，返回False以避免误触发
            return False
        except Exception as e:
            print(f"获取音量时发生异常: {e}")
            return False
    
    def is_media_playing(self):
        """检测是否有媒体播放器正在播放音频/视频"""
        try:
            # 导入psutil模块
            import psutil
            
            # 媒体播放器进程列表（分组管理）
            media_players = [
                # 视频播放器
                'vlc.exe', 'potplayermini64.exe', 'potplayer64.exe', 'mpc-hc.exe', 'mpc-hc64.exe',
                'wmplayer.exe', 'mplayer2.exe', 'kodi.exe', 'plexmediaplayer.exe', 'mxplayer.exe',
                # 音乐播放器
                'spotify.exe', 'itunes.exe', 'qqmusic.exe', 'neteasemusic.exe', 'kuwo.exe',
                'foobar2000.exe', 'winamp.exe', 'groove.exe', 'musicbee.exe', 'aimp.exe',
                # 浏览器
                'chrome.exe', 'firefox.exe', 'edge.exe', 'opera.exe', 'brave.exe', 'safari.exe',
            ]
            
            # 首先，尝试获取所有活动的音频会话，但只检查media_players中的程序
            try:
                from pycaw.pycaw import AudioUtilities
                
                sessions = AudioUtilities.GetAllSessions()
                for session in sessions:
                    if session.State == 1:  # 活动状态
                        volume = session.SimpleAudioVolume
                        if volume and not volume.GetMute() and volume.GetMasterVolume() > 0.0 and session.Process:
                            process_name = session.Process.name().lower()
                            # 只输出并检测media_players中的程序
                            if process_name in media_players:
                                print(f"检测到活动音频会话: {process_name}")
                                print(f"通过音频会话检测确认正在播放媒体")
                                return True
            except Exception as session_error:
                print(f"音频会话检测失败: {session_error}")
            
            # 检查常见媒体播放器进程
            for proc in psutil.process_iter(['name']):
                try:
                    process_name = proc.info['name'].lower()
                    
                    if process_name in media_players:
                        print(f"检测到媒体相关进程: {process_name}")
                        return True
                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                    pass
            
            # 检查Windows Media Player COM对象
            try:
                import win32com.client
                wmp = win32com.client.Dispatch("WMPlayer.OCX")
                if hasattr(wmp, 'playState') and wmp.playState == 3:  # 播放状态为3表示正在播放
                    print("Windows Media Player正在播放")
                    return True
            except Exception as e:
                print(f"WMPlayer检测失败: {e}")
            
            print("未检测到正在播放的媒体")
            return False
        except Exception as e:
            print(f"媒体检测出错: {e}")
            return False
    
    def set_system_mute(self):
        """设置系统为静音状态"""
        try:
            # 方法1：尝试使用pycaw设置系统静音
            try:
                from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
                from ctypes import cast, POINTER
                from comtypes import CLSCTX_ALL
                
                # 获取默认音频输出设备
                devices = AudioUtilities.GetSpeakers()
                
                # 使用正确的pycaw方法获取并设置音量
                try:
                    volume = devices.EndpointVolume
                    
                    # 设置为静音
                    volume.SetMute(1, None)
                    print("已设置系统为静音 (pycaw方法)")
                    return True
                except AttributeError:
                    # 如果Activate方法不存在，尝试通过会话设置静音
                    sessions = AudioUtilities.GetAllSessions()
                    for session in sessions:
                        if session.State == 1:  # 活动状态
                            volume = session.SimpleAudioVolume
                            if volume:
                                volume.SetMute(1, None)
                    print("已设置所有活动音频会话为静音 (会话方法)")
                    return True
            except Exception as pycaw_error:
                print(f"pycaw静音方法失败: {pycaw_error}")
                
                # 方法2：尝试使用win32api设置系统静音
                try:
                    import ctypes
                    
                    # 定义发送按键的函数
                    def press_key(key_code):
                        ctypes.windll.user32.keybd_event(key_code, 0, 0, 0)  # 按下键
                        ctypes.windll.user32.keybd_event(key_code, 0, 0x0002, 0)  # 释放键
                    
                    # 模拟按下静音快捷键 (Windows键 + F3)
                    VK_LWIN = 0x5B
                    VK_F3 = 0x72
                    
                    # 按下Windows键
                    ctypes.windll.user32.keybd_event(VK_LWIN, 0, 0, 0)
                    # 按下F3键
                    ctypes.windll.user32.keybd_event(VK_F3, 0, 0, 0)
                    # 释放F3键
                    ctypes.windll.user32.keybd_event(VK_F3, 0, 0x0002, 0)
                    # 释放Windows键
                    ctypes.windll.user32.keybd_event(VK_LWIN, 0, 0x0002, 0)
                    
                    print("已模拟静音快捷键 (Windows + F3)")
                    return True
                except Exception as win32_error:
                    print(f"win32api静音方法失败: {win32_error}")
                    
                    # 方法3：使用PowerShell命令设置系统静音
                    try:
                        import subprocess
                        
                        # PowerShell命令来切换静音状态
                        ps_command = "(New-Object -ComObject WScript.Shell).SendKeys([char]173)"
                        result = subprocess.run(["powershell", "-Command", ps_command], 
                                              capture_output=True, text=True, timeout=2)
                        
                        if result.returncode == 0:
                            print("已通过PowerShell设置系统静音")
                            return True
                        else:
                            print(f"PowerShell命令失败: {result.stderr}")
                    except Exception as ps_error:
                        print(f"PowerShell静音方法失败: {ps_error}")
            
            # 如果所有方法都失败
            return False
        except Exception as e:
            print(f"设置系统静音时发生异常: {e}")
            return False
    
    def unmute_system(self):
        """取消系统静音状态"""
        try:
            # 方法1：尝试使用pycaw取消系统静音
            try:
                from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
                from ctypes import cast, POINTER
                from comtypes import CLSCTX_ALL
                
                # 获取默认音频输出设备
                devices = AudioUtilities.GetSpeakers()
                
                # 使用正确的pycaw方法获取并设置音量
                try:
                    volume = devices.EndpointVolume
                    
                    # 取消静音
                    volume.SetMute(0, None)
                    # 设置一个合适的音量级别
                    volume.SetMasterVolumeLevelScalar(0.3, None)  # 30%音量
                    print("已取消系统静音并设置音量 (pycaw方法)")
                    return True
                except AttributeError:
                    # 如果Activate方法不存在，尝试通过会话取消静音
                    sessions = AudioUtilities.GetAllSessions()
                    for session in sessions:
                        if session.State == 1:  # 活动状态
                            volume = session.SimpleAudioVolume
                            if volume:
                                volume.SetMute(0, None)
                    print("已取消所有活动音频会话静音 (会话方法)")
                    return True
            except Exception as pycaw_error:
                print(f"pycaw取消静音方法失败: {pycaw_error}")
                
                # 方法2：尝试使用win32api取消系统静音
                try:
                    import ctypes
                    
                    # 再次模拟按下静音快捷键来取消静音
                    VK_LWIN = 0x5B
                    VK_F3 = 0x72
                    
                    # 按下Windows键
                    ctypes.windll.user32.keybd_event(VK_LWIN, 0, 0, 0)
                    # 按下F3键
                    ctypes.windll.user32.keybd_event(VK_F3, 0, 0, 0)
                    # 释放F3键
                    ctypes.windll.user32.keybd_event(VK_F3, 0, 0x0002, 0)
                    # 释放Windows键
                    ctypes.windll.user32.keybd_event(VK_LWIN, 0, 0x0002, 0)
                    
                    print("已模拟取消静音快捷键 (Windows + F3)")
                    return True
                except Exception as win32_error:
                    print(f"win32api取消静音方法失败: {win32_error}")
                    
                    # 方法3：使用PowerShell命令取消系统静音
                    try:
                        import subprocess
                        
                        # 再次使用PowerShell命令来切换静音状态
                        ps_command = "(New-Object -ComObject WScript.Shell).SendKeys([char]173)"
                        result = subprocess.run(["powershell", "-Command", ps_command], 
                                              capture_output=True, text=True, timeout=2)
                        
                        if result.returncode == 0:
                            print("已通过PowerShell取消系统静音")
                            return True
                        else:
                            print(f"PowerShell命令失败: {result.stderr}")
                    except Exception as ps_error:
                        print(f"PowerShell取消静音方法失败: {ps_error}")
            
            # 如果所有方法都失败
            return False
        except Exception as e:
            print(f"取消系统静音时发生异常: {e}")
            return False
    
    def monitor_volume(self):
        """监控音量状态的线程函数"""
        while self.monitoring:
            try:
                # 检查是否处于禁用静音设置的时间窗口内
                current_time = time.time()
                if current_time < self.mute_disabled_until:
                    remaining_time = int(self.mute_disabled_until - current_time)
                    print(f"静音设置已暂时禁用，剩余时间: {remaining_time}秒")
                    # 等待10秒后再次检查
                    time.sleep(10)
                    continue
                
                # 检查音量和媒体播放状态
                if self.get_system_volume() and self.is_media_playing():
                    # 设置系统静音
                    self.set_system_mute()
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
        """显示音量警告消息框（顶层窗口，带取消静音选项）"""
        try:
            # 创建一个自定义的顶级窗口，设置为顶层
            dialog = tk.Toplevel(self.root)
            dialog.title("音量提示")
            dialog.geometry("350x180")
            dialog.resizable(False, False)
            # 设置为顶层窗口
            dialog.attributes('-topmost', True)
            # 禁用父窗口
            dialog.grab_set()
            
            # 计算窗口位置，使其显示在屏幕中间
            dialog.update_idletasks()
            width = dialog.winfo_width()
            height = dialog.winfo_height()
            x = (dialog.winfo_screenwidth() // 2) - (width // 2)
            y = (dialog.winfo_screenheight() // 2) - (height // 2)
            dialog.geometry('{}x{}+{}+{}'.format(width, height, x, y))
            
            # 添加提示信息
            message = tk.Label(dialog, text="检测到当前为外放状态，已设置系统为静音", wraplength=330, pady=20)
            message.pack()
            
            # 创建按钮框架
            button_frame = tk.Frame(dialog)
            button_frame.pack(pady=10)
            
            # 创建确定按钮
            ok_button = tk.Button(button_frame, text="确定", width=10, command=dialog.destroy)
            ok_button.pack(side=tk.LEFT, padx=5)
            
            # 创建取消静音设置按钮
            cancel_mute_button = tk.Button(button_frame, text="取消静音设置", width=15, command=lambda: self.handle_cancel_mute(dialog))
            cancel_mute_button.pack(side=tk.LEFT, padx=5)
            
        except Exception as e:
            print(f"显示消息框时出错: {e}")
    
    def handle_cancel_mute(self, dialog):
        """处理取消静音设置的逻辑"""
        try:
            # 取消系统静音
            self.unmute_system()
            
            # 设置5分钟（300秒）内不再监测音量
            self.mute_disabled_until = time.time() + 300
            print("已取消静音设置，5分钟内不再监测音量")
            
            # 关闭对话框
            dialog.destroy()
            
            # 显示一个确认消息，增大窗口尺寸以确保文字完全显示
            confirm_dialog = tk.Toplevel(self.root)
            confirm_dialog.title("操作确认")
            confirm_dialog.geometry("320x120")
            confirm_dialog.resizable(False, False)
            confirm_dialog.attributes('-topmost', True)
            
            # 计算窗口位置，使其显示在屏幕中间
            confirm_dialog.update_idletasks()
            width = confirm_dialog.winfo_width()
            height = confirm_dialog.winfo_height()
            x = (confirm_dialog.winfo_screenwidth() // 2) - (width // 2)
            y = (confirm_dialog.winfo_screenheight() // 2) - (height // 2)
            confirm_dialog.geometry('{}x{}+{}+{}'.format(width, height, x, y))
            
            # 增加wraplength确保文字能完全显示
            confirm_label = tk.Label(confirm_dialog, text="已取消静音设置，5分钟内不再监测音量", wraplength=300)
            confirm_label.pack(pady=20)
            
            ok_button = tk.Button(confirm_dialog, text="确定", command=confirm_dialog.destroy)
            ok_button.pack()
            
        except Exception as e:
            print(f"处理取消静音设置时出错: {e}")
    
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