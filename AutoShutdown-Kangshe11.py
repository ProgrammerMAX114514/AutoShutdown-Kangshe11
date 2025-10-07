# A simple program to shutdown your computer automatically. -*- python3 -*-
# LICENSE: MIT
# Encoding: utf-8

from tkinter import Tk
from tkinter.messagebox import askyesno, showinfo, showerror
from datetime import datetime, timedelta
import sys
import json
from subprocess import call
import os
import ctypes
from ctypes import wintypes
import pystray
from PIL import Image
import threading
from win10toast import ToastNotifier

# ---- Global Constants ----
global_title = "自动关机实用程序"
__version__ = "v0.3.0-dev"
global_suffix = f"""
Autoshutdown-Kangshe11 {__version__}
Powered by Python 3.13.7-64bit
Copyright © 2025 WuRuizhao@Kangshe11"""
# Global variables for logging filename and start time
start_time = ""
file_name = ""

# ---- Safe Exit in Main Thread ----
def _exit(exit_code=0):
    root.after(0, lambda: sys.exit(exit_code))

# ---- System & Admin Related ----
def is_admin():
    """Check if the script is running with administrator privileges."""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def check_admin():
    """Log whether the program is running as admin or not."""
    result = is_admin()
    if result:
        log("INFO", "正在以管理员身份运行程序。")
    else:
        log("WARN", "正在以非管理员身份运行本程序。虽然大部分时候不会出现问题，但某些功能可能会无法正常使用。")

# ---- Time & Logging Setup ----
def set_time():
    """Set the startup timestamp and generate corresponding log filename."""
    global start_time, file_name
    now = datetime.now()
    start_time = now.strftime("%Y-%m-%d_%H:%M:%S")
    file_name = f'{now.strftime("%Y-%m-%d_%H-%M-%S")}.log'

def log(level, msg):
    """
    Write a log message to a dated log file under ./logs/.

    Args:
        level (str): Log level, e.g., 'INFO', 'ERROR'
        msg (str): The message to log
    """
    if level.lower() not in ['debug', 'info', 'warn', 'error', 'fatal']:
        raise ValueError(f'Invalid log level: {level}')
    if "logs" not in os.listdir('.'):
        os.mkdir('logs')
    with open(f'./logs/{file_name}', 'a', encoding='utf-8') as f:
        f.write(f'[{datetime.now().strftime("%H:%M:%S")}] [{level}] {msg}\n')

# ---- Config Loading ----
def load_config(file_path='config.json'):
    """
    Load configuration from a JSON file.

    Args:
        file_path (str): Path to the JSON config file

    Returns:
        dict: Parsed config, or None if failed
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
        return config
    except FileNotFoundError:
        log("ERROR", f"配置文件 {file_path} 未找到！")
        showerror("ERROR", f"配置文件 {file_path} 未找到！{global_suffix}")
        return None
    except json.JSONDecodeError:
        log("ERROR", f"配置文件 {file_path} 格式错误，不是有效的 JSON！")
        showerror("ERROR", f"配置文件 {file_path} 格式错误，不是有效的 JSON！{global_suffix}")
        return None
    except Exception as e:
        log("ERROR", f"无法加载配置文件 {file_path}：{e}")
        showerror("ERROR", f"无法加载配置文件 {file_path}：{e}{global_suffix}")
        return None

# ---- Shutdown Management ----
def shutdown():
    """
    Execute a delayed shutdown command and prompt user to confirm or cancel.
    Uses tkinter's askyesno() which MUST run in the main thread.
    """
    cmd = ['shutdown', '-s', '-t', '60'] # Schedule shutdown after 60 sec
    log("INFO", f"执行关机命令: {cmd}")
    cmd_result = call(cmd)
    if cmd_result == 0:
        log("INFO", "关机命令执行成功")
    else:
        log("ERROR", "关机命令执行失败！")
        showerror(global_title, f"关机命令执行失败！{global_suffix}")

    # Ask user to confirm or cancel the shutdown
    ans = askyesno(global_title, f"""系统将在 1 分钟后自动关机。
要继续关机，请点击“是”，系统将在 1 分钟后关机。
要取消关机，请点击“否”，系统将取消关机，您可以继续使用系统。{global_suffix}""")
    if ans:
        log("INFO", "用户确认关机，程序将退出。")
        _exit()
    if not ans:
        cancel_shutdown()
        log("INFO", "用户取消关机，已调用取消命令。")
        showinfo(global_title, f"已取消关机。{global_suffix}")
        log("INFO", "用户取消关机，程序退出。")
        _exit()
    return

def cancel_shutdown():
    """
    Cancel a previously scheduled shutdown.
    """
    cmd = ['shutdown', '-a']
    log("INFO", f"执行取消关机命令: {cmd}")
    cmd_result = call(cmd)
    if cmd_result == 0:
        log("INFO", "取消关机命令执行成功")
    else:
        log("ERROR", "取消关机命令执行失败！")
        showinfo(global_title, f"取消关机命令执行失败！{global_suffix}")

# ---- Utility Functions ----
def get_current_weekday():
    """Get current weekday in short format (e.g., 'Mon')."""
    return datetime.now().strftime("%a")

# ---- Scheduled Daily Shutdown Logic ----
def scheduled_daily_shutdown():
    """
    Core function to schedule daily shutdowns at configured times.
    Uses threading.Timer to schedule tasks and root.after to safely trigger GUI actions.
    """
    def trigger_shutdown_action():
        """Trigger the shutdown sequence via root.after to ensure thread safety."""
        log("INFO", "到达预设关机时间，准备触发关机流程...")
        root.after(0, shutdown)

    def schedule_for_day(weekday, time_str):
        """Schedule a one-time shutdown for a specific time on a given weekday."""
        try:
            now = datetime.now()
            target_time = datetime.strptime(time_str, "%H:%M").time()
            target_dt = datetime.combine(now.date(), target_time)

            if target_dt <= now:
                return

            delay_seconds = (target_dt - now).total_seconds()
            log("INFO", f"计划于 {target_dt.strftime('%Y-%m-%d %H:%M:%S')} 触发关机 (延迟 {delay_seconds:.0f} 秒)")
            threading.Timer(delay_seconds, trigger_shutdown_action).start()
            return
        except Exception as e:
            log("ERROR", f"安排关机任务失败 [{weekday} {time_str}]: {e}")

    def schedule_all_for_today():
        """Check config and schedule all configured shutdowns for today."""
        weekday = get_current_weekday()
        log("INFO", f"正在为今天 ({weekday}) 安排关机任务...")

        if weekday in config.get("time", {}):
            times = config["time"][weekday]
            if "noon" in times:
                schedule_for_day(weekday, times["noon"])
            if "night" in times:
                schedule_for_day(weekday, times["night"])
        else:
            log("INFO", f"今天 ({weekday}) 没有配置关机时间，跳过。")

    def reschedule_daily():
        """Reschedule all tasks for the next day, recursively."""
        log("INFO", "等待至明日，重新安排关机任务...")
        tomorrow = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0) + timedelta(days=1)
        delay = (tomorrow - datetime.now()).total_seconds()
        threading.Timer(delay, lambda: [schedule_all_for_today(), reschedule_daily()]).start()
    
    # Start the scheduling process
    schedule_all_for_today()
    reschedule_daily()

    # Start tkinter main loop to allow root.after() GUI actions to work
    root.mainloop()

# ---- System Tray (Pystray) Related ----
def on_tray_exit():
    """Callback when user selects 'Exit' from system tray menu."""
    log("INFO", "用户点击了退出按钮。")
    icon.stop()
    _exit(0)
    return

def tray():
    """Initialize and run the system tray icon."""
    global icon
    icon_image = Image.open('icon.ico')
    icon = pystray.Icon("自动关机实用程序", icon_image, menu=pystray.Menu(pystray.MenuItem("退出", on_tray_exit)))
    icon.run()
    return

# ---- Startup Notification ----
def send_startup_notification():
    """Show a Windows toast notification on program start."""
    global __version__
    try:
        toaster = ToastNotifier()
        toaster.show_toast(
            title="自动关机实用程序",
            msg=f"“自动关机实用程序” v{__version__}已启动，正在后台运行。\n右键托盘图标可以退出。",
            icon_path="icon.ico",
            duration=10,
            threaded=True
        )
        log("INFO", "已发送启动通知。")
    except Exception as e:
        log("WARNING", f"发送启动通知失败: {e}")

# ---- Prevent Multiple Instances ----
def prevent_multiple_instances():
    """Prevent the program from running multiple instances using a Windows mutex."""
    MUTEX_NAME = "Global\\AutoShutdownUtil_Mutex_2024"

    kernel32 = ctypes.WinDLL('kernel32')
    CreateMutex = kernel32.CreateMutexW
    CreateMutex.argtypes = [wintypes.LPVOID, wintypes.BOOL, wintypes.LPCWSTR]
    CreateMutex.restype = wintypes.HANDLE

    GetLastError = kernel32.GetLastError
    GetLastError.restype = wintypes.DWORD

    ERROR_ALREADY_EXISTS = 183

    mutex = CreateMutex(None, False, MUTEX_NAME)
    last_error = GetLastError()

    if last_error == ERROR_ALREADY_EXISTS:
        log("WARN", "检测到程序已在运行，阻止新实例启动。")
        showerror(global_title, f"程序已经在运行中，请勿多开程序！{global_suffix}")
        _exit(1)
    elif mutex:
        log("INFO", "程序启动成功，当前为唯一运行实例。")
    else:
        log("ERROR", "无法创建互斥锁，未知错误。")
        showerror(global_title, f"程序启动失败！{global_suffix}")
        _exit(1)

# ---- Program Entry Point ----
if __name__ == "__main__":
    # Create hidden Tk root window (required for tkinter dialogs and root.after)
    root = Tk()
    root.attributes('-topmost', True)
    root.withdraw()
    try:
        # Load configuration
        try:
            config = load_config()
        except:
            log("ERROR","无法加载配置文件！")
            showerror(global_title, f"无法加载配置文件！{global_suffix}")
            _exit(1)
        set_time()
        check_admin()
        log("INFO", f"程序启动的时间是 {start_time}。")
        log("INFO", "正在初始化……")
        prevent_multiple_instances()
        log("INFO", "初始化完成。")
        send_startup_notification()

        # Start system tray (in a daemon thread)
        tray_thread = threading.Thread(target=tray, daemon=True)
        tray_thread.start()
        # Start the daily shutdown scheduler
        scheduled_daily_shutdown()
    except Exception as e:
        log("ERROR", f"程序异常退出！异常信息: {e}")
        showerror(global_title, f"An error occurred while running the program:\n{e}\nTurn to the log file for more details.{global_suffix}")
