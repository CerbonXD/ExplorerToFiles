import sys
import threading
import subprocess
import time
import json

from win32api import ShellExecute
from win32gui import PostMessage
from win32con import WM_CLOSE
from win32com.client import Dispatch

from pystray import MenuItem, Icon
from PIL import Image

from loguru import logger


exit_event = threading.Event()

def get_explorer_windows(shell):
    return shell.Windows()


@logger.catch(message="Cannot get Explorer folder path from window", default=None)
def get_path_from_window(window):
    return window.Document.Folder.Self.Path


@logger.catch(message="Cannot close Explorer window")
def close_window_by_hwnd(hwnd):
    PostMessage(hwnd, WM_CLOSE, 0, 0)


@logger.catch(message="Error retrieving Files AUMID", default=None)
def get_files_aumid():
    cmd = ['powershell.exe', '-NoProfile', '-Command', 'Get-StartApps | ConvertTo-Json']
    result = subprocess.run(cmd, capture_output=True, text=True, check=True)
    output = result.stdout.strip()

    apps = json.loads(output)

    for app in apps:
        name = app.get("Name", "").strip().lower()

        if name == "files":
            return app.get("AppID")

    logger.info("Files app not found in the output of Get-StartApps.")
    return None


@logger.catch(message="Cannot launch Files app")
def launch_files_app(files_aumid, folder_path):
    ShellExecute(0, 'open', fr'shell:AppsFolder\{files_aumid}', folder_path, None, 1)


@logger.catch(message="Error processing Explorer window")
def monitor_explorer_windows():
    logger.info("Monitoring Explorer windows...")

    files_aumid = get_files_aumid()
    if files_aumid is None: return

    shell = Dispatch("Shell.Application")

    while not exit_event.is_set():
        windows = get_explorer_windows(shell)

        for window in windows:
            hwnd = window.HWND

            folder_path = get_path_from_window(window)
            if folder_path is None: continue
            if folder_path.startswith("::"):
                folder_path = ""

            logger.info(f"Detected Explorer window (HWND {hwnd}) with path: {folder_path}")

            launch_files_app(files_aumid, folder_path)
            logger.info(f"Launched Files app for folder: {folder_path}")

            close_window_by_hwnd(hwnd)

        time.sleep(0.5)


@logger.catch(message="Error opening custom icon", default=None)
def create_image():
    return Image.open("files_redirect_icon.png")


def on_exit(icon):
    exit_event.set()
    icon.stop()


def setup_tray():
    image = create_image()
    menu = (MenuItem('Exit', on_exit),)
    icon = Icon("FilesRedirector", image, "Files Redirector", menu)
    return icon


def main():
    if getattr(sys, "frozen", False):
        logger.remove()

    monitor_thread = threading.Thread(target=monitor_explorer_windows)
    monitor_thread.daemon = True
    monitor_thread.start()

    tray_icon = setup_tray()
    tray_icon.run()

    monitor_thread.join()


if __name__ == "__main__":
    main()