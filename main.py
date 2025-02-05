import threading
import subprocess
import time
import json

from win32api import ShellExecute
from win32gui import PostMessage
from win32con import WM_CLOSE
from win32com.client import Dispatch

from pystray import MenuItem, Icon
from PIL import Image, ImageDraw


should_exit = False


def get_explorer_windows():
    shell = Dispatch("Shell.Application")
    return shell.Windows()


def get_path_from_window(window):
    try:
        return window.Document.Folder.Self.Path

    except Exception as e:
        print("Cannot get Explorer folder path from window:", e)

    return None


def close_window_by_hwnd(hwnd):
    try:
        PostMessage(hwnd, WM_CLOSE, 0, 0)

    except Exception as e:
        print("Cannot close Explorer window:", e)


def get_files_aumid():
    try:
        cmd = ['powershell.exe', '-NoProfile', '-Command', 'Get-StartApps | ConvertTo-Json']
        result = subprocess.run(cmd, capture_output=True, text=True, check=True)
        output = result.stdout.strip()

        apps = json.loads(output)

        for app in apps:
            name = app.get("Name", "").strip().lower()

            if name == "files":
                return app.get("AppID")

        print("Files app not found in the output of Get-StartApps.")

    except Exception as e:
        print("Error retrieving Files AUMID:", e)

    return None


def launch_files_app(files_aumid, folder_path):
    try:
        ShellExecute(0, 'open', fr'shell:AppsFolder\{files_aumid}', folder_path, None, 1)

    except Exception as e:
        print("Cannot launch Files app:", e)


def monitor_explorer_windows():
    print("Monitoring Explorer windows...")

    files_aumid = get_files_aumid()
    if files_aumid is None: return

    global should_exit
    while not should_exit:
        windows = get_explorer_windows()

        for window in windows:
            try:
                hwnd = window.HWND

                folder_path = get_path_from_window(window)
                if folder_path is None: continue
                if folder_path.startswith("::"):
                    folder_path = ""

                print(f"Detected Explorer window (HWND {hwnd}) with path: {folder_path}")

                launch_files_app(files_aumid, folder_path)
                print(f"Launched Files app for folder: {folder_path}")

                close_window_by_hwnd(hwnd)

            except Exception as e:
                print("Error processing Explorer window:", e)

        time.sleep(0.5)


def create_image():
    """
    Creates an icon image for the system tray.
    """
    width = 64
    height = 64
    # Background and foreground colors.
    bg_color = (0, 0, 0)
    fg_color = (255, 255, 255)
    image = Image.new('RGB', (width, height), bg_color)
    dc = ImageDraw.Draw(image)
    # Draw a simple rectangle in the middle.
    dc.rectangle(
        (width // 4, height // 4, width * 3 // 4, height * 3 // 4),
        fill=fg_color)
    return image


def on_exit(icon):
    global should_exit
    should_exit = True
    icon.stop()


def setup_tray():
    image = create_image()
    menu = (MenuItem('Exit', on_exit),)
    icon = Icon("FilesRedirector", image, "Files Redirector", menu)
    return icon


def main():
    monitor_thread = threading.Thread(target=monitor_explorer_windows)
    monitor_thread.daemon = True
    monitor_thread.start()

    tray_icon = setup_tray()
    tray_icon.run()

    monitor_thread.join()


if __name__ == "__main__":
    main()