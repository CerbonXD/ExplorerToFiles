import threading
import subprocess
import json

import win32api
import win32gui
import win32con
import win32com.client

import pystray
from pystray import MenuItem as item
from PIL import Image, ImageDraw

# Global flag for stopping the monitor thread.
should_exit = False

def get_explorer_windows():
    """
    Uses the Shell.Application COM object to enumerate open shell windows.
    Filters only windows that have the typical Explorer window class name.
    """
    shell = win32com.client.Dispatch("Shell.Application")
    explorer_windows = []
    for window in shell.Windows():
        try:
            hwnd = window.HWND
            # Only process windows with class "CabinetWClass" (typical for Explorer)
            class_name = win32gui.GetClassName(hwnd)
            if class_name != "CabinetWClass":
                continue
            # Double-check that the process is explorer.exe
            if window.FullName.lower().endswith("explorer.exe"):
                explorer_windows.append(window)
        except Exception:
            continue
    return explorer_windows


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


def get_path_from_window(window):
    """
    Retrieves the current folder path from an Explorer window.
    Returns None if not available.
    """
    try:
        return window.Document.Folder.Self.Path
    except Exception:
        return None


def close_window_by_hwnd(hwnd):
    """
    Sends a WM_CLOSE message to the window with the given handle.
    """
    try:
        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
    except Exception as e:
        print("Error closing window:", e)


def launch_files_app(files_aumid, folder_path):
    """
    Launches the Files UWP app for the given folder path using its AUMID.
    """
    win32api.ShellExecute(0, 'open', fr'shell:AppsFolder\{files_aumid}', folder_path, None, 1)


def main():
    global should_exit

    files_aumid = get_files_aumid()

    print("Monitoring for Explorer windows...")
    while not should_exit:
        windows = get_explorer_windows()
        for window in windows:
            try:
                hwnd = window.HWND

                folder_path = get_path_from_window(window)
                if not folder_path:
                    continue

                print(f"Detected Explorer window (HWND {hwnd}) with path: {folder_path}")

                # Launch the Files app with the folder path.
                launch_files_app(files_aumid, folder_path)
                print(f"Launched Files app for folder: {folder_path}")

                # Close the original Explorer window.
                close_window_by_hwnd(hwnd)

            except Exception as inner_exc:
                print("Error processing a window:", inner_exc)


def create_image():
    """
    Creates an icon image for the system tray.
    (You can replace this with any image of your choice.)
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
    """
    Callback for the tray menu's "Exit" option.
    """
    global should_exit
    should_exit = True
    icon.stop()

def setup_tray():
    """
    Sets up and returns a pystray Icon instance with an "Exit" menu item.
    """
    image = create_image()
    menu = (item('Exit', on_exit),)
    icon = pystray.Icon("FilesRedirector", image, "Files Redirector", menu)
    return icon

if __name__ == "__main__":
    # Start the monitoring thread.
    monitor_thread = threading.Thread(target=main)
    monitor_thread.daemon = True
    monitor_thread.start()

    # Set up and run the system tray icon.
    tray_icon = setup_tray()
    tray_icon.run()

    # Optionally join the monitor thread on exit.
    monitor_thread.join()
