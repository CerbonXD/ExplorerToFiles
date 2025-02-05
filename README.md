# Explorer To Files

A Python script that runs in the background to monitor any open Windows Explorer window and redirects it to the [Files](https://github.com/files-community/Files) app keeping the same folder path.

I originally created this script for personal use but decided to make it public in case anyone else might want the same feature. I can't guarantee that it will work for everyone, so please feel free to open an issue or submit a pull request if you encounter problems or have improvements.

## Features

- **Background Monitoring:**  
  Constantly monitors for any Windows Explorer window.

- **Redirection:**  
  Redirects Explorer windows to the Files app while preserving the folder path.

- **System Tray Integration:**  
  Runs silently in the background with a system tray icon for easy access and termination.

## Download

You can download the compiled executable **"Explorer To Files.exe"** from the [Releases](https://github.com/CerbonXD/ExplorerToFiles/releases) page.

## How to Use

1. **Start the Program:**  
   After installing the program, double-click the executable to start it.  
   A system tray icon will appear, indicating that the program is running in the background.

2. **Stop the Program:**  
   To stop the program, right-click the system tray icon and select "Exit."

3. **Auto-Start with Windows:**  
   To have the program start automatically when Windows boots:
   - Press `Win + R`, type `shell:startup`, and press Enter.
   - Place a shortcut to the executable in the opened startup folder.

## Building the Project

If you'd like to build the project yourself clone the project and install the dependencies.

### Create .exe file

Install `pyinstaller` and execute the following command:
```bash
pyinstaller --onefile --noconsole --name "Explorer To Files" --icon "files_redirect_icon.png" --add-data "files_redirect_icon.png;." main.py
```

You will find the executable inside `dist` folder.

## License

This project is open source and available under the [MIT License](https://github.com/CerbonXD/ExplorerToFiles/blob/master/LICENSE).
