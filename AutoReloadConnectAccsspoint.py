import pygetwindow as gw
from pywinauto import application, keyboard
import subprocess
import time
import datetime as dt
import os
import pyautogui
import pyperclip
import win32gui
import win32con
import threading
import logging

# Configure logging
logging.basicConfig(filename='putty_automation.log', level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')

def find_putty():
    folder_path = "C:\\Program Files\\PuTTY"
    for file in os.listdir(folder_path):
        if file.endswith(".exe") and file.startswith("putty"):
            return os.path.join(folder_path, file)
    return None

def handle_fatal_error(clipboard_content):
    try:
        error_window = gw.getWindowsWithTitle('PuTTY Fatal Error')[0]
        app = application.Application().connect(handle=error_window._hWnd)
        dlg = app.window(handle=error_window._hWnd)
        dlg['OK'].click()
        if "Network error: Connection timed out" not in clipboard_content:
            logging.info('Change Status to Connected Mode Completed')
        error_window.close()
        logging.info("Handled PuTTY Fatal Error")
    except Exception as e:
        logging.error(f"Error handling PuTTY Fatal Error: {e}")

def close_inactive_putty():
    try:
        for window in gw.getAllWindows():
            if window.title.endswith('PuTTY (inactive)'):
                app = application.Application().connect(handle=window._hWnd)
                dlg = app.window(handle=window._hWnd)
                dlg.close()
                logging.info(f"Closed inactive PuTTY window: {window.title}")
    except Exception as e:
        logging.error(f"Error closing inactive PuTTY window: {e}")

def move_window_to_top_left(window):
    win32gui.SetWindowPos(window, win32con.HWND_TOP, 0, 0, 0, 0, win32con.SWP_NOSIZE | win32con.SWP_NOZORDER)

def save_log(ap_statuses):
    current_date = dt.datetime.now().strftime("%d-%m-%Y")
    with open(f'{current_date}.txt', 'a+', encoding='utf-8') as SaveFile:
        current_time = dt.datetime.now().strftime("%d/%m/%Y %I:%M:%S %p")
        SaveFile.write(f'Check AP Status at {current_time}\n')
        SaveFile.write('\nSummary of AP statuses:\n')
        for status in ap_statuses:
            SaveFile.write(f"{status['index']}. {status['ip']} - {status['status']}\n")
        SaveFile.write(f'Total APs checked: {len(ap_statuses)}\n')

def process_ip(ip, index, ap_statuses):
    program_path = find_putty()
    if not program_path:
        logging.error("putty.exe not found in the specified folder.")
        return

    subprocess.Popen(program_path)
    time.sleep(2)

    window_title = None
    for title in gw.getAllTitles():
        if "PuTTY Configuration" in title:
            window_title = title
            break

    if not window_title:
        logging.error("No PuTTY window found")
        return

    window = gw.getWindowsWithTitle(window_title)[0]
    app = application.Application().connect(handle=window._hWnd)
    dlg = app.window(handle=window._hWnd)
    dlg.set_focus()
    dlg['Host Name (or IP address):Edit'].type_keys(ip)
    dlg['SSH'].click()
    dlg['Open'].click()

    logging.info(f"Filled the Host Name field with {ip}, selected SSH, and clicked Open")

    terminal_window = None
    security_alert_window = None

    while not terminal_window:
        for window in gw.getAllWindows():
            if window.title.startswith('PuTTY Security Alert'):
                security_alert_window = window
                break
            if window.title.startswith(f'{ip} - PuTTY'):
                terminal_window = window
                break

        if security_alert_window:
            app = application.Application().connect(handle=security_alert_window._hWnd)
            dlg = app.window(handle=security_alert_window._hWnd)
            dlg['Accept'].click()
            logging.info("Accepted PuTTY Security Alert")
            time.sleep(1)
            security_alert_window = None

    if terminal_window:
        app = application.Application().connect(handle=terminal_window._hWnd)
        terminal = app.window(handle=terminal_window._hWnd)
        terminal.set_focus()
        time.sleep(2)
        move_window_to_top_left(terminal_window._hWnd)
        terminal.set_focus()
        keyboard.send_keys('admin{ENTER}')
        time.sleep(1)
        keyboard.send_keys('N@xtech!234{ENTER}')
        logging.info(f"Entered username and password, and sent the command on {ip}")
        time.sleep(1)
        keyboard.send_keys('show {SPACE} flexconnect {SPACE} status{ENTER}')
        time.sleep(2)

        terminal.set_focus()
        x = 130
        y = 140
        pyautogui.click(x=x, y=y)
        pyautogui.click(x=x, y=y)
        pyautogui.click(x=x, y=y)
        pyautogui.hotkey('ctrl', 'c')

        time.sleep(2)

        clipboard_content = pyperclip.paste()
        if "AP in Connected Mode" in clipboard_content:
            logging.info(f"AP {ip} in Connected Mode")
            logging.info(f'{clipboard_content}')
            ap_statuses.append({'index': index, 'ip': ip, 'status': 'Connected Mode'})
            keyboard.send_keys('exit{ENTER}')
        elif "Network error: Connection timed out" in clipboard_content:
            logging.warning(f"AP {ip} is Down")
            logging.warning(f'{clipboard_content}')
            ap_statuses.append({'index': index, 'ip': ip, 'status': 'Down'})
            handle_fatal_error(clipboard_content)
            close_inactive_putty()
            time.sleep(3)
        else:
            logging.info(f"AP {ip} in StandAlone Mode")
            logging.info(f'{clipboard_content}')
            ap_statuses.append({'index': index, 'ip': ip, 'status': 'StandAlone Mode'})
            keyboard.send_keys('en{ENTER}')
            time.sleep(1)
            keyboard.send_keys('N@xtech!234{ENTER}')
            time.sleep(1)
            keyboard.send_keys('reload{ENTER}')
            time.sleep(1)
            keyboard.send_keys('{ENTER}')
            time.sleep(3)
            handle_fatal_error(clipboard_content)
            time.sleep(1)
            close_inactive_putty()
        logging.info(f"Handling AP {ip} completed")
    else:
        logging.error("No PuTTY terminal window found")
    time.sleep(2)

def main():
    
    LPG = [
        "172.21.0.202", "172.21.0.207", "172.21.0.204",
        "172.21.0.206", "172.21.0.203", "172.21.0.205", "172.21.0.208", "172.21.0.209",
    ]
    SRI = [
        "172.22.0.205", "172.22.0.203", "172.22.0.204", "172.22.0.206", "172.22.0.201",
        "172.22.0.202"
    ]
    
    KKN = [
        "172.23.0.203", "172.23.0.208", "172.23.0.210", "172.23.0.207", "172.23.0.206",
        "172.23.0.209", "172.23.0.201"
    ]
    
    SNI = [
        "172.24.0.211", "172.24.0.203", "172.24.0.205", "172.24.0.204",
        "172.24.0.207", "172.24.0.202"
    ]
    
    All = [
        "172.23.0.203", "172.23.0.208", "172.23.0.210", "172.23.0.207", "172.23.0.206",
        "172.23.0.209", "172.23.0.201", "172.21.0.202", "172.21.0.207", "172.21.0.204",
        "172.21.0.206", "172.21.0.203", "172.21.0.205", "172.21.0.208", "172.21.0.209",
        "172.22.0.205", "172.22.0.203", "172.22.0.204", "172.22.0.206", "172.22.0.201",
        "172.22.0.202", "172.24.0.211", "172.24.0.203", "172.24.0.205", "172.24.0.204",
        "172.24.0.207", "172.24.0.202"
    ]
    
    central = [
    "172.16.201.1", "172.16.201.2", "172.16.201.3", "172.16.201.4", "172.16.201.5",
    "172.16.201.6", "172.16.201.7", "172.16.201.8", "172.16.201.12", "172.16.201.13",
    "172.16.201.14", "172.16.201.15", "172.16.201.16", "172.16.201.17", "172.16.201.18",
    "172.16.201.20", "172.16.201.21", "172.16.201.22", "172.16.201.23", "172.16.201.24",
    "172.16.201.25", "172.16.201.26", "172.16.201.27", "172.16.201.28", "172.16.201.29",
    "172.16.201.30", "172.16.201.31", "172.16.201.32", "172.16.201.33", "172.16.201.34",
    "172.16.201.35", "172.16.201.36", "172.16.201.37", "172.16.201.38", "172.16.201.39",
    "172.16.201.40", "172.16.201.41", "172.16.201.42", "172.16.201.43", "172.16.201.44",
    "172.16.201.45", "172.16.201.47", "172.16.201.48", "172.16.201.49", "172.16.201.51",
    "172.16.201.52", "172.16.201.53", "172.16.201.54", "172.16.201.56", "172.16.201.57",
    "172.16.201.58", "172.16.201.59", "172.16.201.60", "172.16.201.61", "172.16.201.62",
    "172.16.201.63", "172.16.201.64", "172.16.201.65", "172.16.201.67", "172.16.201.68",
    "172.16.201.121", "172.16.201.122", "172.16.201.142", "172.16.201.194", "172.16.201.198",
    "172.16.201.199"
]

    ip_addresses  = SRI


    ap_statuses = []
    for index, ip in enumerate(ip_addresses, start=1):
        logging.info(f'Round {index}')
        process_ip(ip, index, ap_statuses)
    save_log(ap_statuses)
    logging.info("Handle All Completed")

if __name__ == "__main__":
    time.sleep(2)
    threading.Thread(target=main).start()

