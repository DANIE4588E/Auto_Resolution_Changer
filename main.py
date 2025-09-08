import win32api
import win32con
import psutil
import win32gui
import win32process
from screeninfo import get_monitors
import time
import threading
import tkinter as tk
from tkinter import messagebox

# Resolution management functions
def change_resolution_of_monitor(monitor_index, width, height, bits_per_pixel=32, frequency=60):
    device = win32api.EnumDisplayDevices(None, monitor_index)
    if not device:
        print(f"Monitor with index {monitor_index} not found.")
        return

    devmode = win32api.EnumDisplaySettings(device.DeviceName, win32con.ENUM_CURRENT_SETTINGS)
    devmode.PelsWidth = width
    devmode.PelsHeight = height
    devmode.BitsPerPel = bits_per_pixel
    devmode.DisplayFrequency = frequency
    devmode.Fields = (
        win32con.DM_PELSWIDTH
        | win32con.DM_PELSHEIGHT
        | win32con.DM_BITSPERPEL
        | win32con.DM_DISPLAYFREQUENCY
    )

    result = win32api.ChangeDisplaySettingsEx(device.DeviceName, devmode)

    if result == win32con.DISP_CHANGE_SUCCESSFUL:
        print(f"Resolution changed successfully for monitor {monitor_index}.")
    elif result == win32con.DISP_CHANGE_RESTART:
        print("System needs to restart for changes to take effect.")
    else:
        print(f"Failed to change resolution for monitor {monitor_index}. Error code: {result}")

def is_app_running(app_name):
    for process in psutil.process_iter(['name']):
        try:
            if process.info['name'] and app_name.lower() in process.info['name'].lower():
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
    return False

def get_window_for_process(process_name):
    for process in psutil.process_iter(['name']):
        if process.info['name'] and process_name.lower() in process.info['name'].lower():
            pid = process.pid
            def enum_window_callback(hwnd, result_list):
                try:
                    _, window_pid = win32process.GetWindowThreadProcessId(hwnd)
                    if pid == window_pid and win32gui.IsWindowVisible(hwnd):
                        rect = win32gui.GetWindowRect(hwnd)
                        if rect[0] != rect[2] and rect[1] != rect[3]:
                            result_list.append(hwnd)
                except Exception:
                    pass

            windows = []
            win32gui.EnumWindows(enum_window_callback, windows)
            if windows:
                return windows[0]
    return None

def get_monitor_for_window(hwnd):
    rect = win32gui.GetWindowRect(hwnd)
    left, top, right, bottom = rect
    center_x = (left + right) // 2
    center_y = (top + bottom) // 2
    monitors = get_monitors()

    for i, monitor in enumerate(monitors):
        if (
            monitor.x <= center_x < monitor.x + monitor.width and
            monitor.y <= center_y < monitor.y + monitor.height
        ):
            return i
    return None

def get_current_monitor_resolutions(index):
    monitors = get_monitors()
    return [monitors[index].width, monitors[index].height]

# GUI application
class ResolutionManagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Resolution Manager")

        self.running = False
        self.configurations = []

        # Load configurations from file
        self.load_configurations()

        # UI Elements
        tk.Label(root, text="Application Name:").grid(row=0, column=0, padx=5, pady=5)
        self.app_name_entry = tk.Entry(root)
        self.app_name_entry.grid(row=0, column=1, padx=5, pady=5)

        tk.Label(root, text="Monitor Index:").grid(row=1, column=0, padx=5, pady=5)
        self.monitor_index_entry = tk.Entry(root)
        self.monitor_index_entry.grid(row=1, column=1, padx=5, pady=5)

        tk.Label(root, text="Normal Resolution Width:").grid(row=2, column=0, padx=5, pady=5)
        self.normal_width_entry = tk.Entry(root)
        self.normal_width_entry.grid(row=2, column=1, padx=5, pady=5)

        tk.Label(root, text="Normal Resolution Height:").grid(row=3, column=0, padx=5, pady=5)
        self.normal_height_entry = tk.Entry(root)
        self.normal_height_entry.grid(row=3, column=1, padx=5, pady=5)

        tk.Label(root, text="Target Resolution Width:").grid(row=4, column=0, padx=5, pady=5)
        self.target_width_entry = tk.Entry(root)
        self.target_width_entry.grid(row=4, column=1, padx=5, pady=5)

        tk.Label(root, text="Target Resolution Height:").grid(row=5, column=0, padx=5, pady=5)
        self.target_height_entry = tk.Entry(root)
        self.target_height_entry.grid(row=5, column=1, padx=5, pady=5)

        self.add_button = tk.Button(root, text="Add Configuration", command=self.add_configuration)
        self.add_button.grid(row=6, column=0, columnspan=2, pady=10)

        self.start_button = tk.Button(root, text="Start", command=self.start_monitoring)
        self.start_button.grid(row=7, column=0, pady=10)

        self.stop_button = tk.Button(root, text="Stop", command=self.stop_monitoring)
        self.stop_button.grid(row=7, column=1, pady=10)

        self.save_button = tk.Button(root, text="Save Configurations", command=self.save_configurations)
        self.save_button.grid(row=8, column=0, columnspan=2, pady=10)

        self.config_listbox = tk.Listbox(root, width=70)
        self.config_listbox.grid(row=9, column=0, columnspan=2, pady=10)

        # Populate listbox with loaded configurations
        self.populate_listbox()

    def add_configuration(self):
        app_name = self.app_name_entry.get()
        monitor_index = self.monitor_index_entry.get()
        normal_width = self.normal_width_entry.get()
        normal_height = self.normal_height_entry.get()
        target_width = self.target_width_entry.get()
        target_height = self.target_height_entry.get()

        if app_name and monitor_index.isdigit() and normal_width.isdigit() and normal_height.isdigit() and target_width.isdigit() and target_height.isdigit():
            try:
                monitor_index = int(monitor_index)
                normal_res = [int(normal_width), int(normal_height)]
                target_res = [int(target_width), int(target_height)]

                # Check if app config already exists
                existing_config = next((config for config in self.configurations if config['app_name'] == app_name), None)

                if existing_config:
                    existing_config['monitor_resolutions'][monitor_index] = {
                        "normal_res": normal_res,
                        "target_res": target_res
                    }
                else:
                    self.configurations.append({
                        "app_name": app_name,
                        "monitor_resolutions": {
                            monitor_index: {
                                "normal_res": normal_res,
                                "target_res": target_res
                            }
                        }
                    })

                self.update_listbox()

                # Clear entries
                self.app_name_entry.delete(0, tk.END)
                self.monitor_index_entry.delete(0, tk.END)
                self.normal_width_entry.delete(0, tk.END)
                self.normal_height_entry.delete(0, tk.END)
                self.target_width_entry.delete(0, tk.END)
                self.target_height_entry.delete(0, tk.END)
            except ValueError:
                messagebox.showerror("Error", "Invalid resolution values.")
        else:
            messagebox.showerror("Error", "Please fill all fields.")

    def update_listbox(self):
        self.config_listbox.delete(0, tk.END)
        for config in self.configurations:
            app_name = config['app_name']
            for monitor_index, resolutions in config['monitor_resolutions'].items():
                normal_res = resolutions['normal_res']
                target_res = resolutions['target_res']
                self.config_listbox.insert(tk.END, f"{app_name} (Monitor {monitor_index}): Normal {normal_res}, Target {target_res}")

    def start_monitoring(self):
        if self.running:
            messagebox.showwarning("Warning", "Monitoring is already running.")
            return

        if not self.configurations:
            messagebox.showerror("Error", "No configurations to monitor.")
            return

        self.running = True
        self.monitor_thread = threading.Thread(target=self.monitor_apps)
        self.monitor_thread.start()

    def stop_monitoring(self):
        self.running = False
        if hasattr(self, 'monitor_thread') and self.monitor_thread.is_alive():
            self.monitor_thread.join()

    def save_configurations(self):
        with open("configurations.txt", "w") as f:
            for config in self.configurations:
                app_name = config['app_name']
                for monitor_index, resolutions in config['monitor_resolutions'].items():
                    normal_res = resolutions['normal_res']
                    target_res = resolutions['target_res']
                    f.write(f"{app_name},{monitor_index},{normal_res[0]}x{normal_res[1]},{target_res[0]}x{target_res[1]}\n")
        messagebox.showinfo("Info", "Configurations saved successfully.")

    def load_configurations(self):
        try:
            with open("configurations.txt", "r") as f:
                for line in f:
                    app_name, monitor_index, normal_res, target_res = line.strip().split(',')
                    monitor_index = int(monitor_index)
                    normal_res = list(map(int, normal_res.split('x')))
                    target_res = list(map(int, target_res.split('x')))

                    existing_config = next((config for config in self.configurations if config['app_name'] == app_name), None)

                    if existing_config:
                        existing_config['monitor_resolutions'][monitor_index] = {
                            "normal_res": normal_res,
                            "target_res": target_res
                        }
                    else:
                        self.configurations.append({
                            "app_name": app_name,
                            "monitor_resolutions": {
                                monitor_index: {
                                    "normal_res": normal_res,
                                    "target_res": target_res
                                }
                            }
                        })
        except FileNotFoundError:
            pass

    def populate_listbox(self):
        self.update_listbox()

    def monitor_apps(self):
        while self.running:
            for config in self.configurations:
                app_name = config['app_name']
                monitor_resolutions = config['monitor_resolutions']

                if is_app_running(app_name):
                    hwnd = get_window_for_process(app_name)
                    if hwnd:
                        monitor_index = get_monitor_for_window(hwnd)
                        if monitor_index is not None and monitor_index in monitor_resolutions:
                            resolutions = monitor_resolutions[monitor_index]
                            normal_res = resolutions['normal_res']
                            target_res = resolutions['target_res']

                            if get_current_monitor_resolutions(monitor_index) != target_res:
                                change_resolution_of_monitor(monitor_index, target_res[0], target_res[1])

                            while self.running:
                                try:
                                    current_index = get_monitor_for_window(hwnd)
                                    if current_index != monitor_index or not is_app_running(app_name):
                                        break
                                    time.sleep(2)
                                except Exception as e:
                                    print(f"Error: {e}")
                                    break

                            if get_current_monitor_resolutions(monitor_index) != normal_res:
                                change_resolution_of_monitor(monitor_index, normal_res[0], normal_res[1])
            time.sleep(5)

if __name__ == "__main__":
    root = tk.Tk()
    app = ResolutionManagerApp(root)
    root.mainloop()
