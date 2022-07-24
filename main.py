import platform
import customtkinter
import tkinter
from tkinter import messagebox, PhotoImage
import subprocess
import winreg
from pywinauto import *
import wmi
import psutil
import time

############################## RUN OPEN HARDWARE FIRST ############################

open_hardware = Application(backend="uia").start(r"OpenHardwareMonitor\OpenHardwareMonitor.exe", wait_for_idle=False)
dlg = open_hardware.window()
dlg.child_window(title="Minimize", control_type="Button")
dlg.minimize()
time.sleep(3)


############################### INSTANTIATE ########################################

customtkinter.set_appearance_mode("dark")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_appearance_mode("blue")  # Themes: "blue" (standard), "green", "dark-blue"

app = customtkinter.CTk()
app.geometry("580x450")
app.resizable(width=False, height=False)
app.title("Boost Manager")
icon = PhotoImage(file="processor-xxl.png")
app.iconphoto(True, icon)

############################################ GUI FUNCTIONS #####################################################


def on():
    messagebox.showinfo("Success!", "Boost is now enabled")
    subprocess.run("Powercfg -setacvalueindex scheme_current sub_processor PERFBOOSTMODE 1", shell=True)
    subprocess.run("Powercfg -setdcvalueindex scheme_current sub_processor PERFBOOSTMODE 1", shell=True)
    subprocess.run("Powercfg -setactive scheme_current", shell=True)


def off():
    messagebox.showinfo("Success!", "Boost is now disabled")
    subprocess.run("Powercfg -setacvalueindex scheme_current sub_processor PERFBOOSTMODE 0", shell=True)
    subprocess.run("Powercfg -setdcvalueindex scheme_current sub_processor PERFBOOSTMODE 0", shell=True)
    subprocess.run("Powercfg -setactive scheme_current", shell=True)


def set_value():
    try:
        # open and assign proper value in registry
        x = winreg.OpenKeyEx(winreg.HKEY_LOCAL_MACHINE,
                             r"SYSTEM\CurrentControlSet\Control\Power\PowerSettings\54533251-82be-4824-96c1-47b60b740d00\be337238-0d82-4146-a960-4f3749d470c7",
                             0, winreg.KEY_ALL_ACCESS)
        winreg.SetValueEx(x, "Attributes", 0, winreg.REG_SZ, "2")
        x.Close()
        # open power options directly to boost
        app2 = Application(backend="uia").start("control.exe powercfg.cpl,,3", wait_for_idle=False).connect(title="Power Options", timeout=100)

        main_window = app2.window()

        ppw = main_window.child_window(title="Processor power management", control_type="TreeItem")
        ppw.expand()
        ppbm = ppw.child_window(title="Processor performance boost mode", control_type="TreeItem")
        ppbm.expand()
        onbat = ppbm.child_window(title="On battery: Disabled", control_type="TreeItem").select()
    except:
        pass


def set_label():
    try:
        usage = psutil.cpu_percent()
        w = wmi.WMI(namespace="root\\OpenHardwareMonitor")
        sensors = w.Sensor()
        cpu_temps = []
        gpu_temp = 0
        for sensor in sensors:
            if sensor.SensorType == u'Temperature' and not 'GPU' in sensor.Name:
                cpu_temps += [float(sensor.Value)]
            elif sensor.SensorType == u'Temperature' and 'GPU' in sensor.Name:
                gpu_temp = sensor.Value
            try:
                num = 0
                length = len(cpu_temps)
                for val in cpu_temps:
                    num += val
                avg = num / length
                avg_cpu_temp = f"CPU: {float(avg):.2f} C°"
                gpu_temp_num = f"GPU: {float(gpu_temp):.2f} C°"
            except:
                pass
    except:
        messagebox.showwarning("Missing Service", "Open hardware not compatible!\nTemperature readings will be unavailable")
    try:
        if avg < 78:
            label_cpu_temp.configure(text=f"{avg_cpu_temp}\nUsage: {usage:.2f}%", text_color="#55CC80")
        else:
            label_cpu_temp.configure(text=f"{avg_cpu_temp}\nUsage: {usage:.2f}%", text_color="#f65353")
        if gpu_temp < 78:
            label_gpu_temp.configure(text=f"{gpu_temp_num}", text_color="#55CC80")
        else:
            label_gpu_temp.configure(text=gpu_temp_num, text_color="#f65353")
    except:
        messagebox.showwarning("Open Hardware Monitor", "Open Hardware Monitor not currently running")

    app.after(12000, set_label)


def start_open_hardware():
    try:
        open_hardware = Application(backend="uia").start(r"OpenHardwareMonitor\OpenHardwareMonitor.exe",
                                                    wait_for_idle=False)
        dlg = open_hardware.window()
        dlg.child_window(title="Minimize", control_type="Button")
        dlg.minimize()
    except:
        messagebox.showwarning("Open Hardware Monitor", "Error starting Open Hardware Monitor")

################################################## WIDGETS ##################################################


frame_1 = customtkinter.CTkFrame(master=app)
frame_1.pack(pady=20, padx=60, fill="both", expand=True)
info = f"processor: {platform.processor()}"
label = customtkinter.CTkLabel(master=frame_1, text=info).pack()

label_options = customtkinter.CTkLabel(master=frame_1, justify=tkinter.LEFT, text="BOOST OPTIONS")
label_options.pack(pady=12, padx=10)

label_cpu_temp = customtkinter.CTkLabel(master=frame_1, justify=tkinter.LEFT, text="")
label_cpu_temp.place(x=-10, y=80)

label_gpu_temp = customtkinter.CTkLabel(master=frame_1, justify=tkinter.LEFT, text="")
label_gpu_temp.place(x=16.5, y=130)

button_on = customtkinter.CTkButton(master=frame_1, command=on, text="ENABLE BOOST")
button_on.pack(pady=12, padx=10)

button_off = customtkinter.CTkButton(master=frame_1, command=off, text="DISABLE BOOST")
button_off.pack(pady=12, padx=10)

label_info = customtkinter.CTkLabel(master=frame_1, justify=tkinter.LEFT, text="       Enable manual control:  must run as admin\nIf application is functioning improperly:  run as admin")
label_info.pack(pady=12, padx=10)

button_manual = customtkinter.CTkButton(master=frame_1, command=set_value, text="Enable manual control")
button_manual.pack(pady=12, padx=10)

label_info2 = customtkinter.CTkLabel(master=frame_1, justify=tkinter.LEFT, text="If Open Hardware Monitor not currently running: Click button below")
label_info2.pack(pady=12, padx=10)

button_open_monitor = customtkinter.CTkButton(master=frame_1, command=start_open_hardware, text="Start Open Hardware Monitor")
button_open_monitor.pack(pady=12, padx=10)

############################################### END CALLS ########################################################

set_label()
app.mainloop()
open_hardware.kill()

