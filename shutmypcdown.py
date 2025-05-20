import os
import sys
import threading
import datetime
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from plyer import notification
import winsound
from apscheduler.schedulers.background import BackgroundScheduler
try:
    from tkcalendar import DateEntry
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "tkcalendar"])
    from tkcalendar import DateEntry

try:
    import pystray
    from PIL import Image, ImageDraw
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pystray", "pillow"])
    import pystray
    from PIL import Image, ImageDraw

LOG_FILE = "shutdown_log.txt"

scheduler = BackgroundScheduler()
scheduler.start()

shutdown_job = None
tray_icon = None
root = None

def log_event(event):
    with open(LOG_FILE, "a") as f:
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        f.write(f"[{timestamp}] {event}\n")

def shutdown(force):
    log_event("Shutdown initiated")
    if force:
        os.system("shutdown /s /f /t 0")
    else:
        os.system("shutdown /s /t 0")

def play_notification():
    winsound.MessageBeep(winsound.MB_ICONEXCLAMATION)
    notification.notify(
        title='Shutdown Notice',
        message='System will shut down in 1 minute.',
        timeout=10
    )
    log_event("1-minute shutdown warning issued")

def schedule_shutdown(dt, force):
    global shutdown_job
    cancel_shutdown() 

    warning_time = dt - datetime.timedelta(minutes=1)

    scheduler.add_job(play_notification, 'date', run_date=warning_time)
    shutdown_job = scheduler.add_job(shutdown, 'date', run_date=dt, args=[force])

    log_event(f"Shutdown scheduled at {dt} (forceful={force})")

def cancel_shutdown():
    global shutdown_job
    if shutdown_job:
        shutdown_job.remove()
        shutdown_job = None
        log_event("Shutdown canceled")

def create_image():
    image = Image.new('RGB', (64, 64), "white")
    draw = ImageDraw.Draw(image)
    draw.rectangle((16, 16, 48, 48), fill="black")
    return image

def show_tray():
    global tray_icon

    def on_open():
        root.after(0, root.deiconify)

    def on_exit():
        cancel_shutdown()
        tray_icon.stop()
        scheduler.shutdown()
        os._exit(0)

    tray_icon = pystray.Icon("Shutdown Scheduler")
    tray_icon.icon = create_image()
    tray_icon.menu = pystray.Menu(
        pystray.MenuItem("Open Scheduler", lambda: on_open()),
        pystray.MenuItem("Exit", lambda: on_exit())
    )
    tray_icon.run()

def create_gui():
    global root

    def on_schedule():
        try:
            selected_date = date_entry.get_date()
            selected_hour = int(hour_var.get())
            selected_minute = int(minute_var.get())
            force = force_var.get()

            dt = datetime.datetime.combine(selected_date, datetime.time(selected_hour, selected_minute))
            if dt <= datetime.datetime.now():
                messagebox.showerror("Error", "Time must be in the future.")
                return

            schedule_shutdown(dt, force)
            messagebox.showinfo("Scheduled", f"Shutdown scheduled at {dt}.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def on_cancel():
        cancel_shutdown()
        messagebox.showinfo("Canceled", "Shutdown canceled.")

    root = tk.Tk()
    root.title("Shutdown Scheduler")
    root.geometry("300x280")

    ttk.Label(root, text="Select Date:").pack(pady=5)
    date_entry = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2)
    date_entry.pack()

    ttk.Label(root, text="Select Time (24h):").pack(pady=5)
    time_frame = ttk.Frame(root)
    time_frame.pack()

    hour_var = tk.StringVar(value="12")
    minute_var = tk.StringVar(value="00")

    hour_menu = ttk.Combobox(time_frame, textvariable=hour_var, values=[f"{i:02d}" for i in range(24)], width=5)
    hour_menu.pack(side="left", padx=2)
    ttk.Label(time_frame, text=":").pack(side="left")
    minute_menu = ttk.Combobox(time_frame, textvariable=minute_var, values=[f"{i:02d}" for i in range(60)], width=5)
    minute_menu.pack(side="left", padx=2)

    force_var = tk.BooleanVar()
    force_check = ttk.Checkbutton(root, text="Force shutdown", variable=force_var)
    force_check.pack(pady=5)

    ttk.Button(root, text="Schedule Shutdown", command=on_schedule).pack(pady=5)
    ttk.Button(root, text="Cancel Shutdown", command=on_cancel).pack(pady=5)

    def minimize_to_tray():
        root.withdraw()
        threading.Thread(target=show_tray, daemon=True).start()

    root.protocol("WM_DELETE_WINDOW", minimize_to_tray)
    root.mainloop()

if __name__ == "__main__":
    create_gui()
