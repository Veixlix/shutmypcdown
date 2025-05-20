import os
import sys
import datetime
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from plyer import notification
import winsound
from apscheduler.schedulers.background import BackgroundScheduler

try:
    from tkcalendar import DateEntry, Calendar
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "tkcalendar"])
    from tkcalendar import DateEntry, Calendar

import urllib.request

ICON_URL = "https://github.com/Veixlix/shutmypcdown/blob/release/favicon.ico?raw=true"
ICON_PATH = "shutdown_icon.ico"

if not os.path.exists(ICON_PATH):
    try:
        urllib.request.urlretrieve(ICON_URL, ICON_PATH)
    except Exception as e:
        print("Warning: Could not download icon:", e)

LOG_FILE = "shutdown_log.txt"

scheduler = BackgroundScheduler()
scheduler.start()

shutdown_job = None
shutdown_time = None

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
    global shutdown_job, shutdown_time
    cancel_shutdown()

    warning_time = dt - datetime.timedelta(minutes=1)

    scheduler.add_job(play_notification, 'date', run_date=warning_time)
    shutdown_job = scheduler.add_job(shutdown, 'date', run_date=dt, args=[force])

    shutdown_time = dt
    log_event(f"Shutdown scheduled at {dt} (forceful={force})")

def cancel_shutdown():
    global shutdown_job, shutdown_time
    if shutdown_job:
        shutdown_job.remove()
        shutdown_job = None
        shutdown_time = None
        log_event("Shutdown canceled")

def create_gui():
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
            refresh_calendar()
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def on_cancel():
        cancel_shutdown()
        messagebox.showinfo("Canceled", "Shutdown canceled.")
        refresh_calendar()

    def refresh_calendar():
        calendar.calevent_remove('all')
        if shutdown_time:
            shutdown_str = shutdown_time.strftime("Shutdown at %H:%M")
            calendar.calevent_create(shutdown_time.date(), shutdown_str, 'shutdown')
            calendar.tag_config('shutdown', background='orange', foreground='white')

    def on_date_hover(event):
        tooltip_label.place_forget()
        hovered_date = calendar.selection_get()
        if shutdown_time and hovered_date == shutdown_time.date():
            tooltip_label.config(text=f"Shutdown planned at {shutdown_time.strftime('%H:%M')}")
            tooltip_label.place(x=event.x_root - root.winfo_rootx(), y=event.y_root - root.winfo_rooty() + 20)

    def on_hover_leave(event):
        tooltip_label.place_forget()

    def on_calendar_select(event):
        selected = calendar.selection_get()
        if selected:
            date_entry.set_date(selected)

    root = tk.Tk()
    root.title("Shutdown Scheduler")
    root.geometry("520x360")
    root.resizable(False, False)

    if os.path.exists(ICON_PATH):
        try:
            root.iconbitmap(ICON_PATH)
        except Exception as e:
            print("Warning: Could not set icon:", e)

    root.configure(bg="#526fb7")
    style = ttk.Style()
    style.theme_use("clam")
    style.configure("TLabel", background="#526fb7", foreground="#ffe3a3", font=("Segoe UI", 10))
    style.configure("TButton", background="#ffb347", foreground="black", padding=6, font=("Segoe UI", 10, "bold"))
    style.configure("TCheckbutton", background="#526fb7", foreground="#ffe3a3")
    style.configure("TFrame", background="#526fb7")
    style.configure("TCombobox", fieldbackground="white", background="white", foreground="black")

    left_frame = ttk.Frame(root)
    left_frame.pack(side="left", fill="both", expand=True, padx=10, pady=10)

    right_frame = ttk.Frame(root)
    right_frame.pack(side="right", fill="y", padx=10, pady=10)

    ttk.Label(left_frame, text="Select Date:").pack(pady=5)
    date_entry = DateEntry(left_frame, width=12, background='#526fb7', foreground='black', borderwidth=2)
    date_entry.pack()

    ttk.Label(left_frame, text="Select Time (24h):").pack(pady=5)
    time_frame = ttk.Frame(left_frame)
    time_frame.pack()

    hour_var = tk.StringVar(value="12")
    minute_var = tk.StringVar(value="00")

    hour_menu = ttk.Combobox(time_frame, textvariable=hour_var, values=[f"{i:02d}" for i in range(24)], width=5)
    hour_menu.pack(side="left", padx=2)
    ttk.Label(time_frame, text=":").pack(side="left")
    minute_menu = ttk.Combobox(time_frame, textvariable=minute_var, values=[f"{i:02d}" for i in range(60)], width=5)
    minute_menu.pack(side="left", padx=2)

    force_var = tk.BooleanVar()
    force_check = ttk.Checkbutton(left_frame, text="Force shutdown", variable=force_var)
    force_check.pack(pady=5)

    ttk.Button(left_frame, text="Schedule Shutdown", command=on_schedule).pack(pady=5)
    ttk.Button(left_frame, text="Cancel Shutdown", command=on_cancel).pack(pady=5)

    calendar = Calendar(right_frame, selectmode='day', background="white", foreground="black")
    calendar.pack(fill="both", expand=True)
    calendar.bind("<Motion>", on_date_hover)
    calendar.bind("<Leave>", on_hover_leave)
    calendar.bind("<<CalendarSelected>>", on_calendar_select)

    tooltip_label = tk.Label(root, text="", bg="#fffcbb", relief="solid", borderwidth=1, font=("Segoe UI", 9))
    refresh_calendar()

    def on_close():
        cancel_shutdown()
        scheduler.shutdown()
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_close)
    root.mainloop()

if __name__ == "__main__":
    create_gui()
