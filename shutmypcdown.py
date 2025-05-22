import os
import sys
import json
import datetime
import tkinter as tk
from tkinter import messagebox, ttk
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger
import urllib.request
import winreg
import pystray
from pystray import MenuItem as item
from PIL import Image
import threading

try:
    from tkcalendar import DateEntry, Calendar
except ImportError:
    import subprocess

    subprocess.check_call([sys.executable, "-m", "pip", "install", "tkcalendar"])
    from tkcalendar import DateEntry, Calendar

ICON_URL = "https://github.com/Veixlix/shutmypcdown/blob/release/favicon.ico?raw=true"
ICON_PATH = os.path.join(os.getenv("TEMP"), "shutdown_icon.ico")
if not os.path.exists(ICON_PATH):
    try:
        urllib.request.urlretrieve(ICON_URL, ICON_PATH)
    except:
        pass

APPDATA = os.getenv("APPDATA")
APP_DIR = os.path.join(APPDATA, "ShutdownScheduler")
os.makedirs(APP_DIR, exist_ok=True)
LOG_FILE = os.path.join(APP_DIR, "shutdown_log.txt")
DATA_FILE = os.path.join(APP_DIR, "scheduled_shutdowns.json")

scheduler = BackgroundScheduler()
scheduler.start()
shutdown_jobs = {}


def log_event(event):
    with open(LOG_FILE, "a") as f:
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        f.write(f"[{timestamp}] {event}\n")


def save_jobs():
    data = {
        jid: {
            "datetime": entry["datetime"].strftime("%Y-%m-%d %H:%M"),
            "force": entry["force"],
            "recurrence": entry["recurrence"],
        }
        for jid, entry in shutdown_jobs.items()
    }
    with open(DATA_FILE, "w") as f:
        json.dump(data, f, indent=4)


def cancel_shutdown(job_id):
    if job_id in shutdown_jobs:
        try:
            shutdown_jobs[job_id]["job"].remove()
        except:
            pass
        del shutdown_jobs[job_id]
        save_jobs()
        log_event(f"Shutdown canceled for {job_id}")


def shutdown(force, job_id):
    log_event(f"Shutdown confirmed by user for job {job_id}.")
    os.system("shutdown /s /f /t 0" if force else "shutdown /s /t 0")


def show_shutdown_confirmation(force, job_id, refresh_callback):
    def on_confirm():
        recurrence = shutdown_jobs[job_id]["recurrence"]
        dt = shutdown_jobs[job_id]["datetime"]

        if recurrence == "daily":
            dt += datetime.timedelta(days=1)
        elif recurrence == "weekly":
            dt += datetime.timedelta(weeks=1)
        elif recurrence == "monthly":
            day = dt.day
            month = dt.month + 1
            year = dt.year
            if month > 12:
                month = 1
                year += 1
            while True:
                try:
                    dt = dt.replace(year=year, month=month, day=day)
                    break
                except ValueError:
                    day -= 1

        if recurrence != "once":
            cancel_shutdown(job_id)
            schedule_shutdown(dt, force, recurrence, refresh_callback)

        shutdown(force, job_id)
        popup.destroy()

    def on_cancel():
        cancel_shutdown(job_id)
        popup.destroy()
        messagebox.showinfo("Canceled", "Scheduled shutdown canceled.")
        refresh_callback()

    popup = tk.Tk()
    popup.title("Shutdown Confirmation")
    popup.geometry("300x120")
    popup.attributes("-topmost", True)
    tk.Label(popup, text="Do you want to shutdown now?", font=("Segoe UI", 10)).pack(
        pady=10
    )
    tk.Button(
        popup, text="Shutdown", command=on_confirm, width=10, bg="red", fg="white"
    ).pack(side="left", padx=20, pady=10)
    tk.Button(popup, text="Cancel", command=on_cancel, width=10).pack(
        side="right", padx=20, pady=10
    )
    popup.mainloop()


def schedule_shutdown(dt, force, recurrence, refresh_callback, jid=None):
    job_id = jid if jid else dt.strftime("%Y%m%d%H%M%S") + recurrence
    if job_id in shutdown_jobs:
        return
    now = datetime.datetime.now()
    if recurrence == "once" and dt <= now:
        return
    if recurrence == "once":
        job = scheduler.add_job(
            show_shutdown_confirmation,
            "date",
            run_date=dt,
            args=[force, job_id, refresh_callback],
        )
    else:
        cron_args = {"hour": dt.hour, "minute": dt.minute}
        if recurrence == "weekly":
            cron_args["day_of_week"] = dt.weekday()
        elif recurrence == "monthly":
            cron_args["day"] = dt.day
        job = scheduler.add_job(
            show_shutdown_confirmation,
            CronTrigger(**cron_args),
            args=[force, job_id, refresh_callback],
        )

    shutdown_jobs[job_id] = {
        "job": job,
        "datetime": dt,
        "force": force,
        "recurrence": recurrence,
    }
    save_jobs()
    log_event(f"Shutdown scheduled at {dt} (forceful={force}, recurrence={recurrence})")


def load_jobs(refresh_callback):
    if not os.path.exists(DATA_FILE):
        return
    try:
        with open(DATA_FILE, "r") as f:
            data = json.load(f)
        for jid, entry in data.items():
            dt = datetime.datetime.strptime(entry["datetime"], "%Y-%m-%d %H:%M")
            schedule_shutdown(
                dt, entry["force"], entry["recurrence"], refresh_callback, jid=jid
            )
    except Exception as e:
        print("Failed to load jobs:", e)


def cancel_all():
    for job_id in list(shutdown_jobs):
        cancel_shutdown(job_id)
    log_event("All shutdowns canceled")


def toggle_startup(enable):
    key = r"Software\Microsoft\Windows\CurrentVersion\Run"
    app_name = "ShutdownScheduler"
    exe_path = sys.executable
    try:
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER, key, 0, winreg.KEY_ALL_ACCESS
        ) as reg_key:
            if enable:
                winreg.SetValueEx(reg_key, app_name, 0, winreg.REG_SZ, exe_path)
            else:
                winreg.DeleteValue(reg_key, app_name)
    except Exception as e:
        messagebox.showerror("Startup Error", str(e))


def create_gui():
    root = tk.Tk()
    root.title("Shutdown Scheduler")
    root.geometry("700x500")
    root.configure(bg="#526fb7")
    if os.path.exists(ICON_PATH):
        try:
            root.iconbitmap(ICON_PATH)
        except:
            pass

    style = ttk.Style()
    style.theme_use("clam")
    style.configure(
        "TLabel", background="#526fb7", foreground="#ffe3a3", font=("Segoe UI", 10)
    )
    style.configure(
        "TButton",
        background="#ffb347",
        foreground="black",
        padding=6,
        font=("Segoe UI", 10, "bold"),
    )
    style.map(
        "TButton", foreground=[("active", "black")], background=[("active", "#ffd580")]
    )
    style.configure("TCheckbutton", background="#526fb7", foreground="#ffe3a3")
    style.configure(
        "TCombobox", fieldbackground="white", background="#ffb347", foreground="black"
    )
    style.configure("TFrame", background="#526fb7")

    top_frame = ttk.Frame(root)
    top_frame.pack(pady=10)

    date_entry = DateEntry(top_frame, width=12)
    date_entry.pack(side="left", padx=5)
    hour_var = tk.StringVar(value="22")
    minute_var = tk.StringVar(value="00")

    ttk.Combobox(
        top_frame,
        textvariable=hour_var,
        values=[f"{i:02d}" for i in range(24)],
        width=5,
    ).pack(side="left")
    ttk.Label(top_frame, text=":").pack(side="left")
    ttk.Combobox(
        top_frame,
        textvariable=minute_var,
        values=[f"{i:02d}" for i in range(60)],
        width=5,
    ).pack(side="left", padx=2)

    force_var = tk.BooleanVar()
    ttk.Checkbutton(top_frame, text="Force", variable=force_var).pack(
        side="left", padx=5
    )
    recurrence_var = tk.StringVar(value="once")
    ttk.Combobox(
        top_frame,
        textvariable=recurrence_var,
        values=["once", "daily", "weekly", "monthly"],
        width=10,
    ).pack(side="left", padx=5)

    startup_var = tk.StringVar(value="Disable")
    startup_menu = ttk.Combobox(
        top_frame,
        textvariable=startup_var,
        values=["Enable", "Disable"],
        width=10,
        state="readonly",
    )
    startup_menu.pack(side="right")
    ttk.Label(top_frame, text="Startup:").pack(side="right", padx=(0, 2))

    middle_frame = ttk.Frame(root)
    middle_frame.pack(fill="x", padx=10, pady=10)
    cal = Calendar(middle_frame, selectmode="day")
    cal.pack(fill="x")

    ttk.Label(root, text="Scheduled Shutdowns:").pack(pady=5)
    list_frame = ttk.Frame(root)
    list_frame.pack(fill="both", expand=True, padx=20)

    def refresh_schedule():
        for widget in list_frame.winfo_children():
            widget.destroy()
        for job_id, entry in shutdown_jobs.items():
            dt = entry["datetime"]
            text = f"\u23f0 {dt.strftime('%Y-%m-%d %H:%M')} - Force: {'Yes' if entry['force'] else 'No'} - {entry['recurrence'].capitalize()}"
            row = ttk.Frame(list_frame)
            row.pack(fill="x", pady=2)
            ttk.Label(row, text=text, width=60).pack(side="left")
            ttk.Button(
                row,
                text="Cancel",
                command=lambda jid=job_id: (cancel_shutdown(jid), refresh_schedule()),
            ).pack(side="right")

    def on_schedule():
        try:
            selected_date = date_entry.get_date()
            selected_hour = int(hour_var.get())
            selected_minute = int(minute_var.get())
            recurrence = recurrence_var.get()
            force = force_var.get()

            dt = datetime.datetime.combine(
                selected_date, datetime.time(selected_hour, selected_minute)
            )
            if recurrence == "once" and dt <= datetime.datetime.now():
                messagebox.showerror("Error", "Time must be in the future.")
                return

            schedule_shutdown(dt, force, recurrence, refresh_schedule)
            refresh_schedule()
            messagebox.showinfo(
                "Scheduled",
                f"Shutdown scheduled for {dt.strftime('%Y-%m-%d %H:%M')} ({recurrence})",
            )
        except Exception as e:
            messagebox.showerror("Error", f"Failed to schedule shutdown: {e}")

    def on_cancel_all():
        cancel_all()
        refresh_schedule()
        messagebox.showinfo("Canceled", "All shutdowns canceled.")

    def on_startup_change(event=None):
        toggle_startup(startup_var.get() == "Enable")

    def on_date_selected(event):
        selected = cal.selection_get()
        date_entry.set_date(selected)

    def minimize_to_tray():
        root.withdraw()

        def show():
            root.after(0, root.deiconify)

        image = Image.open(ICON_PATH)
        menu = (item("Show", show), item("Quit", lambda: (icon.stop(), root.quit())))
        icon = pystray.Icon("ShutdownScheduler", image, "Shutdown Scheduler", menu)
        threading.Thread(target=icon.run, daemon=True).start()

    ttk.Button(top_frame, text="Schedule", command=on_schedule).pack(
        side="left", padx=5
    )
    ttk.Button(top_frame, text="Cancel All", command=on_cancel_all).pack(
        side="left", padx=5
    )
    startup_menu.bind("<<ComboboxSelected>>", on_startup_change)
    cal.bind("<<CalendarSelected>>", on_date_selected)

    load_jobs(refresh_schedule)
    refresh_schedule()
    root.protocol("WM_DELETE_WINDOW", minimize_to_tray)
    root.mainloop()


if __name__ == "__main__":
    create_gui()
