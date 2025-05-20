# 💻 Shutdown Scheduler - shutmypcdown

 Easily schedule or cancel system shutdowns using an calendar interface. Hover over marked days to see scheduled shutdown times
 Made on request to give absolute time/date rather than seconds/minutes/hours

---

##  Features

- 📅 Select shutdown date & time with a themed GUI
- 📌 Calendar view highlights scheduled shutdowns
- 🖱️ Hover over a calendar day to see the exact shutdown time
- ✅ Optional force shutdown toggle

---

## 📥 Download

**Grab the latest executable release from:**  
👉 [releases](https://github.com/Veixlix/shutmypcdown/releases)

No Python installation needed — just download and run.

---

## Build from source

### Requirements
- Python 3.9+
- `tkinter`, `tkcalendar`, `apscheduler`, `plyer`, `winsound`
```bash
pip install tkcalendar apscheduler plyer
python shutmypcdown.py
