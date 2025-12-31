import time
import threading
from datetime import datetime
import tkinter as tk
import winsound
import win32com.client

# -------- SETTINGS --------
BEEP_FREQ = 1500
BEEP_DURATION = 700
SNOOZE_MINUTES = 5
# --------------------------

alarm_active = False
task = ""
alarm_time = None

# -------- ALARM LOOP (THREAD-SAFE) --------
def alarm_loop():
    global alarm_active

    # Create voice INSIDE thread (important)
    speaker = win32com.client.Dispatch("SAPI.SpVoice")

    while alarm_active:
        # Loud beep
        winsound.Beep(BEEP_FREQ, BEEP_DURATION)

        # Async speak so it can repeat
        speaker.Speak(f"Time to {task}", 1)

        time.sleep(1)

# -------- BUTTONS --------
def dismiss_alarm():
    global alarm_active
    alarm_active = False
    status_label.config(text="Alarm dismissed")

def snooze_alarm():
    global alarm_active
    alarm_active = False
    status_label.config(text=f"Snoozed for {SNOOZE_MINUTES} minutes")
    root.after(SNOOZE_MINUTES * 60000, start_alarm)

def start_alarm():
    global alarm_active
    if not alarm_active:
        alarm_active = True
        status_label.config(text="🚨 ALARM RINGING 🚨")
        threading.Thread(target=alarm_loop, daemon=True).start()

# -------- TIME CHECK --------
def check_time():
    now = datetime.now()
    if now.hour == alarm_time.hour and now.minute == alarm_time.minute:
        start_alarm()
    else:
        root.after(1000, check_time)

# -------- INPUT --------
alarm_input = input("Enter time (HH:MM or HH:MM AM/PM): ").strip()
task = input("Enter your task: ").strip()

if "AM" not in alarm_input.upper() and "PM" not in alarm_input.upper():
    alarm_input += " AM"

alarm_time = datetime.strptime(alarm_input, "%I:%M %p")
print("⏰ Alarm set for", alarm_input, "| Task:", task)

# -------- GUI --------
root = tk.Tk()
root.title("Alarm System")
root.geometry("320x200")

tk.Label(root, text="Alarm System", font=("Arial", 16)).pack(pady=10)

status_label = tk.Label(root, text="Waiting for alarm...", font=("Arial", 12))
status_label.pack(pady=10)

frame = tk.Frame(root)
frame.pack(pady=20)

tk.Button(frame, text="Snooze", width=10, command=snooze_alarm).pack(side="left", padx=10)
tk.Button(frame, text="Dismiss", width=10, command=dismiss_alarm).pack(side="right", padx=10)

root.after(1000, check_time)
root.mainloop()
