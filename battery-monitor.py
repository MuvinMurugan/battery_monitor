import tkinter as tk
import win32com.client
import psutil
import time

battery = psutil.sensors_battery()

def speak(message):
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
    speaker.Speak(message)

def get_battery_percentage():
    battery = psutil.sensors_battery()
    return battery.percent

def show_battery_status():
    battery_percentage = get_battery_percentage()
    battery_status_label.config(text=f"Battery Percentage: {battery_percentage}%")
    if battery_percentage >= 85 and battery.power_plugged: # and battery.power_plugged
        message = f"Battery Percentage is at {battery_percentage}%. Please unplug the power cable."
        root.bell()
        speak(message)
    elif battery_percentage <= 25 and not battery.power_plugged: # and battery.power_plugged
        message = f"Battery Percentage is at {battery_percentage}%. Please plug the power cable."
        root.bell()
        speak(message)
    root.after(60000, show_battery_status) # Check the battery percentage every 60 seconds

root = tk.Tk()
root.title("Battery Percentage Monitor")
root.geometry("300x150")

battery_status_label = tk.Label(root, text="Battery Percentage: ", font=("Arial", 14))
battery_status_label.pack()

refresh_button = tk.Button(root, text="Start", font=("Arial", 14), command=show_battery_status)
refresh_button.pack()

root.mainloop()
