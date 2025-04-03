import pandas as pd
import pyttsx3
import tkinter as tk
from tkinter import filedialog, messagebox
import time
import threading

# Initialize text-to-speech engine
engine = pyttsx3.init()

paused = False  # Global flag for pausing
stop_reading = False  # Global flag for stopping the reading

def select_file():
    """Opens a file dialog to select an Excel file and loads its columns."""
    global df
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
    
    if file_path:
        try:
            df = pd.read_excel(file_path)
            col_count = len(df.columns)
            col_dropdown["menu"].delete(0, "end")  # Clear existing options
            
            for i in range(col_count):
                col_dropdown["menu"].add_command(label=i, command=tk._setit(col_var, i))

            col_var.set(0)  # Default to first column
            messagebox.showinfo("Success", "File Loaded Successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file\n{str(e)}")

def read_aloud():
    """Reads the selected column numbers aloud with a user-defined delay."""
    global paused, stop_reading
    if df is None:
        messagebox.showerror("Error", "Please load an Excel file first!")
        return
    
    try:
        col_index = int(col_var.get())  # Get selected column index
        delay = float(delay_var.get())  # Get selected delay time
        stop_reading = False  # Reset stop flag

        for number in df.iloc[:, col_index]:
            if stop_reading:  
                break  # Stop reading when stop button is pressed
            
            while paused:
                time.sleep(0.1)  # Wait if paused
            
            text = str(number)
            print(text)
            engine.say(text)
            engine.runAndWait()
            time.sleep(delay)  # Apply delay
            
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read aloud\n{str(e)}")

def start_reading():
    """Runs the reading function in a separate thread."""
    threading.Thread(target=read_aloud, daemon=True).start()

def pause_resume():
    """Toggles the pause/resume state."""
    global paused
    paused = not paused  # Toggle pause state
    if paused:
        pause_btn.config(text="Resume")
    else:
        pause_btn.config(text="Pause")

def stop_speech():
    """Stops reading immediately."""
    global stop_reading
    stop_reading = True  # Set stop flag to terminate reading
    engine.stop()  # Stop current speech
    pause_btn.config(text="Pause")  # Reset pause button

# Initialize GUI
root = tk.Tk()
root.title("Excel Reader & Speech Converter")
root.geometry("400x400")

# UI Elements
tk.Label(root, text="Select an Excel File:").pack(pady=5)
tk.Button(root, text="Browse", command=select_file).pack()

tk.Label(root, text="Select Column Index:").pack(pady=5)
col_var = tk.StringVar(root)
col_dropdown = tk.OptionMenu(root, col_var, "No File Loaded")
col_dropdown.pack()

tk.Label(root, text="Set Delay (Seconds):").pack(pady=5)
delay_var = tk.DoubleVar(root, value=1.0)  # Default delay 1 second
delay_slider = tk.Scale(root, from_=0, to=5, resolution=0.5, orient="horizontal", variable=delay_var)
delay_slider.pack()

tk.Button(root, text="Start Reading", command=start_reading).pack(pady=5)
pause_btn = tk.Button(root, text="Pause", command=pause_resume)
pause_btn.pack(pady=5)
tk.Button(root, text="Stop", command=stop_speech).pack(pady=5)

# Start GUI
df = None  # Initialize dataframe
root.mainloop()
