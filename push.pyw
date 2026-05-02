import sys, subprocess, tkinter as tk
from tkinter.scrolledtext import ScrolledText

import pathlib
SCRIPT = str(pathlib.Path(__file__).resolve().parent / "push.py")

root = tk.Tk()
root.title("Push to GitHub")
root.geometry("500x350")

log = ScrolledText(root, state="disabled", font=("Courier", 10))
log.pack(fill="both", expand=True, padx=10, pady=10)

def write(msg):
    log.config(state="normal")
    log.insert("end", msg + "\n")
    log.see("end")
    log.config(state="disabled")
    root.update()

def run():
    process = subprocess.Popen(
        [sys.executable, SCRIPT],
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True
    )
    for line in process.stdout:
        write(line.rstrip())
    write("\nWindow can be closed.")

root.after(500, run)
root.mainloop()
