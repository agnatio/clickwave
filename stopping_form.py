import tkinter as tk
from tkinter import ttk
import os

def invoke_stopping_form(title, seconds=3):
    seconds = int(seconds)  # Convert seconds to an integer
    timer_running = True

    def update_timer():
        nonlocal seconds, timer_running
        if timer_running and seconds > 0:
            go_button['text'] = f'Press or wait {seconds} seconds'
            seconds -= 1
            root.after(1000, update_timer)
        elif not timer_running:
            go_button['text'] = 'Press to continue'
        else:
            root.destroy()

    def cancel():
        root.destroy()
        exit()

    def pause():
        nonlocal timer_running
        timer_running = not timer_running

    root = tk.Tk()
    root.title(title)
    icon = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'fox.ico')
    root.iconbitmap(icon)

    style = ttk.Style()
    style.theme_use('clam')

    title_label = ttk.Label(root, text=title)
    title_label.pack()

    go_button = ttk.Button(root, text='Go', command=root.destroy)
    go_button.pack(side=tk.LEFT, padx=5, pady=5)

    pause_button = ttk.Button(root, text='Pause', command=pause)
    pause_button.pack(side=tk.LEFT, padx=5, pady=5)

    cancel_button = ttk.Button(root, text='Cancel', command=cancel)
    cancel_button.pack(side=tk.LEFT, padx=5, pady=5)

    root.update_idletasks()  # Ensure window dimensions are updated

    window_width = root.winfo_width()
    window_height = root.winfo_height()
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2

    root.geometry(f"+{x}+{y}")  # Place the window at the center of the screen

    update_timer()  # Initial call to start the timer

    root.mainloop()


if __name__ == '__main__':
    invoke_stopping_form("Stopping Form", seconds=10)
