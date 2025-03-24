import subprocess
import asyncio
import tkinter as tk
from tkinter import messagebox, filedialog
from playwright.async_api import async_playwright
import threading
import os

DEFAULT_EDGE_PATH = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

def browse_edge():
    path = filedialog.askopenfilename(title="Select msedge.exe", filetypes=[("Edge Executable", "msedge.exe")])
    if path:
        edge_path_entry.delete(0, tk.END)
        edge_path_entry.insert(0, path)

def start_edge_and_playwright():
    edge_path = edge_path_entry.get()

    if not os.path.isfile(edge_path):
        messagebox.showerror("Error", "Invalid Edge path!")
        return

    try:
        # Start Edge with remote debugging enabled
        subprocess.Popen([edge_path, "--remote-debugging-port=9222"])
    except Exception as e:
        messagebox.showerror("Error launching Edge", str(e))
        return

    # Run playwright in a separate thread to not block the GUI
    threading.Thread(target=lambda: asyncio.run(run_playwright()), daemon=True).start()

async def run_playwright():
    try:
        async with async_playwright() as p:
            browser = await p.chromium.connect_over_cdp("http://localhost:9222")
            context = browser.contexts[0] if browser.contexts else await browser.new_context()
            page = await context.new_page()
            await page.goto("https://www.kaggle.com")
            await page.wait_for_timeout(5000)  # keep the tab open briefly
    except Exception as e:
        messagebox.showerror("Playwright Error", str(e))

# GUI setup
root = tk.Tk()
root.title("Edge + Playwright Launcher")
root.geometry("500x150")

tk.Label(root, text="Path to Edge Executable:").pack(pady=5)

edge_path_entry = tk.Entry(root, width=60)
edge_path_entry.insert(0, DEFAULT_EDGE_PATH)
edge_path_entry.pack(pady=5)

browse_button = tk.Button(root, text="Browse", command=browse_edge)
browse_button.pack(pady=5)

start_button = tk.Button(root, text="Start Edge & Playwright", command=start_edge_and_playwright)
start_button.pack(pady=10)

root.mainloop()
