import subprocess
import asyncio
import tkinter as tk
from tkinter import messagebox, filedialog
from playwright.async_api import async_playwright
import threading
import os
import psutil
import signal


DEFAULT_EDGE_PATH = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

def browse_edge():
    path = filedialog.askopenfilename(title="Select msedge.exe", filetypes=[("Edge Executable", "msedge.exe")])
    if path:
        edge_path_entry.delete(0, tk.END)
        edge_path_entry.insert(0, path)

def open_edge_debug():
    edge_path = edge_path_entry.get()
    if not os.path.isfile(edge_path):
        messagebox.showerror("Error", "Invalid Edge path!")
        return
    
    try:
        close_edge_instances() # Call the function to close all edge instances running on 9222
        subprocess.Popen([edge_path, "--remote-debugging-port=9222", "--start-maximized"]) #add "--headless" to open browser invisible
        #messagebox.showinfo("Info", "Edge launched in remote debugging mode. Log in and then run Playwright.")
    except Exception as e:
        messagebox.showerror("Error launching Edge", str(e))

def start_playwright():
    threading.Thread(target=lambda: asyncio.run(run_playwright()), daemon=True).start()

async def run_playwright():
    try:
        async with async_playwright() as p:
            browser = await p.chromium.connect_over_cdp("http://localhost:9222")
            context = browser.contexts[0] if browser.contexts else await browser.new_context() 
            #to set default download folder, add this to last braces (accept_downloads=True, downloads_path="C:\\Users\\YourUsername\\Downloads\\CustomFolder")
            page = await context.new_page()
            
            await page.goto("https://www.uni-due.de/studienangebote/studiengang.php?id=99")

            # Call the function to download file with your XPath and target folder
            r"""download_xpath = r'//*[@id="content__standard__main"]/div[1]/div/div/div/div/div/div[2]/div/p[5]/a'
            download_folder = r"C:\Users\shino\Desktop" 
            custom_filename = "my_custom_download"  # Optional
            await download_file(page, download_xpath, download_folder, custom_filename) #dont include the last paramter if you want to download it with default name
            """
       
            await page.close()
            await context.close()
            await browser.close()# this does not work as browser is externally via Chrome Dev Tools(CDT)
            close_edge_instances()
            

    except Exception as e:
        messagebox.showerror("Playwright Error", str(e))

#Function to trigger download to custom folder
async def download_file(page, download_xpath, download_folder, custom_filename=None):
    """ Clicks on a download link using the given XPath and saves the file to a custom location."""
    os.makedirs(download_folder, exist_ok=True)  # Ensure folder exists

    # Expect a download when clicking the link
    async with page.expect_download() as download_info:
        await page.locator(download_xpath).click(modifiers=["Alt"])#replace with click() if you just want to click without alt

    download = await download_info.value  # Get the download object

    # Keep original filename if custom_filename is not given
    if not custom_filename:
        custom_filename = download.suggested_filename
    else:
        # Preserve file extension
        _, ext = os.path.splitext(download.suggested_filename)
        custom_filename = f"{custom_filename}{ext}"

    # Define full path
    custom_download_path = os.path.join(download_folder, custom_filename)

    # Save the file to the new location
    await download.save_as(custom_download_path)


#Funtion to close all edge browsers running on 9222
def close_edge_instances():
    """
    Finds and terminates all instances of Microsoft Edge running on port 9222.
    """
    for process in psutil.process_iter(attrs=["pid", "name", "cmdline"]):
        try:
            # Check if it's Microsoft Edge running with remote debugging
            if process.info["name"] and "msedge" in process.info["name"].lower():
                cmdline = " ".join(process.info["cmdline"]) if process.info["cmdline"] else ""
                
                if "--remote-debugging-port=9222" in cmdline:
                    print(f"Closing Edge instance (PID: {process.info['pid']})")
                    os.kill(process.info["pid"], signal.SIGTERM)  # SIGTERM (graceful shutdown)
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass




######Main Program


# GUI setup
root = tk.Tk()
root.title("Edge + Playwright Launcher")
root.geometry("500x200")

tk.Label(root, text="Path to Edge Executable:").pack(pady=5)
edge_path_entry = tk.Entry(root, width=60)
edge_path_entry.insert(0, DEFAULT_EDGE_PATH)
edge_path_entry.pack(pady=5)

browse_button = tk.Button(root, text="Browse", command=browse_edge)
browse_button.pack(pady=5)

debug_button = tk.Button(root, text="Open Edge for Login", command=open_edge_debug)
debug_button.pack(pady=5)

start_button = tk.Button(root, text="Run Playwright", command=start_playwright)
start_button.pack(pady=10)

root.mainloop()
