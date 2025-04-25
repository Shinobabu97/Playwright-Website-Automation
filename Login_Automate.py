import subprocess
import asyncio
import tkinter as tk
from tkinter import scrolledtext
from playwright.async_api import async_playwright
import threading
import os
import psutil
import signal
import win32com.client as win32
import pythoncom
import time
import datetime

DEFAULT_EDGE_PATH = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
DEFAULT_DOWNLOAD_PATH = os.path.join(os.environ["USERPROFILE"], "Downloads")
RENAMED_FILE = "S24-Brückenmeldung DPD.xlsx"
RENAMED_FILE_2 = "S24-Brückenmeldung GLS.xlsx"
EMAIL_TO = "xyz@gmail.com"

# Log function
def log_message(message, error=False):
    timestamp = datetime.datetime.now().strftime("%H:%M:%S")
    log_text = f"[{timestamp}] {message}\n"
    
    # Use the main thread to update the UI
    if log_area:
        log_area.configure(state=tk.NORMAL)
        if error:
            log_area.insert(tk.END, log_text, "error")
        else:
            log_area.insert(tk.END, log_text)
        log_area.see(tk.END)  # Auto-scroll to the end
        log_area.configure(state=tk.DISABLED)

def open_edge_debug():
    #log_message("Starting Edge browser with debug port...")
    
    try:
        # Only close Edge instances with debugging port 9222
        close_debug_edge_instances()
        
        # Start a new Edge instance with debugging port
        subprocess.Popen([
            DEFAULT_EDGE_PATH,
            "--remote-debugging-port=9222",
            "--user-data-dir=C:\\temp\\edge_debug_profile",  # Use a separate profile to avoid conflicts
            "--start-maximized",
            "https://ssa.h.de/gui/#/menu/home//overview"
        ])
        
        # Give Edge time to start
        time.sleep(1)
        
        # Check if debugging port is active
        if not is_port_in_use(9222):
            log_message("Warning: Edge may not have started with debugging port. Check if it's running.", error=True)
        else:
            log_message("Edge erfolgreich gestartet.")
    except Exception as e:
        log_message(f"Error launching Edge: {str(e)}", error=True)

def is_port_in_use(port):
    """Check if the debugging port is in use"""
    import socket
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(('localhost', port)) == 0

def start_playwright():
    #log_message("Starting automation process...")
    threading.Thread(target=lambda: asyncio.run(run_playwright()), daemon=True).start()

# Playwright automation
async def run_playwright():
    full_save_path = os.path.join(DEFAULT_DOWNLOAD_PATH, RENAMED_FILE)

    try:
        if not is_port_in_use(9222):
            log_message("Debug port not available. Make sure Edge is running with --remote-debugging-port=9222", error=True)
            return
            
        async with async_playwright() as p:
            #log_message("Connecting to Edge browser...")
            browser = await p.chromium.connect_over_cdp("http://localhost:9222")
            context = browser.contexts[0] if browser.contexts else await browser.new_context(accept_downloads=True)
            page = None
            for ptab in context.pages:
                if "https://ssa.x.de/gui/#/menu/home//overview" in ptab.url:
                    page = ptab
                    break

            if not page:
                #log_message("Couldn't find the invoice page.", error=True)
                return
            # Email to DPD
            log_message("Automatisierungsprozess wird gestartet...")
            await page.goto("https://ssa.x.de/gui/#/menu/Dispatch/deliveriesCompleted/dlg/")
            await page.click('xpath=//*[@id="select-button"]')  # confirm date
            #log_message("Confirmed date selection...")
            await page.wait_for_timeout(1000)

            input_xpath = '//*[@id="deliverytable_filterrow_deliverygroup"]/cg-table-text-row-filter/cg-table-text-filter-input/label[1]/input'
            await page.locator(f'xpath={input_xpath}').fill('dpd')
            #log_message("Applied 'dpd' filter...")
            await page.wait_for_timeout(1000)

            await page.click('//*[@id="deliverytable_presets"]/span/div/span')
            await page.wait_for_timeout(1000)
            await page.click('//*[@id="deliverytable_presets"]/ul/div[1]/div[3]/a[1]/div[2]')
            #log_message("Selected preset filter...")
            await page.wait_for_timeout(3000)

            #log_message("Prozess erfolgreich abgeschlossen!")
            async with page.expect_download() as download_info:
                await page.click('//*[@id="deliverytable_export"]/span/span[2]')
                await page.click('//*[@id="deliverytable_export_xlsx"]')
            download = await download_info.value
            await download.save_as(full_save_path)
            #log_message(f"Datei gespeichert unter: {full_save_path}")
            await page.wait_for_timeout(2000)

            #log_message("Sending email with attachment...")
            send_email_with_attachment(full_save_path, "S24-Brückenmeldung DPD")
            log_message("E-Mail erfolgreich an DPD gesendet.")
            #log_message("Process completed successfully!", error=False)

            # Email to GLS
            full_save_path = os.path.join(DEFAULT_DOWNLOAD_PATH, RENAMED_FILE_2)
            #log_message("Connected to page. Starting automation...")
            await page.goto("https://ssax.de/gui/#/menu/Dispatch/deliveriesCompleted/dlg/")
            await page.click('xpath=//*[@id="select-button"]')  # confirm date
            #log_message("Confirmed date selection...")
            await page.wait_for_timeout(1000)

            input_xpath = '//*[@id="deliverytable_filterrow_deliverygroup"]/cg-table-text-row-filter/cg-table-text-filter-input/label[1]/input'
            await page.locator(f'xpath={input_xpath}').fill('gls')
            #log_message("Applied 'gls' filter...")
            await page.wait_for_timeout(1000)

            await page.click('//*[@id="deliverytable_presets"]/span/div/span')
            await page.wait_for_timeout(1000)
            await page.click('//*[@id="deliverytable_presets"]/ul/div[1]/div[3]/a[2]/div[2]')
            #log_message("Selected preset filter...")
            await page.wait_for_timeout(3000)

            #log_message("Datei wird heruntergeladen...")
            async with page.expect_download() as download_info:
                await page.click('//*[@id="deliverytable_export"]/span/span[2]')
                await page.click('//*[@id="deliverytable_export_xlsx"]')
            download = await download_info.value
            await download.save_as(full_save_path)
            #log_message(f"Datei gespeichert unter: {full_save_path}")
            await page.wait_for_timeout(2000)

            #log_message("Sending email with attachment...")
            send_email_with_attachment(full_save_path, "S24-Brückenmeldung GLS")
            log_message("E-Mail erfolgreich an GLS gesendet.")
            log_message("Prozess erfolgreich abgeschlossen!", error=False)

    except Exception as e:
        log_message(f"Automation error: {str(e)}", error=True)

# Close only Edge instances with debug port 9222
def close_debug_edge_instances():
    closed = False
    for process in psutil.process_iter(attrs=["pid", "name", "cmdline"]):
        try:
            if process.info["name"] and "msedge" in process.info["name"].lower():
                cmdline = " ".join(process.info["cmdline"]) if process.info["cmdline"] else ""
                if "--remote-debugging-port=9222" in cmdline:
                    os.kill(process.info["pid"], signal.SIGTERM)
                    closed = True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
    
    if closed:
        log_message("Vorhandene Edge-Debug-Instanzen geschlossen.")

# Email via Outlook (invisible)
def send_email_with_attachment(file_path, subject):
    try:
        pythoncom.CoInitialize()  # Initialize COM for this thread
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = EMAIL_TO
        mail.Subject = subject
        mail.Body = "siehe Anhang."
        mail.Attachments.Add(file_path)
        #mail.Display() #draft 
        #mail.Send()  # send with no GUI shown
        return True
    except Exception as e:
        log_message(f"Outlook Error: {str(e)}", error=True)
        return False
    finally:
        pythoncom.CoUninitialize()  # Cleanup COM for this thread

# GUI setup
root = tk.Tk()
root.title("S24 Brückenmeldung")
root.geometry("800x400")

# Create frame for buttons
button_frame = tk.Frame(root)
button_frame.pack(pady=10)

# Add buttons
tk.Button(button_frame, text="Edge für Anmeldung öffnen", command=open_edge_debug, 
          bg="#4CAF50", fg="white", padx=10, pady=5).pack(side=tk.LEFT, padx=10)
tk.Button(button_frame, text="Automatisierung starten & E-Mail senden", command=start_playwright,
          bg="#2196F3", fg="white", padx=10, pady=5).pack(side=tk.LEFT, padx=10)

# Create log display
log_frame = tk.LabelFrame(root, text="Execution Log")
log_frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

log_area = scrolledtext.ScrolledText(log_frame, state=tk.DISABLED, height=15)
log_area.pack(padx=5, pady=5, fill=tk.BOTH, expand=True)
log_area.tag_configure("error", foreground="red")

log_message("Anwendung gestartet. Bereit zur Ausführung!")

root.mainloop()
