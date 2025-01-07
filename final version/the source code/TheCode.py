import os
import sys
import json
import random
import logging
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
import time
from datetime import datetime

# Copyright (C) 2024 Refter. All rights reserved.
# This software is licensed under the terms of the MIT License. See LICENSE.txt for more details.

class Config:
    DEFAULT_CONFIG = {
        "search_delay": (3, 7),
        "page_load_timeout": 30,
        "max_retries": 3,
        "save_log": True,
        "headless_mode": False
    }
    
    @classmethod
    def get_config_path(cls):
        if getattr(sys, 'frozen', False):
            application_path = os.path.dirname(sys.executable)
        else:
            application_path = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(application_path, 'config.json')

    @classmethod
    def load(cls):
        try:
            with open(cls.get_config_path(), 'r') as f:
                return {**cls.DEFAULT_CONFIG, **json.load(f)}
        except FileNotFoundError:
            return cls.DEFAULT_CONFIG
        except Exception as e:
            logging.error(f"Error loading config: {str(e)}")
            return cls.DEFAULT_CONFIG

    @classmethod
    def save(cls, config):
        try:
            with open(cls.get_config_path(), 'w') as f:
                json.dump(config, f, indent=4)
        except Exception as e:
            logging.error(f"Error saving config: {str(e)}")

def get_log_path():
    if getattr(sys, 'frozen', False):
        return os.path.join(os.path.dirname(sys.executable), 'logs')
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logs')

class BingSearchAutomation:
    def __init__(self, status_callback=None):
        self.config = Config.load()
        self.status_callback = status_callback
        self.setup_logging()

    def setup_logging(self):
        if self.config['save_log']:
            log_dir = get_log_path()
            os.makedirs(log_dir, exist_ok=True)
            log_filename = os.path.join(log_dir, f"search_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
            logging.basicConfig(
                filename=log_filename,
                level=logging.INFO,
                format='%(asctime)s - %(levelname)s - %(message)s'
            )

    def update_status(self, message):
        logging.info(message)
        if self.status_callback:
            self.status_callback(message)

    def setup_driver(self, driver_path):
        options = Options()
        if self.config['headless_mode']:
            options.add_argument('--headless')
        options.add_argument('--start-maximized')
        options.add_argument('--disable-extensions')
        options.page_load_strategy = 'eager'
        
        try:
            service = Service(driver_path)
            driver = webdriver.Edge(service=service, options=options)
            driver.set_page_load_timeout(self.config['page_load_timeout'])
            return driver
        except Exception as e:
            raise Exception(f"Failed to setup WebDriver: {str(e)}")

    def perform_search(self, search_file, driver_path):
        if not all([os.path.isfile(f) for f in [search_file, driver_path]]):
            raise FileNotFoundError("Search file or WebDriver not found")

        try:
            with open(search_file, 'r', encoding='utf-8') as f:
                searches = [line.strip() for line in f if line.strip()]
        except Exception as e:
            raise Exception(f"Error reading search file: {str(e)}")

        if not searches:
            raise ValueError("Search file is empty")

        driver = None
        completed_searches = 0
        
        try:
            driver = self.setup_driver(driver_path)
            
            for i, search_term in enumerate(searches, 1):
                retries = 0
                while retries < self.config['max_retries']:
                    try:
                        self.single_search(driver, search_term, i, len(searches))
                        completed_searches += 1
                        break
                    except TimeoutException:
                        retries += 1
                        self.update_status(f"Timeout on '{search_term}'. Retry {retries}/{self.config['max_retries']}")
                        if retries == self.config['max_retries']:
                            logging.error(f"Failed to complete search for '{search_term}' after {retries} retries")
                    except WebDriverException as e:
                        logging.error(f"WebDriver error: {str(e)}")
                        raise

        finally:
            if driver:
                try:
                    driver.quit()
                except:
                    pass
            self.update_status(f"Completed {completed_searches}/{len(searches)} searches")

    def single_search(self, driver, search_term, current, total):
        driver.get("https://www.bing.com")
        
        wait = WebDriverWait(driver, 10)
        search_box = wait.until(
            EC.presence_of_element_located((By.NAME, "q"))
        )
        
        search_box.clear()
        search_box.send_keys(search_term)
        search_box.send_keys(Keys.RETURN)
        
        delay = random.uniform(*self.config['search_delay'])
        self.update_status(f"Searching ({current}/{total}): {search_term}")
        time.sleep(delay)

class SearchAutomationGUI:
    def __init__(self):
        try:
            self.root = tk.Tk()
            self.root.title("Bing Search Automation")
            self.root.geometry("600x400")
            
            # Set style
            self.style = ttk.Style()
            self.style.configure("Accent.TButton", foreground="white", background="green")
            
            self.automation = BingSearchAutomation(self.update_status)
            self.setup_gui()
        except Exception as e:
            messagebox.showerror("Initialization Error", f"Error starting application: {str(e)}")
            if hasattr(self, 'root') and self.root:
                self.root.destroy()
            sys.exit(1)

    def setup_gui(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.create_file_selection(main_frame)
        self.create_config_frame(main_frame)
        self.create_status_frame(main_frame)
        self.create_copyright_label(main_frame)
        
        # Configure grid weights
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

    def create_file_selection(self, parent):
        files_frame = ttk.LabelFrame(parent, text="File Selection", padding="5")
        files_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=5)

        ttk.Label(files_frame, text="Search File:").grid(row=0, column=0, sticky="w")
        self.search_file_entry = ttk.Entry(files_frame, width=50)
        self.search_file_entry.grid(row=0, column=1, padx=5)
        ttk.Button(files_frame, text="Browse", command=lambda: self.browse_file(self.search_file_entry)).grid(row=0, column=2)

        ttk.Label(files_frame, text="WebDriver File:").grid(row=1, column=0, sticky="w")
        self.driver_file_entry = ttk.Entry(files_frame, width=50)
        self.driver_file_entry.grid(row=1, column=1, padx=5)
        ttk.Button(files_frame, text="Browse", command=lambda: self.browse_file(self.driver_file_entry)).grid(row=1, column=2)

    def create_config_frame(self, parent):
        config_frame = ttk.LabelFrame(parent, text="Settings", padding="5")
        config_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=5)

        self.headless_var = tk.BooleanVar(value=self.automation.config['headless_mode'])
        ttk.Checkbutton(config_frame, text="Headless Mode", variable=self.headless_var).grid(row=0, column=0)

        ttk.Button(config_frame, text="Start", command=self.start_search, style="Accent.TButton").grid(row=0, column=1, padx=5)

    def create_status_frame(self, parent):
        status_frame = ttk.LabelFrame(parent, text="Status", padding="5")
        status_frame.grid(row=2, column=0, columnspan=2, sticky="ew", pady=5)

        self.status_label = ttk.Label(status_frame, text="Ready")
        self.status_label.grid(row=0, column=0, sticky="w")

        self.progress = ttk.Progressbar(status_frame, mode='indeterminate')
        self.progress.grid(row=1, column=0, sticky="ew", pady=5)

    def create_copyright_label(self, parent):
        copyright_label = ttk.Label(parent, text="© 2024 Refter. All rights reserved.", font=("Helvetica", 8))
        copyright_label.grid(row=3, column=0, columnspan=2, sticky="s", pady=5)

    def browse_file(self, entry_field):
        file_path = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")])
        if file_path:
            entry_field.delete(0, tk.END)
            entry_field.insert(0, file_path)

    def update_status(self, message):
        self.status_label.config(text=message)
        self.root.update()

    def start_search(self):
        search_file = self.search_file_entry.get()
        driver_path = self.driver_file_entry.get()

        if not search_file or not driver_path:
            messagebox.showerror("Error", "Please select both the search file and WebDriver file")
            return

        self.automation.config['headless_mode'] = self.headless_var.get()
        Config.save(self.automation.config)

        try:
            self.progress.start()
            self.automation.perform_search(search_file, driver_path)
        except Exception as e:
            messagebox.showerror("Error", str(e))
            logging.error(f"Error during search: {str(e)}")
        finally:
            self.progress.stop()

    def run(self):
        try:
            self.root.mainloop()
        except Exception as e:
            messagebox.showerror("Error", f"Application error: {str(e)}")
            logging.error(f"Application error: {str(e)}")

if __name__ == "__main__":
    try:
        app = SearchAutomationGUI()
        app.run()
    except Exception as e:
        messagebox.showerror("Fatal Error", f"Fatal error: {str(e)}")
        logging.error(f"Fatal error: {str(e)}")
        sys.exit(1)
