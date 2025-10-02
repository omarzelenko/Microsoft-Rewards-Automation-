import os
import sys
import json
import random
import logging
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import customtkinter as ctk
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
import threading

# Copyright (C) 2024 Refter. All rights reserved.
# This software is licensed under the terms of the MIT License. See LICENSE.txt for more details.

# Set the appearance mode and color theme
ctk.set_appearance_mode("System")  # "System", "Dark" or "Light"
ctk.set_default_color_theme("blue")  # "blue", "green", "dark-blue"

class Config:
    DEFAULT_CONFIG = {
        "search_delay": (3, 7),
        "page_load_timeout": 30,
        "max_retries": 3,
        "save_log": True,
        "headless_mode": False,
        "appearance_mode": "System",
		"color_theme": "blue",
		# Auto search settings
		"auto_search_enabled": False,
		"auto_search_interval_hours": 24,
		"last_auto_search": None,
		"run_on_startup": False,
		# Remember last used paths
		"last_search_file": "",
		"last_driver_path": ""
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

def get_resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, 'resources', relative_path)

class BingSearchAutomation:
    def __init__(self, status_callback=None, progress_callback=None):
        self.config = Config.load()
        self.status_callback = status_callback
        self.progress_callback = progress_callback
        self.setup_logging()
        self.is_running = False
        self.search_thread = None

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

    def update_progress(self, current, total):
        if self.progress_callback:
            self.progress_callback(current, total)

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
        if not os.path.isfile(search_file):
            raise FileNotFoundError(f"Search file not found: {search_file}")
        
        if not os.path.isfile(driver_path):
            raise FileNotFoundError(f"WebDriver not found: {driver_path}")

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
                if not self.is_running:
                    break
                
                retries = 0
                while retries < self.config['max_retries'] and self.is_running:
                    try:
                        self.single_search(driver, search_term, i, len(searches))
                        completed_searches += 1
                        self.update_progress(i, len(searches))
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
            self.is_running = False
            # Record timestamp of last auto search
            try:
                updated_config = {**self.config, "last_auto_search": datetime.now().isoformat()}
                Config.save(updated_config)
                self.config = updated_config
            except Exception as e:
                logging.error(f"Failed to update last_auto_search: {str(e)}")

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

    def start_search_thread(self, search_file, driver_path):
        if self.search_thread and self.search_thread.is_alive():
            return  # Already running
        
        self.is_running = True
        self.search_thread = threading.Thread(
            target=self.perform_search, 
            args=(search_file, driver_path)
        )
        self.search_thread.daemon = True
        self.search_thread.start()
    
    def stop_search(self):
        self.is_running = False
        self.update_status("Stopping search operations...")

class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip_window = None
        self.widget.bind("<Enter>", self.show_tip)
        self.widget.bind("<Leave>", self.hide_tip)

    def show_tip(self, event=None):
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 25

        self.tip_window = tk.Toplevel(self.widget)
        self.tip_window.wm_overrideredirect(True)
        self.tip_window.wm_geometry(f"+{x}+{y}")
        
        label = tk.Label(self.tip_window, text=self.text, background="#ffffe0", relief="solid", borderwidth=1, padx=5, pady=2)
        label.pack()

    def hide_tip(self, event=None):
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None

class SearchAutomationGUI:
    def __init__(self):
        try:
            self.config = Config.load()
            
            # Set the appearance mode from config
            ctk.set_appearance_mode(self.config.get("appearance_mode", "System"))
            ctk.set_default_color_theme(self.config.get("color_theme", "blue"))
            
            self.root = ctk.CTk()
            self.root.title("Bing Search Automation")
            self.root.geometry("800x600")
            self.root.minsize(700, 500)
            
            self.setup_variables()
            self.automation = BingSearchAutomation(
                self.update_status, 
                self.update_progress
            )
            self.setup_gui()
            # Attempt auto-run if enabled and due
            self.maybe_auto_run()
            
        except Exception as e:
            messagebox.showerror("Initialization Error", f"Error starting application: {str(e)}")
            if hasattr(self, 'root') and self.root:
                self.root.destroy()
            sys.exit(1)

    def setup_variables(self):
        self.search_file_var = tk.StringVar(value=self.config.get('last_search_file', ''))
        self.driver_file_var = tk.StringVar(value=self.config.get('last_driver_path', ''))
        self.headless_var = tk.BooleanVar(value=self.config.get('headless_mode', False))
        self.search_delay_min = tk.DoubleVar(value=self.config.get('search_delay', [3, 7])[0])
        self.search_delay_max = tk.DoubleVar(value=self.config.get('search_delay', [3, 7])[1])
        self.page_load_timeout = tk.IntVar(value=self.config.get('page_load_timeout', 30))
        self.max_retries = tk.IntVar(value=self.config.get('max_retries', 3))
        self.save_log_var = tk.BooleanVar(value=self.config.get('save_log', True))
        self.appearance_var = tk.StringVar(value=self.config.get('appearance_mode', 'System'))
        # Auto search variables
        self.auto_enabled_var = tk.BooleanVar(value=self.config.get('auto_search_enabled', False))
        self.auto_interval_hours = tk.IntVar(value=self.config.get('auto_search_interval_hours', 24))
        self.run_on_startup_var = tk.BooleanVar(value=self.config.get('run_on_startup', False))

    def setup_gui(self):
        # Create main container
        self.main_frame = ctk.CTkFrame(self.root)
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Header frame with logo and title
        self.create_header_frame()
        
        # Tabview for different sections
        self.create_tabview()
        
        # Status frame
        self.create_status_frame()
        
        # Footer with copyright
        self.create_footer()

    def create_header_frame(self):
        header_frame = ctk.CTkFrame(self.main_frame)
        header_frame.pack(fill="x", pady=(0, 20))
        
        # Try to load a logo
        try:
            # Path to the app icon - create a resources directory and place an icon file there
            logo_path = get_resource_path("logo.png")
            if os.path.exists(logo_path):
                logo_img = Image.open(logo_path)
                logo_img = logo_img.resize((64, 64))
                logo_photo = ImageTk.PhotoImage(logo_img)
                logo_label = ctk.CTkLabel(header_frame, image=logo_photo, text="")
                logo_label.image = logo_photo  # Keep a reference
                logo_label.pack(side="left", padx=10)
        except Exception:
            pass  # If logo loading fails, continue without it
        
        title_label = ctk.CTkLabel(header_frame, text="Bing Search Automation", font=("Helvetica", 22, "bold"))
        title_label.pack(side="left", padx=10)
        
        # Appearance mode toggle
        appearance_frame = ctk.CTkFrame(header_frame)
        appearance_frame.pack(side="right", padx=10)
        
        ctk.CTkLabel(appearance_frame, text="Theme:").pack(side="left", padx=5)
        appearance_menu = ctk.CTkOptionMenu(
            appearance_frame, 
            values=["System", "Light", "Dark"],
            variable=self.appearance_var,
            command=self.change_appearance_mode
        )
        appearance_menu.pack(side="left")

    def create_tabview(self):
        self.tabview = ctk.CTkTabview(self.main_frame)
        self.tabview.pack(fill="both", expand=True)
        
        # Create tabs
        self.tab_search = self.tabview.add("Search")
        self.tab_settings = self.tabview.add("Settings")
        self.tab_about = self.tabview.add("About")
        
        # Fill the search tab
        self.create_search_tab()
        
        # Fill the settings tab
        self.create_settings_tab()
        
        # Fill the about tab
        self.create_about_tab()

    def create_search_tab(self):
        # File selection frame
        files_frame = ctk.CTkFrame(self.tab_search)
        files_frame.pack(fill="x", padx=20, pady=20)
        
        # Search file row
        search_frame = ctk.CTkFrame(files_frame)
        search_frame.pack(fill="x", pady=10)
        
        ctk.CTkLabel(search_frame, text="Search File:", width=100).pack(side="left")
        search_entry = ctk.CTkEntry(search_frame, textvariable=self.search_file_var, width=400)
        search_entry.pack(side="left", padx=10, fill="x", expand=True)
        
        search_button = ctk.CTkButton(
            search_frame, 
            text="Browse", 
            command=lambda: self.browse_file(self.search_file_var, "Select Search File")
        )
        search_button.pack(side="right")
        
        # WebDriver file row
        driver_frame = ctk.CTkFrame(files_frame)
        driver_frame.pack(fill="x", pady=10)
        
        ctk.CTkLabel(driver_frame, text="WebDriver:", width=100).pack(side="left")
        driver_entry = ctk.CTkEntry(driver_frame, textvariable=self.driver_file_var, width=400)
        driver_entry.pack(side="left", padx=10, fill="x", expand=True)
        
        driver_button = ctk.CTkButton(
            driver_frame, 
            text="Browse", 
            command=lambda: self.browse_file(self.driver_file_var, "Select WebDriver File")
        )
        driver_button.pack(side="right")
        
        # Headless mode toggle
        headless_frame = ctk.CTkFrame(files_frame)
        headless_frame.pack(fill="x", pady=10)
        
        headless_switch = ctk.CTkSwitch(
            headless_frame, 
            text="Headless Mode (Run in background)", 
            variable=self.headless_var,
            onvalue=True,
            offvalue=False
        )
        headless_switch.pack(side="left")
        ToolTip(headless_switch, "Run browser in background without UI")
        
        # Control buttons frame
        control_frame = ctk.CTkFrame(self.tab_search)
        control_frame.pack(fill="x", padx=20, pady=10)
        
        self.start_button = ctk.CTkButton(
            control_frame, 
            text="Start Searching", 
            command=self.start_search,
            fg_color="#28a745",
            hover_color="#218838",
            font=("Helvetica", 14, "bold")
        )
        self.start_button.pack(side="left", padx=10, expand=True, fill="x")
        
        self.stop_button = ctk.CTkButton(
            control_frame, 
            text="Stop", 
            command=self.stop_search,
            fg_color="#dc3545",
            hover_color="#c82333",
            state="disabled",
            font=("Helvetica", 14, "bold")
        )
        self.stop_button.pack(side="right", padx=10, expand=True, fill="x")

    def create_settings_tab(self):
        settings_frame = ctk.CTkFrame(self.tab_settings)
        settings_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Search delay settings
        delay_frame = ctk.CTkFrame(settings_frame)
        delay_frame.pack(fill="x", pady=10)
        
        ctk.CTkLabel(delay_frame, text="Search Delay (seconds):", anchor="w").pack(anchor="w")
        
        delay_range_frame = ctk.CTkFrame(delay_frame)
        delay_range_frame.pack(fill="x", pady=5)
        
        ctk.CTkLabel(delay_range_frame, text="Min:").pack(side="left")
        min_delay = ctk.CTkEntry(delay_range_frame, textvariable=self.search_delay_min, width=80)
        min_delay.pack(side="left", padx=10)
        
        ctk.CTkLabel(delay_range_frame, text="Max:").pack(side="left", padx=(20, 0))
        max_delay = ctk.CTkEntry(delay_range_frame, textvariable=self.search_delay_max, width=80)
        max_delay.pack(side="left", padx=10)
        
        # Page load timeout
        timeout_frame = ctk.CTkFrame(settings_frame)
        timeout_frame.pack(fill="x", pady=10)
        
        ctk.CTkLabel(timeout_frame, text="Page Load Timeout (seconds):").pack(side="left")
        timeout_entry = ctk.CTkEntry(timeout_frame, textvariable=self.page_load_timeout, width=80)
        timeout_entry.pack(side="left", padx=10)
        
        # Max retries
        retries_frame = ctk.CTkFrame(settings_frame)
        retries_frame.pack(fill="x", pady=10)
        
        ctk.CTkLabel(retries_frame, text="Max Retries:").pack(side="left")
        retries_entry = ctk.CTkEntry(retries_frame, textvariable=self.max_retries, width=80)
        retries_entry.pack(side="left", padx=10)
        
        # Save log option
        log_frame = ctk.CTkFrame(settings_frame)
        log_frame.pack(fill="x", pady=10)
        
        log_switch = ctk.CTkSwitch(
            log_frame, 
            text="Save Log Files", 
            variable=self.save_log_var,
            onvalue=True,
            offvalue=False
        )
        log_switch.pack(side="left")
        
        # Save button
        save_button = ctk.CTkButton(
            settings_frame, 
            text="Save Settings", 
            command=self.save_settings,
            fg_color="#007bff",
            hover_color="#0069d9"
        )
        save_button.pack(pady=20)

        # Auto search settings
        auto_frame = ctk.CTkFrame(settings_frame)
        auto_frame.pack(fill="x", pady=10)

        auto_switch = ctk.CTkSwitch(
            auto_frame, 
            text="Enable Auto Search (every 24 hours)", 
            variable=self.auto_enabled_var,
            onvalue=True,
            offvalue=False
        )
        auto_switch.pack(side="left")
        ToolTip(auto_switch, "Automatically run searches once per day when the app starts")

        interval_frame = ctk.CTkFrame(settings_frame)
        interval_frame.pack(fill="x", pady=10)
        ctk.CTkLabel(interval_frame, text="Auto Search Interval (hours):").pack(side="left")
        interval_entry = ctk.CTkEntry(interval_frame, textvariable=self.auto_interval_hours, width=80)
        interval_entry.pack(side="left", padx=10)

        startup_frame = ctk.CTkFrame(settings_frame)
        startup_frame.pack(fill="x", pady=10)
        startup_switch = ctk.CTkSwitch(
            startup_frame,
            text="Run on system login (if supported)",
            variable=self.run_on_startup_var,
            onvalue=True,
            offvalue=False
        )
        startup_switch.pack(side="left")
        ToolTip(startup_switch, "Launch the app automatically when you sign in")

    def create_about_tab(self):
        about_frame = ctk.CTkFrame(self.tab_about)
        about_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # App info
        ctk.CTkLabel(
            about_frame, 
            text="Bing Search Automation", 
            font=("Helvetica", 20, "bold")
        ).pack(pady=10)
        
        ctk.CTkLabel(
            about_frame, 
            text="Version 1.2.0", 
            font=("Helvetica", 14)
        ).pack()
        
        ctk.CTkLabel(
            about_frame, 
            text="This application automates Bing searches using Selenium WebDriver.\n"
                 "It's designed to read search terms from a text file and perform searches automatically.",
            wraplength=500
        ).pack(pady=20)
        
        # Instructions
        instructions_frame = ctk.CTkFrame(about_frame)
        instructions_frame.pack(fill="x", pady=10)
        
        ctk.CTkLabel(
            instructions_frame, 
            text="How to use:", 
            font=("Helvetica", 16, "bold")
        ).pack(anchor="w", pady=5)
        
        instructions = (
            "1. Select a text file containing search terms (one per line)\n"
            "2. Select the Edge WebDriver (msedgedriver.exe)\n"
            "3. Configure settings as needed\n"
            "4. Click 'Start Searching' to begin"
        )
        
        ctk.CTkLabel(
            instructions_frame, 
            text=instructions,
            justify="left"
        ).pack(anchor="w", padx=20)
        
        # Copyright
        ctk.CTkLabel(
            about_frame, 
            text="© 2024 Refter. All rights reserved.\nLicensed under the MIT License.",
            font=("Helvetica", 10)
        ).pack(side="bottom", pady=20)

    def create_status_frame(self):
        status_frame = ctk.CTkFrame(self.main_frame)
        status_frame.pack(fill="x", pady=(10, 0))
        
        ctk.CTkLabel(status_frame, text="Status:").pack(side="left", padx=10)
        
        self.status_label = ctk.CTkLabel(status_frame, text="Ready", width=500, anchor="w")
        self.status_label.pack(side="left", fill="x", expand=True, padx=10)
        
        # Progress bar
        progress_frame = ctk.CTkFrame(self.main_frame)
        progress_frame.pack(fill="x", pady=(5, 10))
        
        self.progress_bar = ctk.CTkProgressBar(progress_frame)
        self.progress_bar.pack(fill="x", padx=10, pady=5)
        self.progress_bar.set(0)
        
        self.progress_label = ctk.CTkLabel(progress_frame, text="0/0")
        self.progress_label.pack()

    def create_footer(self):
        footer_frame = ctk.CTkFrame(self.main_frame, height=30)
        footer_frame.pack(fill="x", pady=(10, 0))
        
        copyright_label = ctk.CTkLabel(
            footer_frame, 
            text="© 2024 Refter. All rights reserved.", 
            font=("Helvetica", 10)
        )
        copyright_label.pack(side="right", padx=10)

    def browse_file(self, var, title):
        file_types = [("Text Files", "*.txt"), ("All Files", "*.*")]
        if "WebDriver" in title:
            file_types = [("WebDriver", "*.exe"), ("All Files", "*.*")]
            
        file_path = filedialog.askopenfilename(title=title, filetypes=file_types)
        if file_path:
            var.set(file_path)
            # Persist last paths for auto-run
            try:
                updated = {**self.config}
                if var is self.search_file_var:
                    updated['last_search_file'] = file_path
                if var is self.driver_file_var:
                    updated['last_driver_path'] = file_path
                Config.save(updated)
                self.config = updated
                self.automation.config = updated
            except Exception as e:
                logging.error(f"Failed to persist selected file path: {str(e)}")

    def update_status(self, message):
        self.status_label.configure(text=message)
        self.root.update_idletasks()

    def update_progress(self, current, total):
        if total > 0:
            self.progress_bar.set(current / total)
            self.progress_label.configure(text=f"{current}/{total}")
        else:
            self.progress_bar.set(0)
            self.progress_label.configure(text="0/0")
        self.root.update_idletasks()

    def start_search(self):
        search_file = self.search_file_var.get()
        driver_path = self.driver_file_var.get()

        if not search_file or not driver_path:
            messagebox.showerror("Error", "Please select both the search file and WebDriver file")
            return
        
        # Save current settings first
        self.save_settings()
        
        try:
            self.automation.start_search_thread(search_file, driver_path)
            # Update button states
            self.start_button.configure(state="disabled")
            self.stop_button.configure(state="normal")
        except Exception as e:
            messagebox.showerror("Error", str(e))
            logging.error(f"Error during search: {str(e)}")

    def stop_search(self):
        if hasattr(self.automation, 'stop_search'):
            self.automation.stop_search()
            self.start_button.configure(state="normal")
            self.stop_button.configure(state="disabled")

    def save_settings(self):
        try:
            config = {
                'search_delay': [self.search_delay_min.get(), self.search_delay_max.get()],
                'page_load_timeout': self.page_load_timeout.get(),
                'max_retries': self.max_retries.get(),
                'save_log': self.save_log_var.get(),
                'headless_mode': self.headless_var.get(),
                'appearance_mode': self.appearance_var.get(),
                'color_theme': 'blue',  # Could make this configurable in the future
                # Auto search settings
                'auto_search_enabled': self.auto_enabled_var.get(),
                'auto_search_interval_hours': self.auto_interval_hours.get(),
                'run_on_startup': self.run_on_startup_var.get(),
                # Remember last paths
                'last_search_file': self.search_file_var.get(),
                'last_driver_path': self.driver_file_var.get()
            }
            Config.save(config)
            self.automation.config = config
            self.update_status("Settings saved successfully")
            # Configure OS login startup
            self.configure_run_on_startup(config.get('run_on_startup', False))
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save settings: {str(e)}")

    def change_appearance_mode(self, new_appearance_mode):
        ctk.set_appearance_mode(new_appearance_mode)
        self.appearance_var.set(new_appearance_mode)
        # Save this setting
        self.save_settings()

    def run(self):
        try:
            # Create a resources directory if it doesn't exist
            resources_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'resources')
            os.makedirs(resources_dir, exist_ok=True)
            
            self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
            self.root.mainloop()
        except Exception as e:
            messagebox.showerror("Error", f"Application error: {str(e)}")
            logging.error(f"Application error: {str(e)}")

    def on_closing(self):
        if hasattr(self.automation, 'is_running') and self.automation.is_running:
            if messagebox.askokcancel("Quit", "Search is still running. Do you want to quit anyway?"):
                self.automation.stop_search()
                self.root.destroy()
        else:
            self.root.destroy()

    def maybe_auto_run(self):
        try:
            if not self.config.get('auto_search_enabled', False):
                return
            search_file = self.config.get('last_search_file') or ''
            driver_path = self.config.get('last_driver_path') or ''
            if not search_file or not driver_path or not os.path.isfile(search_file) or not os.path.isfile(driver_path):
                return
            last_run_iso = self.config.get('last_auto_search')
            interval_hours = int(self.config.get('auto_search_interval_hours', 24))
            due = True
            if last_run_iso:
                try:
                    last_dt = datetime.fromisoformat(last_run_iso)
                    due = (datetime.now() - last_dt).total_seconds() >= interval_hours * 3600
                except Exception:
                    due = True
            if not due:
                return
            # Temporarily force headless for auto-run
            previous_headless = bool(self.automation.config.get('headless_mode', False))
            self.automation.config['headless_mode'] = True
            # Start background search
            self.automation.start_search_thread(search_file, driver_path)
            # Reflect UI state
            if hasattr(self, 'start_button') and hasattr(self, 'stop_button'):
                self.start_button.configure(state="disabled")
                self.stop_button.configure(state="normal")
            # Restore headless after completion
            self.root.after(2000, lambda: self._restore_headless_after_run(previous_headless))
        except Exception as e:
            logging.error(f"Auto-run failed: {str(e)}")

    def _restore_headless_after_run(self, previous_headless):
        try:
            if getattr(self.automation, 'is_running', False):
                self.root.after(2000, lambda: self._restore_headless_after_run(previous_headless))
                return
            self.automation.config['headless_mode'] = previous_headless
            self.headless_var.set(previous_headless)
            new_cfg = {**self.config, 'headless_mode': previous_headless}
            Config.save(new_cfg)
            self.config = new_cfg
        except Exception as e:
            logging.error(f"Failed to restore headless setting: {str(e)}")

    def configure_run_on_startup(self, enable):
        try:
            if sys.platform.startswith('win'):
                try:
                    import winreg  # type: ignore
                    run_key_path = r"Software\\Microsoft\\Windows\\CurrentVersion\\Run"
                    app_name = "BingSearchAutomation"
                    exe_path = sys.executable if getattr(sys, 'frozen', False) else sys.argv[0]
                    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, run_key_path, 0, winreg.KEY_ALL_ACCESS) as key:
                        if enable:
                            winreg.SetValueEx(key, app_name, 0, winreg.REG_SZ, f'"{exe_path}"')
                        else:
                            try:
                                winreg.DeleteValue(key, app_name)
                            except FileNotFoundError:
                                pass
                except Exception as e:
                    logging.error(f"Failed to configure Windows startup: {str(e)}")
            elif sys.platform.startswith('linux'):
                try:
                    autostart_dir = os.path.join(os.path.expanduser('~'), '.config', 'autostart')
                    os.makedirs(autostart_dir, exist_ok=True)
                    desktop_file = os.path.join(autostart_dir, 'bing-search-automation.desktop')
                    exec_path = sys.executable if getattr(sys, 'frozen', False) else sys.executable + f" {os.path.abspath(__file__)}"
                    if enable:
                        content = (
                            "[Desktop Entry]\n"
                            "Type=Application\n"
                            "Name=Bing Search Automation\n"
                            f"Exec={exec_path}\n"
                            "X-GNOME-Autostart-enabled=true\n"
                        )
                        with open(desktop_file, 'w') as f:
                            f.write(content)
                    else:
                        if os.path.exists(desktop_file):
                            os.remove(desktop_file)
                except Exception as e:
                    logging.error(f"Failed to configure Linux startup: {str(e)}")
            else:
                # Other OS not supported
                pass
        except Exception as e:
            logging.error(f"configure_run_on_startup error: {str(e)}")

if __name__ == "__main__":
    try:
        app = SearchAutomationGUI()
        app.run()
    except Exception as e:
        messagebox.showerror("Fatal Error", f"Fatal error: {str(e)}")
        logging.error(f"Fatal error: {str(e)}")
        sys.exit(1)
