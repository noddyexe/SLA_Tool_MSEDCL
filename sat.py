import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import psycopg2
from psycopg2.extras import RealDictCursor
import pandas as pd
import os
import datetime
import pickle
import socket
import sys
import time
import threading
import json
from config_loader import get_config



class DBDataFetcher:
    def __init__(self, root, go_back_callback, save_directory, is_exe):
        self.root = root
        self.go_back_callback = go_back_callback
        self.save_directory = save_directory
        self.is_exe = is_exe
        self.root.title("SAT Extraction")
        self.root.geometry("1200x700")
        self.root.configure(bg='white')
        
        # Configuration for retries and batch processing
        self.MAX_RETRIES = 500000
        self.RETRY_INTERVAL_MS = 5000
        self.BATCH_SIZE = 500
        self.METER_BATCH_SIZE = 500
        
        # Batch tracking for resume functionality
        self.batch_state_file = os.path.join(self.save_directory, 'batch_state.json')
        self.current_batch_state = {}
        
        # Style configuration
        self.style = ttk.Style()
        self.style.configure('TFrame', background='MediumAquamarine')
        self.style.configure('TLabel', background='white', font=('Arial', 10))
        self.style.configure('TRadiobutton', background='white')
        self.style.configure('TCheckbutton', background='white')
        
        # Configure custom button styles
        self.style.configure('Flat.TButton', 
                           font=('Arial', 10, 'bold'),
                           borderwidth=0,
                           relief='flat',
                           padding=10,
                           foreground='white',
                           background='#4CAF50')
        
        self.style.map('Flat.TButton',
                      foreground=[('pressed', 'white'), ('active', 'white')],
                      background=[('pressed', '#45a049'), ('active', '#66BB6A')])
        
        self.style.configure('BlueFlat.TButton', 
                           font=('Arial', 10, 'bold'),
                           borderwidth=0,
                           relief='flat',
                           padding=10,
                           foreground='white',
                           background='#2196F3')
        
        self.style.map('BlueFlat.TButton',
                      foreground=[('pressed', 'white'), ('active', 'white')],
                      background=[('pressed', '#1976D2'), ('active', '#42A5F5')])
        
        # Main frames
        self.button_frame = ttk.Frame(root, width=500, padding="10")
        self.button_frame.pack(side=tk.LEFT, fill=tk.Y, expand=False)
        
        self.log_frame = ttk.Frame(root, padding="10")
        self.log_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # Log text widget
        self.log_text = tk.Text(self.log_frame, wrap=tk.WORD, width=80, height=30, 
                               font=('Consolas', 10), bg='black', fg='white')
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # Scrollbar for log
        scrollbar = ttk.Scrollbar(self.log_frame, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
        
        # Add widgets to button frame
        self.create_widgets()
        
        # Initialize variables
        self.excel_data = None
        self.filtered_data = None
        self.loaded_data = None
        self.vpn_connected = False
        self.processing = False
        self.stop_processing = False
        
        # File paths for final summary
        self.master_file_path = ""
        self.blp_file_path = ""
        self.dlp_file_path = ""
        self.billing_file_path = ""
        
        # Database connection cache
        self.db_connections = {}
        
        # Auto mode master data
        self.auto_master_data = None
        
    def log_message(self, message):
        """Add message to log with timestamp and print to terminal"""
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        formatted_message = f"[{timestamp}] {message}\n"
        
        self.log_text.insert(tk.END, formatted_message)
        self.log_text.see(tk.END)
        self.root.update()
        
        try:
            print(formatted_message, end="")
            if sys.stdout is not None:
                sys.stdout.flush()
        except:
            pass
    
    def save_batch_state(self, module, sla_type, data_type, batch_index, total_batches, processed_meters):
        """Save current batch state for resume functionality"""
        state = {
            'module': module,
            'sla_type': sla_type,
            'data_type': data_type,
            'batch_index': batch_index,
            'total_batches': total_batches,
            'processed_meters': processed_meters,
            'from_date': self.from_date_entry.get(),
            'to_date': self.to_date_entry.get(),
            'timestamp': datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        try:
            with open(self.batch_state_file, 'w') as f:
                json.dump(state, f)
        except Exception as e:
            self.log_message(f"Warning: Could not save batch state: {str(e)}")
    
    def load_batch_state(self):
        """Load saved batch state for resume"""
        try:
            if os.path.exists(self.batch_state_file):
                with open(self.batch_state_file, 'r') as f:
                    state = json.load(f)
                return state
        except Exception as e:
            self.log_message(f"Warning: Could not load batch state: {str(e)}")
        return None
    
    def clear_batch_state(self):
        """Clear batch state file after successful completion"""
        try:
            if os.path.exists(self.batch_state_file):
                os.remove(self.batch_state_file)
        except Exception as e:
            self.log_message(f"Warning: Could not clear batch state: {str(e)}")
    
    def create_widgets(self):
        """Create all GUI widgets"""
        # Go Back Button (top right)
        go_back_button = tk.Button(self.button_frame, text="← Go Back", 
                                 command=self.go_back_callback,
                                 bg='#f44336', fg='white',
                                 font=('Arial', 10, 'bold'),
                                 bd=0, padx=15, pady=5,
                                 activebackground='#d32f2f',
                                 activeforeground='white')
        go_back_button.grid(row=0, column=2, sticky=tk.E, pady=5)
        
        go_back_button.bind("<Enter>", lambda e: go_back_button.config(bg='#d32f2f'))
        go_back_button.bind("<Leave>", lambda e: go_back_button.config(bg='#f44336'))
        
        # Title
        title_label = ttk.Label(self.button_frame, text="SAT Data Extraction", 
                              font=('Arial', 14, 'bold'))
        title_label.grid(row=1, column=0, columnspan=2, pady=10)
        
        # VPN Check Button
        vpn_button = tk.Button(self.button_frame, text="Check VPN Connection", 
                              command=self.check_vpn, 
                              bg='#2196F3', fg='white',
                              font=('Arial', 10, 'bold'),
                              bd=0, padx=20, pady=10,
                              activebackground='#42A5F5',
                              activeforeground='white')
        vpn_button.grid(row=2, column=0, columnspan=2, pady=5, sticky=tk.EW)
        
        vpn_button.bind("<Enter>", lambda e: vpn_button.config(bg='#42A5F5'))
        vpn_button.bind("<Leave>", lambda e: vpn_button.config(bg='#2196F3'))
        
        # Sync Data Option (YES/NO)
        self.fetch_var = tk.StringVar(value="NO")
        fetch_label = ttk.Label(self.button_frame, text="Sync Data:")
        fetch_label.grid(row=3, column=0, sticky=tk.W, pady=5)
        
        fetch_yes = ttk.Radiobutton(self.button_frame, text="YES", variable=self.fetch_var, 
                                   value="YES", command=self.toggle_fetch_options)
        fetch_yes.grid(row=3, column=1, sticky=tk.W, pady=5)
        
        fetch_no = ttk.Radiobutton(self.button_frame, text="NO", variable=self.fetch_var, 
                                  value="NO", command=self.toggle_fetch_options)
        fetch_no.grid(row=3, column=2, sticky=tk.W, pady=5)
        
        # Container frame for options shown when Sync Data = YES
        self.options_frame = ttk.Frame(self.button_frame)
        self.options_frame.grid(row=4, column=0, columnspan=3, sticky=tk.EW)
        
        # ── NEW: Fetch Mode (Manual / Auto) ──────────────────────────────────
        self.fetch_mode_var = tk.StringVar(value="Manual")
        fetch_mode_label = ttk.Label(self.options_frame, text="Fetch Mode:")
        fetch_mode_label.grid(row=0, column=0, sticky=tk.W, pady=5)

        fetch_manual_rb = ttk.Radiobutton(self.options_frame, text="Manual",
                                          variable=self.fetch_mode_var,
                                          value="Manual",
                                          command=self.toggle_fetch_mode)
        fetch_manual_rb.grid(row=0, column=1, sticky=tk.W, pady=5)

        fetch_auto_rb = ttk.Radiobutton(self.options_frame, text="Auto",
                                        variable=self.fetch_mode_var,
                                        value="Auto",
                                        command=self.toggle_fetch_mode)
        fetch_auto_rb.grid(row=0, column=2, sticky=tk.W, pady=5)

        # Sub-frame for Manual-only widgets (Select File button)
        self.manual_only_frame = ttk.Frame(self.options_frame)
        self.manual_only_frame.grid(row=1, column=0, columnspan=3, sticky=tk.EW)

        excel_button = tk.Button(self.manual_only_frame, text="Select File", 
                               command=self.load_excel, 
                               bg='#2196F3', fg='white',
                               font=('Arial', 10, 'bold'),
                               bd=0, padx=20, pady=10,
                               activebackground='#42A5F5',
                               activeforeground='white')
        excel_button.grid(row=0, column=0, columnspan=2, pady=5, sticky=tk.EW)
        excel_button.bind("<Enter>", lambda e: excel_button.config(bg='#42A5F5'))
        excel_button.bind("<Leave>", lambda e: excel_button.config(bg='#2196F3'))
        # ─────────────────────────────────────────────────────────────────────

        # Module Dropdown
        self.db_type_var = tk.StringVar()
        db_type_label = ttk.Label(self.options_frame, text="Module:")
        db_type_label.grid(row=2, column=0, sticky=tk.W, pady=5)
        
        self.db_type_combo = ttk.Combobox(self.options_frame, textvariable=self.db_type_var, 
                                        values=["SLA", "MDM", "HES", "ALL"],
                                        state="readonly")
        self.db_type_combo.grid(row=2, column=1, columnspan=2, sticky=tk.EW, pady=5)
        self.db_type_combo.set("SLA")
        
        # Meter Type Dropdown
        self.meter_type_var = tk.StringVar()
        meter_type_label = ttk.Label(self.options_frame, text="Meter Type:")
        meter_type_label.grid(row=3, column=0, sticky=tk.W, pady=5)
        
        self.meter_type_combo = ttk.Combobox(self.options_frame, textvariable=self.meter_type_var, 
                                           values=["SAT1", "SAT2", "SAT3", "SAT4", "SAT5","SAT6","SAT7","SAT8","SAT9","SAT10","SAT11","SAT12","SAT13","SAT14","SAT15","SAT16","SAT17","SAT18","SAT19","SAT20","SAT21","SAT22","SAT23","SAT24","SAT25","SAT26","SAT27","SAT28","SAT29","SAT30","SAT31","SAT32","SAT33","SAT34","SAT35","SAT36","SAT37","SAT38","SAT39","SAT40","ALLSAT"],
                                           state="readonly")
        self.meter_type_combo.grid(row=3, column=1, columnspan=2, sticky=tk.EW, pady=5)
        self.meter_type_combo.set("ALLSAT")
        
        # HES Type Dropdown
        self.hes_type_var = tk.StringVar()
        hes_type_label = ttk.Label(self.options_frame, text="HES Type:")
        hes_type_label.grid(row=4, column=0, sticky=tk.W, pady=5)
        
        self.hes_type_combo = ttk.Combobox(self.options_frame, textvariable=self.hes_type_var, 
                                         values=["BCITS", "Cyan", "Crystal", "ALL"],
                                         state="readonly")
        self.hes_type_combo.grid(row=4, column=1, columnspan=2, sticky=tk.EW, pady=5)
        self.hes_type_combo.set("ALL")
        
        # Data Type Dropdown
        self.data_type_var = tk.StringVar()
        data_type_label = ttk.Label(self.options_frame, text="Data Type:")
        data_type_label.grid(row=5, column=0, sticky=tk.W, pady=5)
        
        self.data_type_combo = ttk.Combobox(self.options_frame, textvariable=self.data_type_var, 
                                          values=["Meter-wise", "Day-wise"],
                                          state="readonly")
        self.data_type_combo.grid(row=5, column=1, columnspan=2, sticky=tk.EW, pady=5)
        self.data_type_combo.set("Meter-wise")
        
        # SLA Type Dropdown
        self.sla_type_var = tk.StringVar()
        sla_type_label = ttk.Label(self.options_frame, text="Profile Type:")
        sla_type_label.grid(row=6, column=0, sticky=tk.W, pady=5)
        
        self.sla_type_combo = ttk.Combobox(self.options_frame, textvariable=self.sla_type_var, 
                                         values=["BLP", "DLP", "BLP-DLP", "BILLING", "ALL"],
                                         state="readonly")
        self.sla_type_combo.grid(row=6, column=1, columnspan=2, sticky=tk.EW, pady=5)
        self.sla_type_combo.set("BLP")
        
        # Date Selection
        date_frame = ttk.Frame(self.options_frame)
        date_frame.grid(row=7, column=0, columnspan=3, pady=5, sticky=tk.EW)
        
        from_date_label = ttk.Label(date_frame, text="From Date:")
        from_date_label.grid(row=0, column=0, padx=5)
        
        self.from_date_entry = ttk.Entry(date_frame, width=12)
        self.from_date_entry.grid(row=0, column=1, padx=5)
        self.from_date_entry.insert(0, datetime.datetime.now().strftime("%Y-%m-%d"))
        
        to_date_label = ttk.Label(date_frame, text="To Date:")
        to_date_label.grid(row=0, column=2, padx=5)
        
        self.to_date_entry = ttk.Entry(date_frame, width=12)
        self.to_date_entry.grid(row=0, column=3, padx=5)
        self.to_date_entry.insert(0, datetime.datetime.now().strftime("%Y-%m-%d"))
        
        # Run Process Button
        self.run_button = tk.Button(self.button_frame, text="Run Process", 
                             command=self.start_process_thread, 
                             bg='#4CAF50', fg='white',
                             font=('Arial', 10, 'bold'),
                             bd=0, padx=20, pady=10,
                             activebackground='#66BB6A',
                             activeforeground='white')
        self.run_button.grid(row=5, column=0, columnspan=3, pady=20, sticky=tk.EW)
        
        # Stop Process Button (initially hidden)
        self.stop_button = tk.Button(self.button_frame, text="Stop Process", 
                             command=self.stop_process, 
                             bg='#f44336', fg='white',
                             font=('Arial', 10, 'bold'),
                             bd=0, padx=20, pady=10,
                             activebackground='#d32f2f',
                             activeforeground='white')
        
        self.run_button.bind("<Enter>", lambda e: self.run_button.config(bg='#66BB6A'))
        self.run_button.bind("<Leave>", lambda e: self.run_button.config(bg='#4CAF50'))
        
        # Final Summary Frame (initially hidden)
        self.summary_frame = ttk.Frame(self.button_frame)
        
        self.summary_label = ttk.Label(self.summary_frame, text="Select folders for the final summary...", 
                                     font=('Arial', 12, 'bold'))
        self.summary_label.grid(row=0, column=0, columnspan=3, pady=10)
        
        self.final_summary_var = tk.StringVar(value="NO")
        final_summary_label = ttk.Label(self.summary_frame, text="Run Final Summary:")
        final_summary_label.grid(row=1, column=0, sticky=tk.W, pady=5)
        
        final_summary_yes = ttk.Radiobutton(self.summary_frame, text="YES", variable=self.final_summary_var, 
                                         value="YES", command=self.toggle_file_selection_buttons)
        final_summary_yes.grid(row=1, column=1, sticky=tk.W, pady=5)
        
        final_summary_no = ttk.Radiobutton(self.summary_frame, text="NO", variable=self.final_summary_var, 
                                        value="NO", command=self.toggle_file_selection_buttons)
        final_summary_no.grid(row=1, column=2, sticky=tk.W, pady=5)
        
        self.file_selection_frame = ttk.Frame(self.summary_frame)
        
        self.master_button = tk.Button(self.file_selection_frame, text="Select Master File", 
                                    command=self.select_master_file,
                                    bg='#2196F3', fg='white',
                                    font=('Arial', 10, 'bold'),
                                    bd=0, padx=20, pady=10,
                                    activebackground='#42A5F5',
                                    activeforeground='white',
                                    state=tk.DISABLED)
        self.master_button.grid(row=0, column=0, columnspan=3, pady=5, sticky=tk.EW)
        
        self.blp_button = tk.Button(self.file_selection_frame, text="Select BLP RAW File", 
                                  command=self.select_blp_file,
                                  bg='#2196F3', fg='white',
                                  font=('Arial', 10, 'bold'),
                                  bd=0, padx=20, pady=10,
                                  activebackground='#42A5F5',
                                  activeforeground='white',
                                  state=tk.DISABLED)
        self.blp_button.grid(row=1, column=0, columnspan=3, pady=5, sticky=tk.EW)
        
        self.dlp_button = tk.Button(self.file_selection_frame, text="Select DLP RAW File", 
                                  command=self.select_dlp_file,
                                  bg='#2196F3', fg='white',
                                  font=('Arial', 10, 'bold'),
                                  bd=0, padx=20, pady=10,
                                  activebackground='#42A5F5',
                                  activeforeground='white',
                                  state=tk.DISABLED)
        self.dlp_button.grid(row=2, column=0, columnspan=3, pady=5, sticky=tk.EW)
        
        self.billing_button = tk.Button(self.file_selection_frame, text="Select BILLING RAW File", 
                                      command=self.select_billing_file,
                                      bg='#2196F3', fg='white',
                                      font=('Arial', 10, 'bold'),
                                      bd=0, padx=20, pady=10,
                                      activebackground='#42A5F5',
                                      activeforeground='white',
                                      state=tk.DISABLED)
        self.billing_button.grid(row=3, column=0, columnspan=3, pady=5, sticky=tk.EW)
        
        self.run_final_button = tk.Button(self.summary_frame, text="Run Final Process", 
                                       command=self.run_final_process,
                                       bg='#FF5722', fg='white',
                                       font=('Arial', 10, 'bold'),
                                       bd=0, padx=20, pady=10,
                                       activebackground='#E64A19',
                                       activeforeground='white',
                                       state=tk.DISABLED)
        self.run_final_button.grid(row=3, column=0, columnspan=3, pady=20, sticky=tk.EW)
        
        self.run_final_button.bind("<Enter>", lambda e: self.run_final_button.config(bg='#E64A19'))
        self.run_final_button.bind("<Leave>", lambda e: self.run_final_button.config(bg='#FF5722'))
        
        # Initially hide options frame if fetch is NO
        self.toggle_fetch_options()
    
    def start_process_thread(self):
        """Start the process in a separate thread to keep GUI responsive"""
        if not self.processing:
            self.processing = True
            self.stop_processing = False
            self.run_button.grid_remove()
            self.stop_button.grid(row=5, column=0, columnspan=3, pady=20, sticky=tk.EW)
            
            process_thread = threading.Thread(target=self.run_process)
            process_thread.daemon = True
            process_thread.start()
    
    def stop_process(self):
        """Stop the current process"""
        self.stop_processing = True
        self.log_message("Stopping process...")
    
    def toggle_fetch_options(self):
        """Show/hide fetch options based on Sync Data YES/NO selection"""
        if self.fetch_var.get() == "YES":
            self.options_frame.grid()
            self.summary_frame.grid_remove()
        else:
            self.options_frame.grid_remove()
            self.summary_frame.grid(row=6, column=0, columnspan=3, pady=10, sticky=tk.EW)

    def toggle_fetch_mode(self):
        """Show/hide Select File button and adjust Module options based on Manual/Auto"""
        mode = self.fetch_mode_var.get()
        if mode == "Manual":
            # Show Select File button
            self.manual_only_frame.grid()
            # Restore full module list (includes HES)
            self.db_type_combo.config(values=["SLA", "MDM", "HES", "ALL"])
            if self.db_type_var.get() not in ["SLA", "MDM", "HES", "ALL"]:
                self.db_type_combo.set("SLA")
        else:  # Auto
            # Hide Select File button
            self.manual_only_frame.grid_remove()
            # Auto mode: no HES option
            self.db_type_combo.config(values=["SLA", "MDM", "ALL"])
            if self.db_type_var.get() == "HES":
                self.db_type_combo.set("SLA")
    
    def toggle_file_selection_buttons(self):
        """Enable/disable file selection buttons based on final summary selection"""
        if self.final_summary_var.get() == "YES":
            self.file_selection_frame.grid(row=2, column=0, columnspan=3, pady=5, sticky=tk.EW)
            self.master_button.config(state=tk.NORMAL)
            self.blp_button.config(state=tk.NORMAL)
            self.dlp_button.config(state=tk.NORMAL)
            self.billing_button.config(state=tk.NORMAL)
            self.run_final_button.config(state=tk.NORMAL)
        else:
            self.file_selection_frame.grid_remove()
            self.master_button.config(state=tk.DISABLED)
            self.blp_button.config(state=tk.DISABLED)
            self.dlp_button.config(state=tk.DISABLED)
            self.billing_button.config(state=tk.DISABLED)
            self.run_final_button.config(state=tk.DISABLED)
            self.log_message("Final summary process cancelled.")

    
    def check_vpn(self):
        """Check VPN connection by testing database server connectivity based on selected module"""
        module = self.db_type_var.get()
        self.log_message(f"Checking VPN connection for {module} module...")
        
        if module == "SLA":
            host = get_config()['SLA']['ip']
        elif module == "MDM":
            host = get_config()['MDM']['ip']
        elif module == "HES":
            host = get_config()['HES']['ip']
        else:  # ALL
            host = get_config()['SLA']['ip']
        
        port = int(get_config()['SLA']['port'])
        
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            s.settimeout(5)
            s.connect((host, port))
            s.close()
            self.log_message(f"VPN connected successfully for {module} module!")
            self.vpn_connected = True
        except socket.timeout:
            self.log_message(f"VPN connection timeout for {module} module - server not responding")
            self.vpn_connected = False
        except ConnectionRefusedError:
            self.log_message(f"VPN connection refused for {module} module - server may be down")
            self.vpn_connected = False
        except socket.gaierror:
            self.log_message(f"VPN hostname resolution failed for {module} module - check DNS or VPN")
            self.vpn_connected = False
        except Exception as e:
            self.log_message(f"VPN check error for {module} module: {str(e)}")
            self.vpn_connected = False
    
    def wait_for_vpn(self, module, max_retries=1000, retry_interval=5):
        """Wait for VPN connection with retries"""
        self.log_message(f"Waiting for VPN connection for {module} module...")
        
        if module == "SLA":
            host = get_config()['SLA']['ip']
        elif module == "MDM":
            host = get_config()['MDM']['ip']
        elif module == "HES":
            host = get_config()['HES']['ip']
        else:
            host = get_config()['SLA']['ip']
        
        port = int(get_config()['SLA']['port'])
        
        retries = 0
        while retries < max_retries and not self.stop_processing:
            try:
                s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                s.settimeout(5)
                s.connect((host, port))
                s.close()
                self.log_message(f"VPN connected successfully for {module} module!")
                self.vpn_connected = True
                return True
            except Exception as e:
                self.log_message(f"VPN connection attempt {retries + 1}/{max_retries} failed: {str(e)}")
                self.vpn_connected = False
                retries += 1
                for i in range(retry_interval):
                    if self.stop_processing:
                        return False
                    time.sleep(1)
        
        return False
    
    def load_excel(self):
        """Load Excel file and save to local data_selve.dat"""
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
        )
        
        if not file_path:
            self.log_message("No file selected. Operation cancelled.")
            return
        
        try:
            self.log_message(f"Loading Excel file: {file_path}")
            self.excel_data = pd.read_excel(file_path)
            
            required_columns = ['Meter Serial Number', 'SAT_No', 'IP', 'HES']
            missing_cols = [col for col in required_columns if col not in self.excel_data.columns]
            
            if missing_cols:
                self.log_message(f"Error: Missing columns in Excel: {', '.join(missing_cols)}")
                messagebox.showerror("Error", f"Missing columns in Excel: {', '.join(missing_cols)}")
                return
            
            data_file = os.path.join(self.save_directory, 'data_selve.dat')
            with open(data_file, 'wb') as f:
                pickle.dump(self.excel_data, f)
            
            self.log_message(f"Total records loaded: {len(self.excel_data)}")
            
        except Exception as e:
            self.log_message(f"Error loading Excel file: {str(e)}")
            messagebox.showerror("Error", f"Failed to load Excel file: {str(e)}")
    
    def load_previous_data(self):
        """Load previously saved data if exists"""
        data_file = os.path.join(self.save_directory, 'data_selve.dat')
        if os.path.exists(data_file):
            try:
                with open(data_file, 'rb') as f:
                    self.excel_data = pickle.load(f)
                    self.log_message(f"Loaded previous data with {len(self.excel_data)} records")
            except Exception as e:
                self.log_message(f"Error loading previous data: {str(e)}")
    
    def filter_data(self):
        """Filter data based on selected options"""
        if self.excel_data is None:
            self.log_message("No data available. Please load Excel file first.")
            return None
        
        self.filtered_data = self.excel_data.copy()
        
        meter_type = self.meter_type_var.get()
        if meter_type != "ALLSAT":
            self.filtered_data = self.filtered_data[self.filtered_data['SAT_No'] == meter_type]
            self.log_message(f"Filtered by Meter Type: {meter_type}. Records: {len(self.filtered_data)}")
        
        hes_type = self.hes_type_var.get()
        if hes_type != "ALL":
            self.filtered_data = self.filtered_data[self.filtered_data['HES'] == hes_type]
            self.log_message(f"Filtered by HES Type: {hes_type}. Records: {len(self.filtered_data)}")
        
        return self.filtered_data
    
    def get_meter_numbers(self):
        """Get meter numbers in required format for SQL query"""
        if self.filtered_data is None:
            self.log_message("No filtered data available.")
            return []
        
        meter_numbers = self.filtered_data['Meter Serial Number'].unique()
        return [str(num) for num in meter_numbers if pd.notna(num) and str(num).strip()]
    
    def format_meter_list(self, meter_list):
        """Format meter list for SQL IN clause"""
        if not meter_list:
            return "('')"
        return "('" + "','".join(meter_list) + "')"
    
    def get_hes_name(self):
        """Get HES name for BILLING query"""
        if self.hes_type_var.get() == "ALL":
            return ""
        return self.hes_type_var.get()


    
    def connect_db(self, module):
        """Connect to database using config.ini credentials."""
        if module in self.db_connections:
            conn, cur = self.db_connections[module]
            try:
                cur.execute("SELECT 1")
                return cur, conn
            except:
                pass
    
        section_map = {"SLA": "SLA", "MDM": "MDM", "HES": "HES"}
        if module not in section_map:
            self.log_message(f"Invalid module: {module}")
            return None, None
    
        cfg = get_config()[section_map[module]]
        DBIP, DBPort = cfg['ip'], cfg['port']
        DBUID, DBPwd = cfg['uid'], cfg['pwd']
        BaseDB = cfg['database']
    
        try:
            self.log_message(f"Connecting to {module} database...")
            conn = psycopg2.connect(
                host=DBIP, port=DBPort, database=BaseDB,
                user=DBUID, password=DBPwd
            )
            cur = conn.cursor()
            self.db_connections[module] = (conn, cur)
            return cur, conn
        except (Exception, psycopg2.DatabaseError) as error:
            self.log_message(f'{module} DB Connection Failed. Error: {error}')
            return None, None



    
    def close_db_connection(self, module):
        """Close database connection for a module"""
        if module in self.db_connections:
            try:
                conn, cur = self.db_connections[module]
                cur.close()
                conn.close()
                del self.db_connections[module]
            except:
                pass
    
    def execute_single_query(self, cur, query):
        """Execute single query with batch fetching"""
        results = []
        try:
            cur.execute(query)
            
            while True:
                rows = cur.fetchmany(self.BATCH_SIZE)
                if not rows:
                    break
                
                columns = [desc[0] for desc in cur.description]
                for row in rows:
                    results.append(dict(zip(columns, row)))
                
                if len(results) % (self.BATCH_SIZE * 10) == 0:
                    self.log_message(f"Fetched {len(results)}")
                
                if self.stop_processing:
                    break
            
            return results
        except (Exception, psycopg2.DatabaseError) as error:
            self.log_message(f'Query execution failed. Error: {error}')
            raise

    # ═══════════════════════════════════════════════════════════════════════════
    # AUTO MODE - Master Data Fetch
    # ═══════════════════════════════════════════════════════════════════════════
    def fetch_auto_master_data(self):
        """Fetch master data in Auto mode and save by HES type"""
        self.log_message("\n[Auto Mode] Fetching master data from MDM database...")
        
        query = """
        select meterno as "Meter Serial Number",hes,supported_id as IP,sat_pkg as SAT_No
        from unified_hes.uni_consumer_master
        where satc = 'SATC'
        """
        
        retry_count = 0
        while retry_count < self.MAX_RETRIES and not self.stop_processing:
            try:
                cur, conn = self.connect_db("SLA")
                if cur is None:
                    if not self.wait_for_vpn("SLA", max_retries=10000, retry_interval=5):
                        break
                    cur, conn = self.connect_db("SLA")
                    if cur is None:
                        break
                
                self.log_message("Executing master data query...")
                results = self.execute_single_query(cur, query)
                
                if not results:
                    self.log_message("No master data found.")
                    return None
                
                df = pd.DataFrame(results)
                self.log_message(f"Total master records fetched: {len(df)}")
                
                # Get distinct HES values
                hes_values = df['hes'].unique()
                self.log_message(f"Distinct HES values found: {', '.join(hes_values)}")
                
                # Save data in separate sheets by HES type
                current_date = datetime.datetime.now().strftime("%d%m%y")
                filename = os.path.join(self.save_directory, f"SATC_MtrMaster_RAW_{current_date}.xlsx")
                
                with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                    for hes in hes_values:
                        if pd.notna(hes) and hes.strip():
                            hes_data = df[df['hes'] == hes]
                            sheet_name = f"{hes}_MtrMaster"
                            # Excel sheet names have 31 character limit
                            if len(sheet_name) > 31:
                                sheet_name = sheet_name[:31]
                            hes_data.to_excel(writer, sheet_name=sheet_name, index=False)
                            self.log_message(f"Saved {len(hes_data)} records for HES: {hes} to sheet: {sheet_name}")
                
                self.log_message(f"Master data saved to: {filename}")
                self.auto_master_data = df
                
                # Also save as pickle for use in other queries
                master_pickle_file = os.path.join(self.save_directory, 'auto_master_data.pkl')
                with open(master_pickle_file, 'wb') as f:
                    pickle.dump(df, f)
                
                self.close_db_connection("SLA")
                return df
                
            except Exception as e:
                self.log_message(f"Master data fetch failed (Attempt {retry_count + 1}/{self.MAX_RETRIES}): {str(e)}")
                retry_count += 1
                self.close_db_connection("SLA")
                if retry_count >= self.MAX_RETRIES:
                    self.log_message("Max retries reached for master data fetch.")
                    break
                for i in range(int(self.RETRY_INTERVAL_MS/1000)):
                    if self.stop_processing:
                        break
                    time.sleep(1)
        
        return None

    def get_meter_numbers_auto(self):
        """Get meter numbers from auto master data"""
        if self.auto_master_data is None:
            # Try to load from pickle file
            master_pickle_file = os.path.join(self.save_directory, 'auto_master_data.pkl')
            if os.path.exists(master_pickle_file):
                try:
                    with open(master_pickle_file, 'rb') as f:
                        self.auto_master_data = pickle.load(f)
                        self.log_message(f"Loaded auto master data from pickle: {len(self.auto_master_data)} records")
                except Exception as e:
                    self.log_message(f"Error loading auto master data: {str(e)}")
                    return []
            else:
                return []
        
        # Filter by HES type if selected
        hes_type = self.hes_type_var.get()
        if hes_type != "ALL":
            filtered_data = self.auto_master_data[self.auto_master_data['hes'] == hes_type]
        else:
            filtered_data = self.auto_master_data
        
        # Filter by SAT type if selected
        meter_type = self.meter_type_var.get()
        if meter_type != "ALLSAT":
            filtered_data = filtered_data[filtered_data['Sat_No'] == meter_type]
        
        meter_numbers = filtered_data['meter_serial_number'].unique()
        return [str(num) for num in meter_numbers if pd.notna(num) and str(num).strip()]

    def build_query_auto(self, module, sla_type, data_type, from_date, to_date):
        """Build query for Auto mode - meter list comes from DB subquery."""
        to_date_dt = datetime.datetime.strptime(to_date, "%Y-%m-%d") + datetime.timedelta(days=1)
        query_to_date = to_date_dt.strftime("%Y-%m-%d")
        is_single_day = from_date == to_date

        meter_type = self.meter_type_var.get()

        # Determine subquery filter based on meter type selection
        if meter_type == "ALLSAT":
            sla_subquery = "select meterno from unified_hes.uni_consumer_master where satc = 'SATC'"
            mdm_subquery = "select mtrno from ami_master.master_main where mla = 'SATC'"
        else:
            sla_subquery = f"select meterno from unified_hes.uni_consumer_master where sat_pkg= '{meter_type}'"
            mdm_subquery = f"select mtrno from ami_master.master_main where village = '{meter_type}'"

        if module == "SLA":
            if sla_type == "BLP":
                if data_type == "Meter-wise":
                    return f"""
                    select meter_number,hes_name,sum(ls_tcount_first_timeslot) as H8,
                    sum(ls_tcount_second_timeslot) as H12,sum(ls_tcount_third_timeslot) as H24 
                    from unified_hes.uni_ca_agg_load_data 
                    where report_date >= '{from_date}' and report_date < '{query_to_date}'
                    and meter_number in ({sla_subquery})
                    group by meter_number,hes_name
                    """
                elif data_type == "Day-wise":
                    if is_single_day:
                        return f"""
                        select meter_number,hes_name,report_date,
                        sum(ls_tcount_first_timeslot) as H8,sum(ls_tcount_second_timeslot) as H12,
                        sum(ls_tcount_third_timeslot) as H24 from unified_hes.uni_ca_agg_load_data 
                        where report_date = '{from_date}'
                        and meter_number in ({sla_subquery})
                        group by meter_number,hes_name,report_date
                        """
                    else:
                        return f"""
                        select meter_number,hes_name,report_date,
                        sum(ls_tcount_first_timeslot) as H8,sum(ls_tcount_second_timeslot) as H12,
                        sum(ls_tcount_third_timeslot) as H24 from unified_hes.uni_ca_agg_load_data 
                        where report_date >= '{from_date}' and report_date < '{query_to_date}'
                        and meter_number in ({sla_subquery})
                        group by meter_number,hes_name,report_date
                        """
            elif sla_type == "DLP":
                if data_type == "Meter-wise":
                    return f"""
                    select meter_number, hes_name, sum(ls_tcount_third_timeslot) as h24 from unified_hes.uni_ca_agg_daily_data
                    where report_date >= '{from_date}' and report_date < '{query_to_date}'
                    and meter_number in ({sla_subquery})
                    group by meter_number, hes_name
                    """
                elif data_type == "Day-wise":
                    if is_single_day:
                        return f"""
                        select meter_number, hes_name,report_date, count(*) as h24 from unified_hes.uni_agg_daily_data
                        where time_period = 'M+24h' 
                        and report_date = '{from_date}'
                        and meter_number in ({sla_subquery})
                        group by meter_number, hes_name,report_date
                        """
                    else:
                        return f"""
                        select meter_number, hes_name,report_date, count(*) as h24 from unified_hes.uni_agg_daily_data
                        where time_period = 'M+24h' 
                        and report_date >= '{from_date}' and report_date < '{query_to_date}'
                        and meter_number in ({sla_subquery})
                        group by meter_number, hes_name,report_date
                        """
            elif sla_type == "BILLING":
                from_date_with_time = f"{from_date} 00:00:00"
                to_date_with_time = f"{query_to_date} 00:00:00"
                hes_name = self.get_hes_name()
                hes_condition = f"and hes_name = '{hes_name}'" if hes_name else ""
                return f"""
                select hes_name,time_stamp,meter_number,bill_date,server_time,sla_id from unified_hes.uni_bill_data
                where bill_date >= '2026-08-01' and bill_date < '2026-08-02'
                and server_time <= '2026-08-01 12:00:00'
                and meter_number in ({sla_subquery})
                {hes_condition}
                order by bill_date
                """

# where bill_date >= '2026-08-01' and bill_date < '2026-08-02'
#                 and server_time >= '2026-08-01 12:00:00'


        elif module == "MDM":
            if sla_type == "BILLING":
                return f"""
                select meter_number,billing_date,server_time as HES_time,time_stamp as MDM_time from meter_data.bill_history
                where billing_date >= '{from_date}' and billing_date < '{query_to_date}'
                and meter_number in ({mdm_subquery})
                """
            elif sla_type == "BLP":
                return f"""
                select meter_number, count(*) from meter_data.load_survey 
                where read_time >= '{from_date}' and read_time < '{query_to_date}'
                and meter_number in ({mdm_subquery})
                group by meter_number
                """
            elif sla_type == "DLP":
                return f"""
                select mtrno, count(mtrno) as MDM_DLP from meter_data.daily_load 
                where rtc_date_time >= '{from_date}' and rtc_date_time < '{query_to_date}'
                and mtrno in ({mdm_subquery})
                group by mtrno
                """

        return None

    def execute_day_wise_query_auto(self, module, sla_type, from_date, to_date):
        """Execute day-wise query for Auto mode (no meter batching)"""
        all_results = []

        start_date = datetime.datetime.strptime(from_date, "%Y-%m-%d")
        end_date = datetime.datetime.strptime(to_date, "%Y-%m-%d")

        current_date = start_date
        while current_date <= end_date and not self.stop_processing:
            date_str = current_date.strftime("%Y-%m-%d")

            query = self.build_query_auto(module, sla_type, "Day-wise", date_str, date_str)

            if query:
                retry_count = 0
                while retry_count < self.MAX_RETRIES and not self.stop_processing:
                    try:
                        cur, conn = self.connect_db(module)
                        if cur is None:
                            if not self.wait_for_vpn(module, max_retries=10000, retry_interval=5):
                                break
                            cur, conn = self.connect_db(module)
                            if cur is None:
                                break

                        day_results = self.execute_single_query(cur, query)
                        all_results.extend(day_results)
                        self.log_message(f"Data for {date_str}: {len(day_results)} records")
                        break

                    except Exception as e:
                        self.log_message(f"Day-wise query failed for {date_str} (Attempt {retry_count + 1}/{self.MAX_RETRIES}): {str(e)}")
                        retry_count += 1
                        if retry_count >= self.MAX_RETRIES:
                            self.log_message(f"Max retries reached for day {date_str}")
                            break
                        for i in range(int(self.RETRY_INTERVAL_MS / 1000)):
                            if self.stop_processing:
                                break
                            time.sleep(1)

            current_date += datetime.timedelta(days=1)

        return all_results

    def execute_query_auto(self, module, sla_type, data_type, from_date, to_date):
        """Execute query in Auto mode - single query per module, no meter batching."""
        if module == "ALL":
            combined = []
            combined.extend(self.execute_query_auto("SLA", sla_type, data_type, from_date, to_date))
            combined.extend(self.execute_query_auto("MDM", sla_type, data_type, from_date, to_date))
            return combined

        if sla_type == "BLP-DLP":
            combined = []
            combined.extend(self.execute_query_auto(module, "BLP", data_type, from_date, to_date))
            combined.extend(self.execute_query_auto(module, "DLP", data_type, from_date, to_date))
            return combined

        if sla_type == "ALL":
            combined = []
            combined.extend(self.execute_query_auto(module, "BLP", data_type, from_date, to_date))
            combined.extend(self.execute_query_auto(module, "DLP", data_type, from_date, to_date))
            combined.extend(self.execute_query_auto(module, "BILLING", data_type, from_date, to_date))
            return combined

        retry_count = 0
        while retry_count < self.MAX_RETRIES and not self.stop_processing:
            try:
                cur, conn = self.connect_db(module)
                if cur is None:
                    if not self.wait_for_vpn(module, max_retries=10000, retry_interval=5):
                        break
                    cur, conn = self.connect_db(module)
                    if cur is None:
                        break

                if data_type == "Day-wise" and module == "SLA":
                    results = self.execute_day_wise_query_auto(module, sla_type, from_date, to_date)
                else:
                    query = self.build_query_auto(module, sla_type, data_type, from_date, to_date)
                    if not query:
                        self.log_message(f"No query built for {module}/{sla_type}/{data_type}")
                        return []
                    results = self.execute_single_query(cur, query)

                self.log_message(f"[Auto] {module} {sla_type} fetched {len(results)} records")
                self.close_db_connection(module)
                return results

            except Exception as e:
                self.log_message(f"[Auto][{module}] Query failed (Attempt {retry_count + 1}/{self.MAX_RETRIES}): {str(e)}")
                retry_count += 1
                self.close_db_connection(module)
                if retry_count >= self.MAX_RETRIES:
                    self.log_message(f"Max retries reached for module {module}")
                    break
                for i in range(int(self.RETRY_INTERVAL_MS / 1000)):
                    if self.stop_processing:
                        break
                    time.sleep(1)

        return []

    def get_query_auto(self, sla_type, module):
        """Entry point for Auto mode query execution"""
        data_type = self.data_type_var.get()
        from_date = self.from_date_entry.get()
        to_date = self.to_date_entry.get()

        results = self.execute_query_auto(module, sla_type, data_type, from_date, to_date)

        if not results:
            return pd.DataFrame()

        return pd.DataFrame(results)

    # ═══════════════════════════════════════════════════════════════════════════
    # MANUAL MODE - original query builder (uses meter_list from Excel)
    # ═══════════════════════════════════════════════════════════════════════════
    def build_query(self, module, sla_type, data_type, from_date, to_date, meter_list):
        """Build query based on parameters (Manual mode - uses meter list from Excel)"""
        meter_in_clause = self.format_meter_list(meter_list)
        
        to_date_dt = datetime.datetime.strptime(to_date, "%Y-%m-%d") + datetime.timedelta(days=1)
        query_to_date = to_date_dt.strftime("%Y-%m-%d")
        
        is_single_day = from_date == to_date
        
        if module == "SLA":
            if sla_type == "BLP":
                if data_type == "Meter-wise":
                    return f"""
                    select meter_number,hes_name,sum(ls_tcount_first_timeslot) as H8,
                    sum(ls_tcount_second_timeslot) as H12,sum(ls_tcount_third_timeslot) as H24 
                    from unified_hes.uni_ca_agg_load_data 
                    where report_date >= '{from_date}' and report_date < '{query_to_date}'
                    and meter_number in {meter_in_clause}
                    group by meter_number,hes_name
                    """
                elif data_type == "Day-wise":
                    if is_single_day:
                        return f"""
                        select meter_number,hes_name,report_date,
                        sum(ls_tcount_first_timeslot) as H8,sum(ls_tcount_second_timeslot) as H12,
                        sum(ls_tcount_third_timeslot) as H24 from unified_hes.uni_ca_agg_load_data 
                        where report_date = '{from_date}'
                        and meter_number in {meter_in_clause}
                        group by meter_number,hes_name,report_date
                        """
                    else:
                        return f"""
                        select meter_number,hes_name,report_date,
                        sum(ls_tcount_first_timeslot) as H8,sum(ls_tcount_second_timeslot) as H12,
                        sum(ls_tcount_third_timeslot) as H24 from unified_hes.uni_ca_agg_load_data 
                        where report_date >= '{from_date}' and report_date < '{query_to_date}'
                        and meter_number in {meter_in_clause}
                        group by meter_number,hes_name,report_date
                        """
            elif sla_type == "DLP":
                if data_type == "Meter-wise":
                    return f"""
                    select meter_number, hes_name, SUM(ls_tcount_third_timeslot) as h24 from unified_hes.uni_ca_agg_daily_data
                    where report_date >= '{from_date}' and report_date < '{query_to_date}'
                    and meter_number in {meter_in_clause}
                    group by meter_number, hes_name
                    """
                elif data_type == "Day-wise":
                    if is_single_day:
                        return f"""
                        select meter_number, hes_name,report_date, count(*) as h24 from unified_hes.uni_agg_daily_data
                        where time_period = 'M+24h' 
                        and report_date = '{from_date}'
                        and meter_number in {meter_in_clause}
                        group by meter_number, hes_name,report_date
                        """
                    else:
                        return f"""
                        select meter_number, hes_name,report_date, count(*) as h24 from unified_hes.uni_agg_daily_data
                        where time_period = 'M+24h' 
                        and report_date >= '{from_date}' and report_date < '{query_to_date}'
                        and meter_number in {meter_in_clause}
                        group by meter_number, hes_name,report_date
                        """
            elif sla_type == "BILLING":
                from_date_with_time = f"{from_date} 00:00:00"
                to_date_with_time = f"{query_to_date} 00:00:00"
                
                hes_name = self.get_hes_name()
                hes_condition = f"and hes_name = '{hes_name}'" if hes_name else ""
                
                return f"""
                select hes_name,time_stamp,meter_number,bill_date,server_time,sla_id from unified_hes.uni_bill_data
                where bill_date >= '2026-08-01' and bill_date < '2026-08-02'
                and server_time <= '2026-08-01 12:00:00'
                and meter_number in {meter_in_clause}
                {hes_condition}
                order by bill_date
                """
                
# where bill_date >= '2026-08-01' and bill_date < '2026-08-02'
#                 and server_time <= '2026-08-01 12:00:00'                
                
        elif module == "MDM":
            if sla_type == "BILLING":
                return f"""
                select meter_number,billing_date,server_time as HES_time,time_stamp as MDM_time from meter_data.bill_history
                where billing_date >= '{from_date}' and billing_date < '{query_to_date}'
                and meter_number in {meter_in_clause}
                """
            elif sla_type == "BLP":
                return f"""
                select meter_number, count(*) as MDM_BLP from meter_data.load_survey 
                where read_time >= '{from_date}' and read_time < '{query_to_date}'
                and meter_number in {meter_in_clause}
                group by meter_number
                """
            elif sla_type == "DLP":
                return f"""
                select mtrno, count(mtrno) as MDM_DLP from meter_data.daily_load 
                where rtc_date_time >= '{from_date}' and rtc_date_time < '{query_to_date}'
                and mtrno in {meter_in_clause}
                group by mtrno
                """
        elif module == "HES":
            if sla_type == "BILLING":
                return f"""
                select meter_number,bill_date,time_stamp as HES_time from  hes.data_bill 
                where bill_date >= '{from_date}' and bill_date < '{query_to_date}'
                and meter_number in {meter_in_clause}
                """
            elif sla_type == "BLP":
                return f"""
                select meter_number,count(meter_number) as HES_BLP from hes.data_block_load 
                where load_time >= '{from_date}' and load_time < '{query_to_date}'
                and meter_number in {meter_in_clause}
                group by meter_number
                """
            elif sla_type == "DLP":
                return f"""
                select meter_number,count(meter_number) as HES_DLP from hes.data_daily_load 
                where load_time >= '{from_date}' and load_time < '{query_to_date}'
                and meter_number in {meter_in_clause}
                group by meter_number
                """
        
        return None

    def execute_day_wise_query(self, module, sla_type, from_date, to_date, meter_numbers):
        """Execute day-wise query by processing each day separately"""
        all_results = []
        
        start_date = datetime.datetime.strptime(from_date, "%Y-%m-%d")
        end_date = datetime.datetime.strptime(to_date, "%Y-%m-%d")
        
        current_date = start_date
        while current_date <= end_date and not self.stop_processing:
            date_str = current_date.strftime("%Y-%m-%d")
            
            query = self.build_query(module, sla_type, "Day-wise", date_str, date_str, meter_numbers)
            
            if query:
                retry_count = 0
                while retry_count < self.MAX_RETRIES and not self.stop_processing:
                    try:
                        cur, conn = self.connect_db(module)
                        if cur is None:
                            if not self.wait_for_vpn(module, max_retries=10000, retry_interval=5):
                                break
                            cur, conn = self.connect_db(module)
                            if cur is None:
                                break
                        
                        day_results = self.execute_single_query(cur, query)
                        all_results.extend(day_results)
                        self.log_message(f"Data for {date_str}: {len(day_results)} records")
                        break
                        
                    except Exception as e:
                        self.log_message(f"Day-wise query failed for {date_str} (Attempt {retry_count + 1}/{self.MAX_RETRIES}): {str(e)}")
                        retry_count += 1
                        if retry_count >= self.MAX_RETRIES:
                            self.log_message(f"Max retries reached for day {date_str}")
                            break
                        all_results = [r for r in all_results if r.get('report_date', '') != date_str]
                        self.log_message(f"Waiting {self.RETRY_INTERVAL_MS/1000}s before retry...")
                        for i in range(int(self.RETRY_INTERVAL_MS/1000)):
                            if self.stop_processing:
                                break
                            time.sleep(1)
            
            current_date += datetime.timedelta(days=1)
        
        return all_results
    
    def execute_query_optimized(self, module, sla_type, data_type, from_date, to_date, meter_numbers):
        """Optimized query execution with resume capability"""
        if not meter_numbers:
            return []
        
        if module == "ALL":
            all_results = []
            all_results.extend(self.execute_query_optimized("SLA", sla_type, data_type, from_date, to_date, meter_numbers))
            all_results.extend(self.execute_query_optimized("MDM", sla_type, data_type, from_date, to_date, meter_numbers))
            all_results.extend(self.execute_query_optimized("HES", sla_type, data_type, from_date, to_date, meter_numbers))
            return all_results
        
        if sla_type == "BLP-DLP":
            combined = []
            combined.extend(self.execute_query_optimized(module, "BLP", data_type, from_date, to_date, meter_numbers))
            combined.extend(self.execute_query_optimized(module, "DLP", data_type, from_date, to_date, meter_numbers))
            return combined
        
        if sla_type == "ALL":
            combined = []
            combined.extend(self.execute_query_optimized(module, "BLP", data_type, from_date, to_date, meter_numbers))
            combined.extend(self.execute_query_optimized(module, "DLP", data_type, from_date, to_date, meter_numbers))
            combined.extend(self.execute_query_optimized(module, "BILLING", data_type, from_date, to_date, meter_numbers))
            return combined
        
        aggregated_results = []
        
        valid_meters = []
        for m in meter_numbers:
            if m and str(m).strip():
                valid_meters.append(str(m).strip())
        
        if not valid_meters:
            return aggregated_results
        
        saved_state = self.load_batch_state()
        start_batch_index = 0
        
        if saved_state and saved_state.get('module') == module and saved_state.get('sla_type') == sla_type and saved_state.get('data_type') == data_type:
            if (saved_state.get('from_date') == from_date and saved_state.get('to_date') == to_date):
                start_batch_index = saved_state.get('batch_index', 0)
                self.log_message(f"Resuming {module} {sla_type} from batch {start_batch_index + 1}/{saved_state.get('total_batches', 0)}")
            else:
                self.log_message("Saved batch state dates don't match current selection. Starting fresh.")
                self.clear_batch_state()
        
        retry_count = 0
        while retry_count < self.MAX_RETRIES and not self.stop_processing:
            try:
                cur, conn = self.connect_db(module)
                if cur is None:
                    if not self.wait_for_vpn(module, max_retries=10000, retry_interval=5):
                        break
                    cur, conn = self.connect_db(module)
                    if cur is None:
                        break
                
                total_batches = ((len(valid_meters) - 1) // self.METER_BATCH_SIZE) + 1
                
                for i in range(start_batch_index * self.METER_BATCH_SIZE, len(valid_meters), self.METER_BATCH_SIZE):
                    if self.stop_processing:
                        break
                    
                    end = min(len(valid_meters), i + self.METER_BATCH_SIZE)
                    batch = valid_meters[i:end]
                    current_batch_num = (i // self.METER_BATCH_SIZE) + 1
                    
                    try:
                        if data_type == "Day-wise" and module == "SLA":
                            batch_results = self.execute_day_wise_query(module, sla_type, from_date, to_date, batch)
                            aggregated_results.extend(batch_results)
                        else:
                            query = self.build_query(module, sla_type, data_type, from_date, to_date, batch)
                            if query:
                                batch_results = self.execute_single_query(cur, query)
                                aggregated_results.extend(batch_results)
                        
                        self.log_message(f"Processing batch {current_batch_num}/{total_batches} ({len(aggregated_results)})")
                        self.save_batch_state(module, sla_type, data_type, current_batch_num, total_batches, i + len(batch))
                        
                    except Exception as batch_error:
                        self.log_message(f"Error processing batch {current_batch_num}: {str(batch_error)}")
                        self.save_batch_state(module, sla_type, data_type, current_batch_num - 1, total_batches, i)
                        raise batch_error
                
                self.clear_batch_state()
                self.close_db_connection(module)
                return aggregated_results
                
            except Exception as e:
                self.log_message(f"[{module}] Connection or Batch failed (Attempt {retry_count + 1}/{self.MAX_RETRIES}): {str(e)}")
                retry_count += 1
                self.close_db_connection(module)
                if retry_count >= self.MAX_RETRIES:
                    self.log_message(f"Max retries reached for module {module}")
                    break
                self.log_message(f"Waiting {self.RETRY_INTERVAL_MS/1000}s before full module retry...")
                for i in range(int(self.RETRY_INTERVAL_MS/1000)):
                    if self.stop_processing:
                        break
                    time.sleep(1)
        
        return aggregated_results
    
    def get_query(self, sla_type, module):
        """Get the appropriate query based on selections and module (Manual mode)"""
        data_type = self.data_type_var.get()
        from_date = self.from_date_entry.get()
        to_date = self.to_date_entry.get()
        meter_numbers = self.get_meter_numbers()
        
        results = self.execute_query_optimized(module, sla_type, data_type, from_date, to_date, meter_numbers)
        
        if not results:
            return pd.DataFrame()
        
        return pd.DataFrame(results)
    
    def get_month_year_from_date(self, date_str):
        """Extract month and year in MMYY format from date string"""
        date_obj = datetime.datetime.strptime(date_str, "%Y-%m-%d")
        return date_obj.strftime("%m%y")
    
    def save_results(self, df, sla_type, module):
        """Save results to Excel file or CSV if too large for Excel"""
        if df is None or df.empty:
            self.log_message(f"No {sla_type} data to save.")
            return
        
        # Add SAT_No column from filtered data (only in Manual mode, when filtered_data is available)
        if self.filtered_data is not None and sla_type != "BILLING":
            meter_col_name = 'meter_number'
            
            if module == "MDM" and sla_type == "DLP":
                meter_col_name = 'mtrno'
            
            if meter_col_name in df.columns:
                sat_mapping = self.filtered_data.set_index('Meter Serial Number')['SAT_No'].to_dict()
                df['SAT_No'] = df[meter_col_name].map(sat_mapping)
            else:
                self.log_message(f"Warning: Meter column '{meter_col_name}' not found in DataFrame. Skipping SAT_No mapping.")
        
        data_type = self.data_type_var.get()
        current_date = datetime.datetime.now().strftime("%Y%m%d")
        from_date = self.from_date_entry.get()
        month_year = self.get_month_year_from_date(from_date)
        
        if sla_type == "BILLING":
            base_filename = f"BILLING_RAW_METER_{module}_{month_year}_{current_date}"
        else:
            data_type_str = "METER" if data_type == "Meter-wise" else "DAY"
            base_filename = f"{sla_type}_RAW_{data_type_str}_{module}_{month_year}_{current_date}"
        
        try:
            if len(df) > 1048576:
                filename = os.path.join(self.save_directory, f"{base_filename}.csv")
                df.to_csv(filename, index=False)
                self.log_message(f"Data too large for Excel ({len(df)} rows). {sla_type} results saved to CSV: {filename}")
            else:
                filename = os.path.join(self.save_directory, f"{base_filename}.xlsx")
                df.to_excel(filename, index=False)
                self.log_message(f"{sla_type} results saved to Excel: {filename}")
                
            # Also save to data_selve.dat if we have Excel data loaded
            if self.excel_data is not None:
                data_file = os.path.join(self.save_directory, 'data_selve.dat')
                with open(data_file, 'wb') as f:
                    pickle.dump(self.excel_data, f)
        except Exception as e:
            self.log_message(f"Error saving {sla_type} results: {str(e)}")
    
    def load_all_data_files(self):
        """Load all available data files and display record counts"""
        self.log_message("\nLoading all available data files...")
        
        data_file = os.path.join(self.save_directory, 'data_selve.dat')
        if os.path.exists(data_file):
            try:
                with open(data_file, 'rb') as f:
                    self.excel_data = pickle.load(f)
                    self.log_message(f"Loaded Excel Data: {len(self.excel_data)} records")
            except Exception as e:
                self.log_message(f"Error loading Excel data: {str(e)}")
                self.excel_data = None
        
        data_files = {
            'BLP_RAW_DAY_SLA': None,
            'BLP_RAW_METER_SLA': None,
            'DLP_RAW_DAY_SLA': None,
            'DLP_RAW_METER_SLA': None,
            'BILLING_RAW_METER_SLA': None,
            'BLP_RAW_DAY_MDM': None,
            'BLP_RAW_METER_MDM': None,
            'DLP_RAW_DAY_MDM': None,
            'DLP_RAW_METER_MDM': None,
            'BILLING_RAW_METER_MDM': None,
            'BLP_RAW_DAY_HES': None,
            'BLP_RAW_METER_HES': None,
            'DLP_RAW_DAY_HES': None,
            'DLP_RAW_METER_HES': None,
            'BILLING_RAW_METER_HES': None
        }
        
        files = os.listdir(self.save_directory)
        
        for file_type in data_files:
            matching_files = [f for f in files if f.startswith(file_type)]
            if matching_files:
                latest_file = max(matching_files, key=lambda f: os.path.getctime(os.path.join(self.save_directory, f)))
                data_files[file_type] = os.path.join(self.save_directory, latest_file)
        
        for file_type, file_path in data_files.items():
            if file_path and os.path.exists(file_path):
                try:
                    if file_path.endswith('.xlsx'):
                        data = pd.read_excel(file_path)
                    elif file_path.endswith('.csv'):
                        data = pd.read_csv(file_path)
                    else:
                        continue
                    self.log_message(f"Loaded {file_type}: {len(data)} records")
                except Exception as e:
                    self.log_message(f"Error loading {file_type}: {str(e)}")
            else:
                self.log_message(f"No file found for {file_type}")
    
    def find_latest_file(self, prefix):
        """Find the latest file with given prefix"""
        files = os.listdir(self.save_directory)
        matching_files = [f for f in files if f.startswith(prefix)]
        
        if matching_files:
            latest_file = max(matching_files, key=lambda f: os.path.getctime(os.path.join(self.save_directory, f)))
            return os.path.join(self.save_directory, latest_file)
        return None
    
    def load_raw_data(self, file_type):
        """Load raw data from file"""
        filename = self.find_latest_file(file_type)
        if filename:
            try:
                if filename.endswith('.xlsx'):
                    return pd.read_excel(filename)
                elif filename.endswith('.csv'):
                    return pd.read_csv(filename)
            except Exception as e:
                self.log_message(f"Error loading {file_type}: {str(e)}")
        return None
    
    def select_master_file(self):
        """Select master file for final summary"""
        file_path = filedialog.askopenfilename(
            title="Select Master File",
            filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
        )
        if file_path:
            self.master_file_path = file_path
            self.log_message(f"Master file selected: {file_path}")
    
    def select_blp_file(self):
        """Select BLP RAW file for final summary"""
        file_path = filedialog.askopenfilename(
            title="Select BLP RAW File",
            filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
        )
        if file_path:
            self.blp_file_path = file_path
            self.log_message(f"BLP RAW file selected: {file_path}")
    
    def select_dlp_file(self):
        """Select DLP RAW file for final summary"""
        file_path = filedialog.askopenfilename(
            title="Select DLP RAW File",
            filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
        )
        if file_path:
            self.dlp_file_path = file_path
            self.log_message(f"DLP RAW file selected: {file_path}")
    
    def select_billing_file(self):
        """Select BILLING RAW file for final summary"""
        file_path = filedialog.askopenfilename(
            title="Select BILLING RAW File",
            filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
        )
        if file_path:
            self.billing_file_path = file_path
            self.log_message(f"BILLING RAW file selected: {file_path}")
    
    def run_final_process(self):
        """Run the final summary process"""
        if self.final_summary_var.get() == "NO":
            self.log_message("Final summary process cancelled.")
            return
        
        if not all([self.master_file_path, self.blp_file_path, self.dlp_file_path, self.billing_file_path]):
            self.log_message("Error: Please select all required files for the final summary.")
            messagebox.showerror("Error", "Please select all required files for the final summary.")
            return
        
        try:
            self.log_message("Starting final summary process...")
            
            billing_filename = os.path.basename(self.billing_file_path)
            filename_parts = billing_filename.split('_')
            if len(filename_parts) >= 5:
                sla_period = filename_parts[4]
            else:
                sla_period = datetime.datetime.now().strftime("%m%y")
                self.log_message(f"Warning: Could not extract period from filename. Using current period: {sla_period}")
            
            self.log_message(f"SLA Period extracted: {sla_period}")
            
            MM = int(sla_period[:2])
            YY = int("20" + sla_period[2:])
            
            if MM in [4, 6, 9, 11]:
                NOD = 30
            elif MM == 2:
                if (YY % 4 == 0 and YY % 100 != 0) or (YY % 400 == 0):
                    NOD = 29
                else:
                    NOD = 28
            else:
                NOD = 31
            
            self.log_message(f"Month: {MM}, Year: {YY}, Number of Days: {NOD}")
            
            cutoff1 = datetime.datetime(YY, MM, 4)
            cutoff2 = datetime.datetime(YY, MM, 8)
            
            self.log_message("Loading raw data files...")
            BLP_Data = pd.read_excel(self.blp_file_path)
            DLP_Data = pd.read_excel(self.dlp_file_path)
            Bill_Data_raw = pd.read_excel(self.billing_file_path)
            MtrMaster = pd.read_excel(self.master_file_path)
            
            self.log_message("Raw data files loaded successfully. Processing data files...")
            
            Bill_Data = (
                Bill_Data_raw.sort_values(['meter_number', 'bill_date'])
                .drop_duplicates(subset='meter_number', keep='first')
                .reset_index(drop=True)
            )
            Bill_Data.loc[:, 'bill_72'] = 0
            Bill_Data.loc[:, 'bill_168'] = 0
            
            Bill_Data['bill_72'] = (
                (Bill_Data['bill_date'] < cutoff1) &
                (Bill_Data['server_time'] < cutoff1)
            ).astype(int)
            
            Bill_Data['bill_168'] = (
                (Bill_Data['bill_date'] < cutoff2) &
                (Bill_Data['server_time'] < cutoff2)
            ).astype(int)
            
            MtrMaster = MtrMaster.loc[:, ['Meter Serial Number', 'HES', 'IP', 'SAT_No']]
            
            col = ['SDO', 'MeterNo', 'HES', 'Comm Tech', 'SAT No', 'IP', 'NOD', 'BLP_Tg', 'BLP_8', 'BLP_12', 'BLP_24', '%BLP_8', '%BLP_12', '%BLP_24', 'DLP_24', '%DLP_24', 'Bill Server Time', 'Bill Time', 'Bill_3', 'Bill_7', '%Bill_3', '%Bill_7']
            SLA_Final_Raw = pd.DataFrame(columns=col)
            
            SLA_Final_Raw['HES'] = MtrMaster['HES']
            SLA_Final_Raw['IP'] = MtrMaster['IP']
            SLA_Final_Raw['SAT No'] = MtrMaster['SAT_No']
            SLA_Final_Raw['MeterNo'] = MtrMaster['Meter Serial Number'].astype(str)
            SLA_Final_Raw['NOD'] = NOD
            SLA_Final_Raw['BLP_Tg'] = SLA_Final_Raw['NOD'] * SLA_Final_Raw['IP']
            
            BLP_Data.set_index("meter_number", inplace=True)
            SLA_Final_Raw["BLP_8"] = SLA_Final_Raw.MeterNo.map(BLP_Data["h8"])
            SLA_Final_Raw["BLP_12"] = SLA_Final_Raw.MeterNo.map(BLP_Data["h12"])
            SLA_Final_Raw["BLP_24"] = SLA_Final_Raw.MeterNo.map(BLP_Data["h24"])
            
            SLA_Final_Raw["%BLP_8"] = SLA_Final_Raw["BLP_8"] * 100 / SLA_Final_Raw['BLP_Tg']
            SLA_Final_Raw["%BLP_12"] = SLA_Final_Raw["BLP_12"] * 100 / SLA_Final_Raw['BLP_Tg']
            SLA_Final_Raw["%BLP_24"] = SLA_Final_Raw["BLP_24"] * 100 / SLA_Final_Raw['BLP_Tg']
            
            DLP_Data.set_index("meter_number", inplace=True)
            SLA_Final_Raw["DLP_24"] = SLA_Final_Raw.MeterNo.map(DLP_Data["h24"])
            SLA_Final_Raw["%DLP_24"] = SLA_Final_Raw["DLP_24"] * 100 / SLA_Final_Raw['NOD']
            
            Bill_Data.set_index("meter_number", inplace=True)
            SLA_Final_Raw["Bill_3"] = SLA_Final_Raw.MeterNo.map(Bill_Data["bill_72"])
            SLA_Final_Raw["Bill_7"] = SLA_Final_Raw.MeterNo.map(Bill_Data["bill_168"])
            SLA_Final_Raw["%Bill_3"] = SLA_Final_Raw["Bill_3"] * 100
            SLA_Final_Raw['%Bill_7'] = SLA_Final_Raw["Bill_7"] * 100
            
            SLA_Summ_HES_SAT_Final = (
                SLA_Final_Raw.groupby(['HES', 'SAT No'], as_index=False)
                .agg({'MeterNo': 'count', 'BLP_Tg': 'sum', 'BLP_8': 'sum', 'BLP_12': 'sum', 'BLP_24': 'sum', 'NOD': 'sum', 'DLP_24': 'sum', 'Bill_3': 'sum', 'Bill_7': 'sum'})
                .rename(columns={'NOD': 'DLP_Tg', 'MeterNo': 'Mtr_Qty'})
            )
            
            SLA_Summ_HES_SAT_Final['%BLP_8'] = SLA_Summ_HES_SAT_Final['BLP_8'] * 100 / SLA_Summ_HES_SAT_Final['BLP_Tg']
            SLA_Summ_HES_SAT_Final['%BLP_12'] = SLA_Summ_HES_SAT_Final['BLP_12'] * 100 / SLA_Summ_HES_SAT_Final['BLP_Tg']
            SLA_Summ_HES_SAT_Final['%BLP_24'] = SLA_Summ_HES_SAT_Final['BLP_24'] * 100 / SLA_Summ_HES_SAT_Final['BLP_Tg']
            SLA_Summ_HES_SAT_Final['%DLP_24'] = SLA_Summ_HES_SAT_Final['DLP_24'] * 100 / SLA_Summ_HES_SAT_Final['DLP_Tg']
            SLA_Summ_HES_SAT_Final['%Bill_3'] = SLA_Summ_HES_SAT_Final['Bill_3'] * 100 / SLA_Summ_HES_SAT_Final['Mtr_Qty']
            SLA_Summ_HES_SAT_Final['%Bill_7'] = SLA_Summ_HES_SAT_Final['Bill_7'] * 100 / SLA_Summ_HES_SAT_Final['Mtr_Qty']
            
            SLA_Summ_HES_SAT_Final['BLP_8'] = SLA_Summ_HES_SAT_Final['BLP_8'].round(0).astype(int)
            SLA_Summ_HES_SAT_Final['BLP_12'] = SLA_Summ_HES_SAT_Final['BLP_12'].round(0).astype(int)
            SLA_Summ_HES_SAT_Final['BLP_24'] = SLA_Summ_HES_SAT_Final['BLP_24'].round(0).astype(int)
            SLA_Summ_HES_SAT_Final['DLP_24'] = SLA_Summ_HES_SAT_Final['DLP_24'].round(0).astype(int)
            SLA_Summ_HES_SAT_Final['%BLP_8'] = SLA_Summ_HES_SAT_Final['%BLP_8'].round(3)
            SLA_Summ_HES_SAT_Final['%BLP_12'] = SLA_Summ_HES_SAT_Final['%BLP_12'].round(3)
            SLA_Summ_HES_SAT_Final['%BLP_24'] = SLA_Summ_HES_SAT_Final['%BLP_24'].round(3)
            SLA_Summ_HES_SAT_Final['%DLP_24'] = SLA_Summ_HES_SAT_Final['%DLP_24'].round(3)
            SLA_Summ_HES_SAT_Final['%Bill_3'] = SLA_Summ_HES_SAT_Final['%Bill_3'].round(3)
            SLA_Summ_HES_SAT_Final['%Bill_7'] = SLA_Summ_HES_SAT_Final['%Bill_7'].round(3)
            
            SLA_Summ_HES_Final = (
                SLA_Final_Raw.groupby(['HES'], as_index=False)
                .agg({'MeterNo': 'count', 'BLP_Tg': 'sum', 'BLP_8': 'sum', 'BLP_12': 'sum', 'BLP_24': 'sum', 'NOD': 'sum', 'DLP_24': 'sum', 'Bill_3': 'sum', 'Bill_7': 'sum'})
                .rename(columns={'NOD': 'DLP_Tg', 'MeterNo': 'Mtr_Qty'})
            )
            
            all_hes_row = {
                'HES': 'All HES',
                'Mtr_Qty': SLA_Summ_HES_Final['Mtr_Qty'].sum().round(0).astype(int),
                'BLP_Tg': SLA_Summ_HES_Final['BLP_Tg'].sum().round(0).astype(int),
                'BLP_8': SLA_Summ_HES_Final['BLP_8'].sum().round(0).astype(int),
                'BLP_12': SLA_Summ_HES_Final['BLP_12'].sum().round(0).astype(int),
                'BLP_24': SLA_Summ_HES_Final['BLP_24'].sum().round(0).astype(int),
                'DLP_Tg': SLA_Summ_HES_Final['DLP_Tg'].sum().round(0).astype(int),
                'DLP_24': SLA_Summ_HES_Final['DLP_24'].sum().round(0).astype(int),
                'Bill_3': SLA_Summ_HES_Final['Bill_3'].sum().round(0).astype(int),
                'Bill_7': SLA_Summ_HES_Final['Bill_7'].sum().round(0).astype(int)
            }
            
            SLA_Summ_HES_Final = pd.concat([SLA_Summ_HES_Final, pd.DataFrame([all_hes_row])], ignore_index=True)
            
            SLA_Summ_HES_Final['%BLP_8'] = SLA_Summ_HES_Final['BLP_8'] * 100 / SLA_Summ_HES_Final['BLP_Tg']
            SLA_Summ_HES_Final['%BLP_12'] = SLA_Summ_HES_Final['BLP_12'] * 100 / SLA_Summ_HES_Final['BLP_Tg']
            SLA_Summ_HES_Final['%BLP_24'] = SLA_Summ_HES_Final['BLP_24'] * 100 / SLA_Summ_HES_Final['BLP_Tg']
            SLA_Summ_HES_Final['%DLP_24'] = SLA_Summ_HES_Final['DLP_24'] * 100 / SLA_Summ_HES_Final['DLP_Tg']
            SLA_Summ_HES_Final['%Bill_3'] = SLA_Summ_HES_Final['Bill_3'] * 100 / SLA_Summ_HES_Final['Mtr_Qty']
            SLA_Summ_HES_Final['%Bill_7'] = SLA_Summ_HES_Final['Bill_7'] * 100 / SLA_Summ_HES_Final['Mtr_Qty']
            
            SLA_Summ_HES_Final['Mtr_Qty'] = SLA_Summ_HES_Final['Mtr_Qty'].round(0).astype(int)
            SLA_Summ_HES_Final['BLP_8'] = SLA_Summ_HES_Final['BLP_8'].round(0).astype(int)
            SLA_Summ_HES_Final['BLP_12'] = SLA_Summ_HES_Final['BLP_12'].round(0).astype(int)
            SLA_Summ_HES_Final['BLP_24'] = SLA_Summ_HES_Final['BLP_24'].round(0).astype(int)
            SLA_Summ_HES_Final['DLP_24'] = SLA_Summ_HES_Final['DLP_24'].round(0).astype(int)
            SLA_Summ_HES_Final['BLP_Tg'] = SLA_Summ_HES_Final['BLP_Tg'].round(0).astype(int)
            SLA_Summ_HES_Final['DLP_Tg'] = SLA_Summ_HES_Final['DLP_Tg'].round(0).astype(int)
            SLA_Summ_HES_Final['Bill_3'] = SLA_Summ_HES_Final['Bill_3'].round(0).astype(int)
            SLA_Summ_HES_Final['Bill_7'] = SLA_Summ_HES_Final['Bill_7'].round(0).astype(int)
            SLA_Summ_HES_Final['%BLP_8'] = SLA_Summ_HES_Final['%BLP_8'].round(3)
            SLA_Summ_HES_Final['%BLP_12'] = SLA_Summ_HES_Final['%BLP_12'].round(3)
            SLA_Summ_HES_Final['%BLP_24'] = SLA_Summ_HES_Final['%BLP_24'].round(3)
            SLA_Summ_HES_Final['%DLP_24'] = SLA_Summ_HES_Final['%DLP_24'].round(3)
            SLA_Summ_HES_Final['%Bill_3'] = SLA_Summ_HES_Final['%Bill_3'].round(3)
            SLA_Summ_HES_Final['%Bill_7'] = SLA_Summ_HES_Final['%Bill_7'].round(3)
            
            self.log_message("Data processed, saving data in Excel...")
            
            file_path = os.path.join(self.save_directory, f'SLA_Cumm_Worksheet_Final_{sla_period}_{datetime.datetime.now().day}{datetime.datetime.now().month}{datetime.datetime.now().year}.xlsx')
            
            with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                SLA_Summ_HES_Final.to_excel(writer, sheet_name='HES_Wise_SLA_Summ', index=False)
                SLA_Summ_HES_SAT_Final.to_excel(writer, sheet_name='SAT_Wise_SLA_Summ', index=True)
                SLA_Final_Raw.to_excel(writer, sheet_name='SLA_Final_Raw', index=True)
                BLP_Data.to_excel(writer, sheet_name='BLP_Data', index=True)
                DLP_Data.to_excel(writer, sheet_name='DLP_Data', index=True)
                Bill_Data.to_excel(writer, sheet_name='Bill_Data', index=True)
                MtrMaster.to_excel(writer, sheet_name='MtrMaster', index=True)
            
            self.log_message(f"Data saved successfully in Excel: {file_path}")
            self.log_message("Final summary process completed successfully!")
            
        except Exception as e:
            self.log_message(f"Error in final summary process: {str(e)}")
            messagebox.showerror("Error", f"Failed to run final summary process: {str(e)}")
    
    def run_process(self):
        """Main process to run based on selections"""
        is_auto = self.fetch_mode_var.get() == "Auto"

        if self.fetch_var.get() == "YES" and not self.vpn_connected:
            self.log_message("VPN not connected. Please check VPN first.")
            messagebox.showwarning("VPN Not Connected", "Please check VPN connection before fetching data.")
            self.processing = False
            self.run_button.grid(row=5, column=0, columnspan=3, pady=20, sticky=tk.EW)
            self.stop_button.grid_remove()
            return

        # ── AUTO MODE ─────────────────────────────────────────────────────────
        if is_auto:
            module = self.db_type_var.get()
            sla_type = self.sla_type_var.get()

            self.log_message(f"[Auto Mode] Starting data fetch... Module={module}, Profile={sla_type}")

            # FIRST: Fetch master data from SLA database
            master_df = self.fetch_auto_master_data()
            if master_df is None or master_df.empty:
                self.log_message("Warning: No master data fetched. Continuing with queries...")
            else:
                # Update filtered_data with master data for SAT mapping
                self.filtered_data = master_df.copy()
                # Rename columns to match expected format for mapping
                self.filtered_data.rename(columns={
                    'meter_serial_number': 'Meter Serial Number',
                    'hes': 'HES',
                    'ip': 'IP',
                    'Sat_No': 'SAT_No'
                }, inplace=True)

            # For ALL module: process SLA then MDM only (no HES in Auto mode)
            modules_to_process = ["SLA", "MDM"] if module == "ALL" else [module]

            for mod in modules_to_process:
                if self.stop_processing:
                    break

                self.log_message(f"\n[Auto] Processing {mod} module...")

                # Check/wait for VPN
                self.db_type_var.set(mod)
                self.check_vpn()
                if not self.vpn_connected:
                    self.log_message(f"VPN not connected for {mod}. Waiting...")
                    if not self.wait_for_vpn(mod, max_retries=10000, retry_interval=10):
                        self.log_message(f"Failed to connect for {mod}. Skipping...")
                        continue
                self.db_type_var.set(module)  # restore original selection

                if sla_type == "ALL":
                    # BILLING first, then BLP and DLP
                    self.log_message(f"\n[Auto] Processing {mod} BILLING...")
                    billing_df = self.get_query_auto("BILLING", mod)
                    if billing_df is not None and not billing_df.empty:
                        self.save_results_auto(billing_df, "BILLING", mod)

                    for st in ["BLP", "DLP"]:
                        if self.stop_processing:
                            break
                        self.log_message(f"\n[Auto] Processing {mod} {st}...")
                        result_df = self.get_query_auto(st, mod)
                        if result_df is not None and not result_df.empty:
                            self.save_results_auto(result_df, st, mod)

                elif sla_type == "BILLING":
                    self.log_message(f"\n[Auto] Processing {mod} BILLING...")
                    billing_df = self.get_query_auto("BILLING", mod)
                    if billing_df is not None and not billing_df.empty:
                        self.save_results_auto(billing_df, "BILLING", mod)

                elif sla_type == "BLP-DLP":
                    for st in ["BLP", "DLP"]:
                        if self.stop_processing:
                            break
                        self.log_message(f"\n[Auto] Processing {mod} {st}...")
                        result_df = self.get_query_auto(st, mod)
                        if result_df is not None and not result_df.empty:
                            self.save_results_auto(result_df, st, mod)

                else:
                    self.log_message(f"\n[Auto] Processing {mod} {sla_type}...")
                    result_df = self.get_query_auto(sla_type, mod)
                    if result_df is not None and not result_df.empty:
                        self.save_results_auto(result_df, sla_type, mod)

            if not self.stop_processing:
                self.log_message("\n[Auto] Process completed.")
                if module in ["ALL", "SLA"]:
                    self.log_message("\nDo you want final Cummulative Summary..?")
                    self.summary_frame.grid(row=6, column=0, columnspan=3, pady=10, sticky=tk.EW)

            self.processing = False
            self.run_button.grid(row=5, column=0, columnspan=3, pady=20, sticky=tk.EW)
            self.stop_button.grid_remove()
            return

        # ── MANUAL MODE ───────────────────────────────────────────────────────
        if not self.vpn_connected and self.fetch_var.get() == "YES":
            self.log_message("VPN not connected. Please check VPN first.")
            self.processing = False
            self.run_button.grid(row=5, column=0, columnspan=3, pady=20, sticky=tk.EW)
            self.stop_button.grid_remove()
            return
        
        self.filter_data()
        
        if self.filtered_data is None or self.filtered_data.empty:
            self.log_message("No data to process after filtering.")
            self.processing = False
            self.run_button.grid(row=5, column=0, columnspan=3, pady=20, sticky=tk.EW)
            self.stop_button.grid_remove()
            return
        
        module = self.db_type_var.get()
        sla_type = self.sla_type_var.get()
        
        if self.fetch_var.get() == "YES":
            self.log_message("Starting data fetch process...")
            
            if module == "ALL":
                modules_to_process = ["SLA", "MDM", "HES"]
                
                for mod in modules_to_process:
                    if self.stop_processing:
                        break
                        
                    self.log_message(f"\nProcessing {mod} module...")
                    
                    self.db_type_var.set(mod)
                    self.check_vpn()
                    
                    if not self.vpn_connected:
                        self.log_message(f"VPN not connected for {mod} module. Waiting for connection...")
                        if not self.wait_for_vpn(mod, max_retries=10000, retry_interval=10):
                            self.log_message(f"Failed to connect to VPN for {mod} module. Skipping...")
                            continue
                    
                    self.log_message(f"\nProcessing {mod} BILLING data first...")
                    billing_result_df = self.get_query("BILLING", mod)
                    if billing_result_df is not None and not billing_result_df.empty:
                        self.save_results(billing_result_df, "BILLING", mod)
                    
                    for st in ["BLP", "DLP"]:
                        if self.stop_processing:
                            break
                        self.log_message(f"\nProcessing {mod} {st} data...")
                        result_df = self.get_query(st, mod)
                        if result_df is not None and not result_df.empty:
                            self.save_results(result_df, st, mod)
                
                if not self.stop_processing:
                    self.log_message("\nProcess completed for ALL modules (SLA, MDM, HES)")
                    self.log_message("\nDo you want final Cummulative Summary..?")
                    self.summary_frame.grid(row=6, column=0, columnspan=3, pady=10, sticky=tk.EW)
                
                self.processing = False
                self.run_button.grid(row=5, column=0, columnspan=3, pady=20, sticky=tk.EW)
                self.stop_button.grid_remove()
                return
            
            else:
                self.log_message(f"\nProcessing {module} module...")
                
                if sla_type == "ALL":
                    self.log_message(f"\nProcessing {module} BILLING data first...")
                    billing_result_df = self.get_query("BILLING", module)
                    if billing_result_df is not None and not billing_result_df.empty:
                        self.save_results(billing_result_df, "BILLING", module)
                    
                    for st in ["BLP", "DLP"]:
                        if self.stop_processing:
                            break
                        self.log_message(f"\nProcessing {module} {st} data...")
                        result_df = self.get_query(st, module)
                        if result_df is not None and not result_df.empty:
                            self.save_results(result_df, st, module)
                    
                    if not self.stop_processing:
                        self.log_message(f"\nProcess completed for {module} ALL selection (BILLING, BLP, DLP)")
                        if module in ["ALL", "SLA"]:
                            self.log_message("\nDo you want final Cummulative Summary..?")
                            self.summary_frame.grid(row=6, column=0, columnspan=3, pady=10, sticky=tk.EW)
                    
                    self.processing = False
                    self.run_button.grid(row=5, column=0, columnspan=3, pady=20, sticky=tk.EW)
                    self.stop_button.grid_remove()
                    return
                
                elif sla_type == "BILLING":
                    self.log_message(f"\nProcessing {module} BILLING data...")
                    billing_result_df = self.get_query("BILLING", module)
                    if billing_result_df is not None and not billing_result_df.empty:
                        self.save_results(billing_result_df, "BILLING", module)
                        self.log_message(f"\nProcess completed for {module} BILLING selection")
                        if module in ["ALL", "SLA"]:
                            self.log_message("\nDo you want final Cummulative Summary..?")
                            self.summary_frame.grid(row=6, column=0, columnspan=3, pady=10, sticky=tk.EW)
                        self.processing = False
                        self.run_button.grid(row=5, column=0, columnspan=3, pady=20, sticky=tk.EW)
                        self.stop_button.grid_remove()
                        return
                    
                    self.log_message(f"No {module} BILLING data found or query failed")
                    self.processing = False
                    self.run_button.grid(row=5, column=0, columnspan=3, pady=20, sticky=tk.EW)
                    self.stop_button.grid_remove()
                    return
                
                elif sla_type == "BLP-DLP":
                    for st in ["BLP", "DLP"]:
                        if self.stop_processing:
                            break
                        self.log_message(f"\nProcessing {module} {st} data...")
                        result_df = self.get_query(st, module)
                        if result_df is not None and not result_df.empty:
                            self.save_results(result_df, st, module)
                    
                    if not self.stop_processing:
                        if module in ["ALL", "SLA"]:
                            self.log_message("\nDo you want final Cummulative Summary..?")
                            self.summary_frame.grid(row=6, column=0, columnspan=3, pady=10, sticky=tk.EW)
                    
                    self.processing = False
                    self.run_button.grid(row=5, column=0, columnspan=3, pady=20, sticky=tk.EW)
                    self.stop_button.grid_remove()
                    return
                
                else:
                    result_df = self.get_query(sla_type, module)
                    if result_df is not None and not result_df.empty:
                        self.save_results(result_df, sla_type, module)
                    
                    if not self.stop_processing:
                        if module in ["ALL", "SLA"]:
                            self.log_message("\nDo you want final Cummulative Summary..?")
                            self.summary_frame.grid(row=6, column=0, columnspan=3, pady=10, sticky=tk.EW)
                    
                    self.processing = False
                    self.run_button.grid(row=5, column=0, columnspan=3, pady=20, sticky=tk.EW)
                    self.stop_button.grid_remove()
                    return
        
        self.processing = False
        self.run_button.grid(row=5, column=0, columnspan=3, pady=20, sticky=tk.EW)
        self.stop_button.grid_remove()

    def save_results_auto(self, df, sla_type, module):
        """Save results for Auto mode - uses auto master data for SAT mapping"""
        if df is None or df.empty:
            self.log_message(f"No {sla_type} data to save.")
            return
        
        # Add SAT_No column from auto master data
        if sla_type != "BILLING":
            meter_col_name = 'meter_number'
            if module == "MDM" and sla_type == "DLP":
                meter_col_name = 'mtrno'
            
            if meter_col_name in df.columns and self.auto_master_data is not None:
                sat_mapping = self.auto_master_data.set_index('meter_serial_number')['Sat_No'].to_dict()
                df['SAT_No'] = df[meter_col_name].map(sat_mapping)
            elif self.auto_master_data is None and self.filtered_data is not None:
                # Fallback to filtered_data if available
                sat_mapping = self.filtered_data.set_index('Meter Serial Number')['SAT_No'].to_dict()
                if meter_col_name in df.columns:
                    df['SAT_No'] = df[meter_col_name].map(sat_mapping)
            else:
                self.log_message(f"Warning: No master data available for SAT mapping.")
        
        data_type = self.data_type_var.get()
        current_date = datetime.datetime.now().strftime("%Y%m%d")
        from_date = self.from_date_entry.get()
        month_year = self.get_month_year_from_date(from_date)
        
        if sla_type == "BILLING":
            base_filename = f"BILLING_RAW_METER_{module}_{month_year}_{current_date}"
        else:
            data_type_str = "METER" if data_type == "Meter-wise" else "DAY"
            base_filename = f"{sla_type}_RAW_{data_type_str}_{module}_{month_year}_{current_date}"
        
        try:
            if len(df) > 1048576:
                filename = os.path.join(self.save_directory, f"{base_filename}.csv")
                df.to_csv(filename, index=False)
                self.log_message(f"Data too large for Excel ({len(df)} rows). {sla_type} results saved to CSV: {filename}")
            else:
                filename = os.path.join(self.save_directory, f"{base_filename}.xlsx")
                df.to_excel(filename, index=False)
                self.log_message(f"{sla_type} results saved to Excel: {filename}")
        except Exception as e:
            self.log_message(f"Error saving {sla_type} results: {str(e)}")