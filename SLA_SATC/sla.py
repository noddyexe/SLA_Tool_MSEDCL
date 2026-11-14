import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import psycopg2
import pandas as pd
import numpy as np
import os
import datetime
import socket
import sys
import time
import threading
import re
from concurrent.futures import ThreadPoolExecutor
from config_loader import get_config
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def _select_excel_engine():
    """Use python-calamine (Rust, ~20x faster) when installed, otherwise openpyxl."""
    try:
        import python_calamine  # noqa: F401
        return 'calamine'
    except ImportError:
        return 'openpyxl'


EXCEL_ENGINE = _select_excel_engine()


def read_all_sheets(path):
    """Read every sheet of an Excel file into a dict of DataFrames."""
    return pd.read_excel(path, sheet_name=None, engine=EXCEL_ENGINE)


def _to_str(series):
    """Same result as str(value) for every element (missing values become 'nan', like the original)."""
    return series.astype(object).map(str)


def _obj_array(values):
    arr = np.empty(len(values), dtype=object)
    arr[:] = values
    return arr


class SLAExtractionApp:
    def __init__(self, root, go_back_callback, save_directory, is_exe):
        self.root = root
        self.go_back_callback = go_back_callback
        self.save_directory = save_directory
        self.is_exe = is_exe
        self.root.title("SLA Extraction")
        self.root.geometry("1200x700")
        self.root.configure(bg='white')
        
        # Configuration for retries and batch processing
        self.MAX_RETRIES = 500000
        self.RETRY_INTERVAL_MS = 5000
        self.BATCH_SIZE = 1000
        
        # Database connection
        self.db_connection = None
        self.rc_dc_db_connection = None  # Separate connection for RC_DC
        
        # Style configuration
        self.style = ttk.Style()
        self.style.configure('TFrame', background='MediumAquamarine')
        self.style.configure('TLabel', background='white', font=('Arial', 10))
        self.style.configure('TRadiobutton', background='white')
        

        # Main frames
        # --- Scrollable container for the left button panel ---
        self.button_frame_container = tk.Frame(root, width=500, bg='MediumAquamarine')
        self.button_frame_container.pack(side=tk.LEFT, fill=tk.Y, expand=False)
        self.button_frame_container.pack_propagate(False)

        self.button_canvas = tk.Canvas(self.button_frame_container, width=500,
                                        bg='MediumAquamarine', highlightthickness=0)
        self.button_scrollbar = ttk.Scrollbar(self.button_frame_container, orient="vertical",
                                               command=self.button_canvas.yview)
        self.button_canvas.configure(yscrollcommand=self.button_scrollbar.set)

        self.button_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.button_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.button_frame = ttk.Frame(self.button_canvas, width=500, padding="10")
        self.button_frame_window = self.button_canvas.create_window((0, 0), window=self.button_frame, anchor="nw")

        def _on_button_frame_configure(event):
            self.button_canvas.configure(scrollregion=self.button_canvas.bbox("all"))
        self.button_frame.bind("<Configure>", _on_button_frame_configure)

        def _on_canvas_configure(event):
            self.button_canvas.itemconfig(self.button_frame_window, width=event.width)
        self.button_canvas.bind("<Configure>", _on_canvas_configure)

        # Mouse wheel scrolling (Windows)
        def _on_mousewheel(event):
            self.button_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        def _bind_mousewheel(event):
            self.button_canvas.bind_all("<MouseWheel>", _on_mousewheel)
        def _unbind_mousewheel(event):
            self.button_canvas.unbind_all("<MouseWheel>")
        self.button_canvas.bind("<Enter>", _bind_mousewheel)
        self.button_canvas.bind("<Leave>", _unbind_mousewheel)
        
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
        self.vpn_connected = False
        self.processing = False
        self.stop_processing = False
        self.from_date_selected = None
        self.to_date_selected = None
        self.month_year_from_filename = None
        self.exclusion_enabled = True
        
        # File paths for final summary
        self.satc_file_path = ""
        self.blp_file_path = ""
        self.dlp_file_path = ""
        self.billing_file_path = ""
        self.pdrc_file_path = ""
        self.sm_to_sm_file_path = ""
        self.non_opp_file_path = ""
        self.power_off_file_path = ""
        self.rc_dc_file_path = ""
        
        # Get the directory where the script is located
        if getattr(sys, 'frozen', False):
            self.script_dir = os.path.dirname(sys.executable)
        else:
            self.script_dir = os.path.dirname(os.path.abspath(__file__))
        
        # Path to ami_loc.xlsx
        self.ami_loc_path = os.path.join(self.script_dir, 'ami_loc.xlsx')
        
        # Load ami_loc data if it exists
        self.ami_loc_df = None
        self.load_ami_loc_data()
    
    def load_ami_loc_data(self):
        """Load ami_loc.xlsx file from the script directory"""
        try:
            if os.path.exists(self.ami_loc_path):
                self.ami_loc_df = pd.read_excel(self.ami_loc_path)
                self.log_message(f"Loaded ami_loc.xlsx from: {self.ami_loc_path}")
                self.log_message(f"  Found {len(self.ami_loc_df)} records in ami_loc")
            else:
                self.log_message(f"Warning: ami_loc.xlsx not found at {self.ami_loc_path}")
                self.ami_loc_df = None
        except Exception as e:
            self.log_message(f"Error loading ami_loc.xlsx: {str(e)}")
            self.ami_loc_df = None
    
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
        title_label = ttk.Label(self.button_frame, text="SLA Data Extraction", 
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
        
        # Exclusion Option (YES/NO) - added after Sync Data
        exclusion_label = ttk.Label(self.options_frame, text="Exclusion:")
        exclusion_label.grid(row=0, column=0, sticky=tk.W, pady=5, padx=5)
        
        self.exclusion_var = tk.StringVar(value="YES")
        exclusion_yes = ttk.Radiobutton(self.options_frame, text="YES", variable=self.exclusion_var, 
                                       value="YES", command=self.toggle_exclusion_options)
        exclusion_yes.grid(row=0, column=1, sticky=tk.W, pady=5)
        
        exclusion_no = ttk.Radiobutton(self.options_frame, text="NO", variable=self.exclusion_var, 
                                      value="NO", command=self.toggle_exclusion_options)
        exclusion_no.grid(row=0, column=2, sticky=tk.W, pady=5)
        
        # Date Selection
        date_frame = ttk.Frame(self.options_frame)
        date_frame.grid(row=1, column=0, columnspan=3, pady=5, sticky=tk.EW)
        
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
        
        self.satc_button = tk.Button(self.file_selection_frame, text="Select SATC RAW File", 
                                   command=self.select_satc_file,
                                   bg='#2196F3', fg='white',
                                   font=('Arial', 10, 'bold'),
                                   bd=0, padx=20, pady=10,
                                   activebackground='#42A5F5',
                                   activeforeground='white',
                                   state=tk.DISABLED)
        self.satc_button.grid(row=0, column=0, columnspan=3, pady=5, sticky=tk.EW)
        
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
        
        self.pdrc_button = tk.Button(self.file_selection_frame, text="Select PDC RAW File", 
                                   command=self.select_pdrc_file,
                                   bg='#2196F3', fg='white',
                                   font=('Arial', 10, 'bold'),
                                   bd=0, padx=20, pady=10,
                                   activebackground='#42A5F5',
                                   activeforeground='white',
                                   state=tk.DISABLED)
        self.pdrc_button.grid(row=4, column=0, columnspan=3, pady=5, sticky=tk.EW)
        
        self.sm_to_sm_button = tk.Button(self.file_selection_frame, text="Select SM to SM RAW File", 
                                       command=self.select_sm_to_sm_file,
                                       bg='#2196F3', fg='white',
                                       font=('Arial', 10, 'bold'),
                                       bd=0, padx=20, pady=10,
                                       activebackground='#42A5F5',
                                       activeforeground='white',
                                       state=tk.DISABLED)
        self.sm_to_sm_button.grid(row=5, column=0, columnspan=3, pady=5, sticky=tk.EW)
        
        self.non_opp_button = tk.Button(self.file_selection_frame, text="Select Non Opp RAW File", 
                                      command=self.select_non_opp_file,
                                      bg='#2196F3', fg='white',
                                      font=('Arial', 10, 'bold'),
                                      bd=0, padx=20, pady=10,
                                      activebackground='#42A5F5',
                                      activeforeground='white',
                                      state=tk.DISABLED)
        self.non_opp_button.grid(row=6, column=0, columnspan=3, pady=5, sticky=tk.EW)
        
        self.power_off_button = tk.Button(self.file_selection_frame, text="Select Power OFF RAW File", 
                                       command=self.select_power_off_file,
                                       bg='#2196F3', fg='white',
                                       font=('Arial', 10, 'bold'),
                                       bd=0, padx=20, pady=10,
                                       activebackground='#42A5F5',
                                       activeforeground='white',
                                       state=tk.DISABLED)
        self.power_off_button.grid(row=7, column=0, columnspan=3, pady=5, sticky=tk.EW)
        
        # RC_DC button
        self.rc_dc_button = tk.Button(self.file_selection_frame, text="Select RC_DC RAW File", 
                                    command=self.select_rc_dc_file,
                                    bg='#2196F3', fg='white',
                                    font=('Arial', 10, 'bold'),
                                    bd=0, padx=20, pady=10,
                                    activebackground='#42A5F5',
                                    activeforeground='white',
                                    state=tk.DISABLED)
        self.rc_dc_button.grid(row=8, column=0, columnspan=3, pady=5, sticky=tk.EW)
        
        self.run_final_button = tk.Button(self.summary_frame, text="Run Final Process", 
                                       command=self.run_final_process,
                                       bg='#FF5722', fg='white',
                                       font=('Arial', 10, 'bold'),
                                       bd=0, padx=20, pady=10,
                                       activebackground='#E64A19',
                                       activeforeground='white',
                                       state=tk.DISABLED)
        self.run_final_button.grid(row=9, column=0, columnspan=3, pady=20, sticky=tk.EW)
        
        self.run_final_button.bind("<Enter>", lambda e: self.run_final_button.config(bg='#E64A19'))
        self.run_final_button.bind("<Leave>", lambda e: self.run_final_button.config(bg='#FF5722'))
        
        # Initially hide options frame if fetch is NO
        self.toggle_fetch_options()
        
        # Set exclusion based on default
        self.toggle_exclusion_options()
    
    def toggle_exclusion_options(self):
        """Handle exclusion option selection"""
        if self.exclusion_var.get() == "YES":
            self.exclusion_enabled = True
            self.log_message("Exclusion: ENABLED")
        else:
            self.exclusion_enabled = False
            self.log_message("Exclusion: DISABLED")
            self.log_message("Note: PDC, SM to SM, Non-Operational, and RC_DC queries will be skipped during fetch.")
            self.log_message("Note: Power OFF, RC_DC and SM to SM will still run.")
            self.log_message("Note: Power OFF query will still run")
    
    def start_process_thread(self):
        """Start the process in a separate thread to keep GUI responsive"""
        if not self.processing:
            self.processing = True
            self.stop_processing = False
            self.run_button.grid_remove()
            self.stop_button.grid(row=5, column=0, columnspan=3, pady=20, sticky=tk.EW)
            
            # Store from_date and to_date for final summary
            self.from_date_selected = self.from_date_entry.get()
            self.to_date_selected = self.to_date_entry.get()
            
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
    
    def toggle_file_selection_buttons(self):
        """Enable/disable file selection buttons based on final summary selection"""
        if self.final_summary_var.get() == "YES":
            self.file_selection_frame.grid(row=2, column=0, columnspan=3, pady=5, sticky=tk.EW)
            
            # Always show Power OFF button regardless of exclusion setting
            self.power_off_button.grid(row=7, column=0, columnspan=3, pady=5, sticky=tk.EW)
            
            if not self.exclusion_enabled:
                # Hide exclusion-related buttons (except Power OFF, RC_DC, SM to SM)
                self.pdrc_button.grid_remove()
                self.non_opp_button.grid_remove()
                
                # Show only required buttons + Power OFF + RC_DC + SM to SM
                self.satc_button.grid(row=0, column=0, columnspan=3, pady=5, sticky=tk.EW)
                self.blp_button.grid(row=1, column=0, columnspan=3, pady=5, sticky=tk.EW)
                self.dlp_button.grid(row=2, column=0, columnspan=3, pady=5, sticky=tk.EW)
                self.billing_button.grid(row=3, column=0, columnspan=3, pady=5, sticky=tk.EW)
                
                # Power OFF is already at row=7
                # RC_DC and SM to SM below Power OFF
                self.rc_dc_button.grid(row=8, column=0, columnspan=3, pady=5, sticky=tk.EW)
                self.sm_to_sm_button.grid(row=9, column=0, columnspan=3, pady=5, sticky=tk.EW)
                
                # Enable only visible buttons
                self.satc_button.config(state=tk.NORMAL)
                self.blp_button.config(state=tk.NORMAL)
                self.dlp_button.config(state=tk.NORMAL)
                self.billing_button.config(state=tk.NORMAL)
                self.power_off_button.config(state=tk.NORMAL)
                self.rc_dc_button.config(state=tk.NORMAL)
                self.sm_to_sm_button.config(state=tk.NORMAL)
            else:
                # Show all buttons
                self.pdrc_button.grid(row=4, column=0, columnspan=3, pady=5, sticky=tk.EW)
                self.sm_to_sm_button.grid(row=5, column=0, columnspan=3, pady=5, sticky=tk.EW)
                self.non_opp_button.grid(row=6, column=0, columnspan=3, pady=5, sticky=tk.EW)
                self.power_off_button.grid(row=7, column=0, columnspan=3, pady=5, sticky=tk.EW)
                self.rc_dc_button.grid(row=8, column=0, columnspan=3, pady=5, sticky=tk.EW)
                
                self.satc_button.config(state=tk.NORMAL)
                self.blp_button.config(state=tk.NORMAL)
                self.dlp_button.config(state=tk.NORMAL)
                self.billing_button.config(state=tk.NORMAL)
                self.pdrc_button.config(state=tk.NORMAL)
                self.sm_to_sm_button.config(state=tk.NORMAL)
                self.non_opp_button.config(state=tk.NORMAL)
                self.power_off_button.config(state=tk.NORMAL)
                self.rc_dc_button.config(state=tk.NORMAL)
            
            self.run_final_button.config(state=tk.NORMAL)
        else:
            self.file_selection_frame.grid_remove()
            self.satc_button.config(state=tk.DISABLED)
            self.blp_button.config(state=tk.DISABLED)
            self.dlp_button.config(state=tk.DISABLED)
            self.billing_button.config(state=tk.DISABLED)
            self.pdrc_button.config(state=tk.DISABLED)
            self.sm_to_sm_button.config(state=tk.DISABLED)
            self.non_opp_button.config(state=tk.DISABLED)
            self.power_off_button.config(state=tk.DISABLED)
            self.rc_dc_button.config(state=tk.DISABLED)
            self.run_final_button.config(state=tk.DISABLED)
            self.log_message("Final summary process cancelled.")

    def check_vpn(self):
        """Check VPN connection by testing database server connectivity"""
        self.log_message("Checking VPN connection for SLA module...")
        
        host = get_config()['SLA']['ip']
        port = int(get_config()['SLA']['port'])
        
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            s.settimeout(5)
            s.connect((host, port))
            s.close()
            self.log_message("VPN connected successfully for SLA module!")
            self.vpn_connected = True
        except socket.timeout:
            self.log_message("VPN connection timeout - server not responding")
            self.vpn_connected = False
        except ConnectionRefusedError:
            self.log_message("VPN connection refused - server may be down")
            self.vpn_connected = False
        except socket.gaierror:
            self.log_message("VPN hostname resolution failed - check DNS or VPN")
            self.vpn_connected = False
        except Exception as e:
            self.log_message(f"VPN check error: {str(e)}")
            self.vpn_connected = False
    
    def wait_for_vpn(self, max_retries=1000, retry_interval=5):
        """Wait for VPN connection with retries"""
        self.log_message("Waiting for VPN connection...")
        
        host = get_config()['SLA']['ip']
        port = int(get_config()['SLA']['port'])
        
        retries = 0
        while retries < max_retries and not self.stop_processing:
            try:
                s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                s.settimeout(5)
                s.connect((host, port))
                s.close()
                self.log_message("VPN connected successfully!")
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
    
    def connect_db(self):
        if self.db_connection:
            try:
                conn, cur = self.db_connection
                cur.execute("SELECT 1")
                return cur, conn
            except:
                pass
    
        cfg = get_config()['SLA']
        try:
            conn = psycopg2.connect(
                host=cfg['ip'], port=cfg['port'], database=cfg['database'],
                user=cfg['uid'], password=cfg['pwd']
            )
            cur = conn.cursor()
            self.db_connection = (conn, cur)
            return cur, conn
        except (Exception, psycopg2.DatabaseError) as error:
            self.log_message(f'SLA DB Connection Failed. Error: {error}')
            return None, None

    def connect_rc_dc_db(self):
        if self.rc_dc_db_connection:
            try:
                conn, cur = self.rc_dc_db_connection
                cur.execute("SELECT 1")
                return cur, conn
            except:
                pass
    
        cfg = get_config()['RC_DC']
        try:
            conn = psycopg2.connect(
                host=cfg['ip'], port=cfg['port'], database=cfg['database'],
                user=cfg['uid'], password=cfg['pwd']
            )
            cur = conn.cursor()
            self.rc_dc_db_connection = (conn, cur)
            return cur, conn
        except (Exception, psycopg2.DatabaseError) as error:
            self.log_message(f'RC_DC DB Connection Failed. Error: {error}')
            return None, None

    def close_db_connection(self):
        """Close database connection"""
        if self.db_connection:
            try:
                conn, cur = self.db_connection
                cur.close()
                conn.close()
                self.db_connection = None
            except Exception as e:
                self.log_message(f"Error closing database connection: {str(e)}")
        
        if self.rc_dc_db_connection:
            try:
                conn, cur = self.rc_dc_db_connection
                cur.close()
                conn.close()
                self.rc_dc_db_connection = None
            except Exception as e:
                self.log_message(f"Error closing RC_DC database connection: {str(e)}")
    
    def execute_query_with_retry(self, query, query_name):
        """Execute query with retry logic"""
        retry_count = 0
        
        while retry_count < self.MAX_RETRIES and not self.stop_processing:
            try:
                cur, conn = self.connect_db()
                if cur is None:
                    if not self.wait_for_vpn(max_retries=10000, retry_interval=5):
                        break
                    cur, conn = self.connect_db()
                    if cur is None:
                        break
                
                cur.execute(query)
                
                results = []
                while True:
                    rows = cur.fetchmany(self.BATCH_SIZE)
                    if not rows:
                        break
                    
                    columns = [desc[0] for desc in cur.description]
                    for row in rows:
                        results.append(dict(zip(columns, row)))
                    
                    self.log_message(f"Fetched {len(results)} records")
                    
                    if self.stop_processing:
                        break
                
                return results
                
            except Exception as e:
                self.log_message(f"{query_name} failed (Attempt {retry_count + 1}/{self.MAX_RETRIES}): {str(e)}")
                retry_count += 1
                self.close_db_connection()
                
                if retry_count >= self.MAX_RETRIES:
                    self.log_message(f"Max retries reached for {query_name}")
                    break
                
                for i in range(int(self.RETRY_INTERVAL_MS/1000)):
                    if self.stop_processing:
                        break
                    time.sleep(1)
        
        return []
    
    def execute_rc_dc_query_with_retry(self, query, query_name):
        """Execute RC_DC query with retry logic using separate database connection"""
        retry_count = 0
        
        while retry_count < self.MAX_RETRIES and not self.stop_processing:
            try:
                cur, conn = self.connect_rc_dc_db()
                if cur is None:
                    if not self.wait_for_vpn(max_retries=10000, retry_interval=5):
                        break
                    cur, conn = self.connect_rc_dc_db()
                    if cur is None:
                        break
                
                cur.execute(query)
                
                results = []
                while True:
                    rows = cur.fetchmany(self.BATCH_SIZE)
                    if not rows:
                        break
                    
                    columns = [desc[0] for desc in cur.description]
                    for row in rows:
                        results.append(dict(zip(columns, row)))
                    
                    self.log_message(f"Fetched {len(results)} records")
                    
                    if self.stop_processing:
                        break
                
                return results
                
            except Exception as e:
                self.log_message(f"{query_name} failed (Attempt {retry_count + 1}/{self.MAX_RETRIES}): {str(e)}")
                retry_count += 1
                
                if self.rc_dc_db_connection:
                    try:
                        conn, cur = self.rc_dc_db_connection
                        cur.close()
                        conn.close()
                        self.rc_dc_db_connection = None
                    except:
                        pass
                
                if retry_count >= self.MAX_RETRIES:
                    self.log_message(f"Max retries reached for {query_name}")
                    break
                
                for i in range(int(self.RETRY_INTERVAL_MS/1000)):
                    if self.stop_processing:
                        break
                    time.sleep(1)
        
        return []
    
    def get_month_year_from_date(self, date_str):
        """Extract month and year in MMYY format from date string"""
        date_obj = datetime.datetime.strptime(date_str, "%Y-%m-%d")
        return date_obj.strftime("%m%y")
    
    def get_month_year_dash_format(self, date_str):
        """Extract month and year in MM-YYYY format from date string (e.g. 09-2026)"""
        date_obj = datetime.datetime.strptime(date_str, "%Y-%m-%d")
        return date_obj.strftime("%m-%Y")
    
    def get_nod_from_month_year(self, month_year_str):
        """Calculate NOD from MMYY format"""
        try:
            month = int(month_year_str[:2])
            year = int(month_year_str[2:])
            
            if month in [1, 3, 5, 7, 8, 10, 12]:
                return 31
            elif month in [4, 6, 9, 11]:
                return 30
            else:
                if (year % 4 == 0 and year % 100 != 0) or (year % 400 == 0):
                    return 29
                else:
                    return 28
        except:
            return 30
    
    def extract_month_year_from_filename(self, filename):
        """Extract MMYY from filename like SATC_RAW_0526_20260703"""
        try:
            match = re.search(r'_(\d{4})_', filename)
            if match:
                return match.group(1)
            
            match = re.search(r'_(\d{4})$', filename)
            if match:
                return match.group(1)
                
            return None
        except:
            return None
    
    def get_from_date_from_mmyy(self, mmyy_str):
        """Convert MMYY to from_date (first day of that month)"""
        try:
            month = int(mmyy_str[:2])
            year = int(mmyy_str[2:])
            
            if year < 100:
                year = 2000 + year
            
            return f"{year}-{month:02d}-01"
        except:
            return None
    
    def save_excel_by_hes(self, df, base_filename):
        """Save Excel file with separate sheets for each HES name"""
        if df is None or df.empty:
            self.log_message(f"No data to save for {base_filename}")
            return
        
        current_date = datetime.datetime.now().strftime("%Y%m%d")
        filename = os.path.join(self.save_directory, f"{base_filename}_{current_date}.xlsx")
        
        try:
            hes_column = None
            if 'hes_name' in df.columns:
                hes_column = 'hes_name'
            elif 'hes' in df.columns:
                hes_column = 'hes'
            
            if hes_column:
                hes_names = df[hes_column].unique()
                
                with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
                    for hes in hes_names:
                        if hes and str(hes) != '0' and str(hes) != 'nan':
                            hes_data = df[df[hes_column] == hes].copy()
                            sheet_name = str(hes)[:31]
                            sheet_name = sheet_name.replace('/', '_').replace('\\', '_').replace('*', '_').replace('?', '_').replace(':', '_')
                            
                            if 'location_code' in hes_data.columns:
                                hes_data['location_code'] = hes_data['location_code'].astype(str).str.zfill(4)
                            if 'consumer_id' in hes_data.columns:
                                hes_data['consumer_id'] = hes_data['consumer_id'].astype(str)
                            
                            hes_data.to_excel(writer, sheet_name=sheet_name, index=False)
                            self.log_message(f"'{sheet_name}' with {len(hes_data)} records")
                    
                self.log_message(f"Data saved to: {filename}")
            else:
                df_copy = df.copy()
                if 'location_code' in df_copy.columns:
                    df_copy['location_code'] = df_copy['location_code'].astype(str).str.zfill(4)
                if 'consumer_id' in df_copy.columns:
                    df_copy['consumer_id'] = df_copy['consumer_id'].astype(str)
                df_copy.to_excel(filename, index=False)
                self.log_message(f"Data saved to: {filename}")
                
        except Exception as e:
            self.log_message(f"Error saving Excel file: {str(e)}")
            csv_filename = filename.replace('.xlsx', '.csv')
            df_copy = df.copy()
            if 'location_code' in df_copy.columns:
                df_copy['location_code'] = df_copy['location_code'].astype(str).str.zfill(4)
            if 'consumer_id' in df_copy.columns:
                df_copy['consumer_id'] = df_copy['consumer_id'].astype(str)
            df_copy.to_csv(csv_filename, index=False)
            self.log_message(f"Saved as CSV instead: {csv_filename}")
    
    def save_dataframe(self, df, base_filename):
        """Save dataframe to Excel or CSV"""
        if df is None or df.empty:
            self.log_message(f"No data to save for {base_filename}")
            return False
        
        current_date = datetime.datetime.now().strftime("%Y%m%d")
        filename = os.path.join(self.save_directory, f"{base_filename}_{current_date}.xlsx")
        
        try:
            if len(df) > 1048576:
                csv_filename = filename.replace('.xlsx', '.csv')
                df.to_csv(csv_filename, index=False)
                self.log_message(f"Data too large for Excel ({len(df)} rows). Saved to CSV: {csv_filename}")
            else:
                df.to_excel(filename, index=False)
                self.log_message(f"Data saved to Excel: {filename}")
            return True
        except Exception as e:
            self.log_message(f"Error saving file: {str(e)}")
            return False
    
    def run_query_satc(self):
        """Run query to fetch SATC master data"""
        query = """
        SELECT location_code, consumer_id, meter_type, category, meterno, 
               meter_make, installation_date, hes, supported_ip, sat_pkg, satc 
        FROM unified_hes.uni_consumer_master
        WHERE satc = 'SATC'
        """
        
        results = self.execute_query_with_retry(query, "SATC Query")
        
        if results:
            df = pd.DataFrame(results)
            self.log_message(f"Total records fetched: {len(df)}")
            return df
        else:
            self.log_message("SATC Query completed but no data found")
            return pd.DataFrame()
    
    def run_query_1_blp(self, from_date, to_date):
        """Run first query for BLP data"""
        to_date_dt = datetime.datetime.strptime(to_date, "%Y-%m-%d") + datetime.timedelta(days=1)
        query_to_date = to_date_dt.strftime("%Y-%m-%d")
        
        query = f"""
        SELECT
            c.meterno AS meter_number,
            COALESCE(a.hes_name, c.hes) AS hes_name,
            COALESCE(a.H8, 0)  AS H8,
            COALESCE(a.H12, 0) AS H12,
            COALESCE(a.H24, 0) AS H24
        FROM unified_hes.uni_consumer_master c
        LEFT JOIN (
            SELECT
                meter_number,
                MAX(hes_name) AS hes_name,
                SUM(ls_tcount_first_timeslot)  AS H8,
                SUM(ls_tcount_second_timeslot) AS H12,
                SUM(ls_tcount_third_timeslot)  AS H24
            FROM unified_hes.uni_ca_agg_load_data
            WHERE report_date >= '{from_date}'
              AND report_date <  '{query_to_date}'
            GROUP BY meter_number
        ) a
        ON a.meter_number = c.meterno
        WHERE c.satc = 'SATC'
        """
        
        results = self.execute_query_with_retry(query, "BLP Query")
        
        if results:
            df = pd.DataFrame(results)
            self.log_message(f"Total records fetched: {len(df)}")
            return df
        else:
            self.log_message("BLP Query completed but no data found")
            return pd.DataFrame()
    
    def run_query_2_dlp(self, from_date, to_date):
        """Run second query for DLP data"""
        to_date_dt = datetime.datetime.strptime(to_date, "%Y-%m-%d") + datetime.timedelta(days=1)
        query_to_date = to_date_dt.strftime("%Y-%m-%d")
        
        query = f"""
        SELECT
            c.meterno AS meter_number,
            COALESCE(a.hes_name, c.hes) AS hes_name,
            COALESCE(a.H24, 0) AS H24
        FROM unified_hes.uni_consumer_master c
        LEFT JOIN (
            SELECT
                meter_number,
                MAX(hes_name) AS hes_name,
                SUM(ls_tcount_third_timeslot)  AS H24
            FROM unified_hes.uni_ca_agg_daily_data a
            where a.report_date >= '{from_date}'
              AND a.report_date < '{query_to_date}'
            GROUP BY meter_number
        ) a
        ON a.meter_number = c.meterno
        WHERE c.satc = 'SATC'
        """
        
        results = self.execute_query_with_retry(query, "DLP Query")
        
        if results:
            df = pd.DataFrame(results)
            self.log_message(f"Total records fetched: {len(df)}")
            return df
        else:
            self.log_message("DLP Query completed but no data found")
            return pd.DataFrame()
    
    def run_query_3_billing(self, from_date):
        """Run third query for Billing data with dynamic date conditions"""
        from_date_obj = datetime.datetime.strptime(from_date, "%Y-%m-%d")
        year = from_date_obj.year
        month = from_date_obj.month
        
        first_day = datetime.datetime(year, month, 1)
        date1_start_dt = first_day - datetime.timedelta(days=3)
        date1_start = date1_start_dt.strftime("%Y-%m-%d %H:%M:%S")
        
        date1_end = f"{year}-{month:02d}-04 00:00:00"
        date2_start = f"{year}-{month:02d}-04 00:00:00"
        date2_end = f"{year}-{month:02d}-08 00:00:00"
        
        query = f"""
        SELECT 
            COALESCE(b.meter_number, cm.meterno) AS meter_number,
            COALESCE(b.hes_name, cm.hes) AS hes_name,
            COALESCE(b.time_stamp::text, '0') AS time_stamp,
            COALESCE(b.bill_date::text, '0') AS bill_date,
            COALESCE(b.server_time::text, '0') AS server_time,
            COALESCE(b.sla_id, '0') AS sla_id,
            CASE 
                WHEN b.server_time >= '{date1_start}' AND b.server_time < '{date1_end}' THEN 1
                WHEN b.server_time >= '{date2_start}' AND b.server_time < '{date2_end}' THEN 0
                ELSE 0
            END AS "72hrs",
            CASE 
                WHEN b.server_time >= '{date1_start}' AND b.server_time < '{date1_end}' THEN 1
                WHEN b.server_time >= '{date2_start}' AND b.server_time < '{date2_end}' THEN 1
                ELSE 0
            END AS "168hrs"
        FROM unified_hes.uni_consumer_master cm
        LEFT JOIN (
            SELECT hes_name, time_stamp, meter_number, bill_date, server_time, sla_id
            FROM unified_hes.uni_bill_data
            where bill_date >= '2026-09-01' and bill_date < '2026-09-02'
            and((server_time >= '{date1_start}' AND server_time < '{date1_end}')
               OR (server_time >= '{date2_start}' AND server_time < '{date2_end}'))
        ) b ON cm.meterno = b.meter_number
        WHERE cm.satc = 'SATC'
        ORDER BY b.bill_date, cm.meterno
        """
        
        results = self.execute_query_with_retry(query, "Billing Query")
        
        if results:
            df = pd.DataFrame(results)
            self.log_message(f"Total records fetched: {len(df)}")
            return df
        else:
            self.log_message("Billing Query completed but no data found")
            return pd.DataFrame()
    
    def run_query_4_pdc(self):
        """Run fourth query for PDC data"""
        query = """
        SELECT consumer_id, meterno, meter_make, category, installation_date, 
               hes, decommissioning_reason, decommissioning_date, supported_ip, 
               sat_pkg, satc 
        FROM unified_hes.uni_consumer_master
        WHERE satc = 'SATC'
        AND decommissioning_reason = 'PDC'
        """
        
        results = self.execute_query_with_retry(query, "PDC Query")
        
        if results:
            df = pd.DataFrame(results)
            self.log_message(f"Total records fetched: {len(df)}")
            return df
        else:
            self.log_message("PDC Query completed but no data found")
            return pd.DataFrame()
    
    def run_query_5_sm_to_sm(self, from_date):
        """Run fifth query for SM to SM data with dynamic month_year from from_date"""
        month_year = self.get_month_year_dash_format(from_date)
        
        query = f"""
        SELECT * FROM unified_hes.uni_non_operational_meters 
        WHERE exception LIKE 'SM%'
        and month_year = '{month_year}'
        """
        
        results = self.execute_query_with_retry(query, "SM to SM Query")
        
        if results:
            df = pd.DataFrame(results)
            self.log_message(f"Total records fetched: {len(df)}")
            return df
        else:
            self.log_message("SM to SM Query completed but no data found")
            return pd.DataFrame()
    
    def run_query_6_non_opp(self, from_date):
        """Run sixth query for Non-Operational data with dynamic month_year from from_date"""
        month_year = self.get_month_year_dash_format(from_date)
        
        query = f"""
        SELECT * FROM unified_hes.uni_non_operational_meters 
        WHERE exception NOT LIKE 'SM%'
        and month_year = '{month_year}'
        """
        
        results = self.execute_query_with_retry(query, "Non-Operational Query")
        
        if results:
            df = pd.DataFrame(results)
            self.log_message(f"Total records fetched: {len(df)}")
            return df
        else:
            self.log_message("Non-Operational Query completed but no data found")
            return pd.DataFrame()
    
    def run_query_7_power_off(self, from_date, to_date):
        """Run seventh query for Power OFF data"""
        from_date_start = f"{from_date} 00:00:00"
        to_date_end = f"{to_date} 23:59:59"
        
        query = f"""
        SELECT DD.consumer_id, DD.meterno,
    DD.sat_pkg,
    CC.hes_name,
    SUM((CAST(CC.date_difference AS NUMERIC) - 1) * DD.supported_ip) AS power_off_exceptions_count
FROM
(
    SELECT
        hes_name,
        event_occ_date,
        month_year,
        meter_number,
        date_difference
    FROM
    (
        SELECT
            hes_name,
            month_year,
            meter_number,
            event_restore_date,
            event_occ_date,
            restoration_date,
            occurrence_date,
            date_difference
        FROM
        (
            SELECT
                hes_name,
                month_year,
                meter_number,
                event_restore_date,
                event_occ_date,
                duration,
                diff,
                split_part(CAST(diff AS varchar), 'days', 1) AS diff_days,
                restoration_date,
                occurrence_date,
                DATE(restoration_date) - DATE(occurrence_date) AS date_difference
            FROM
            (
                SELECT
                    hes_name,
                    TO_CHAR(event_occ_date,'yyyymm') AS month_year,
                    meter_number,
                    event_restore_date,
                    event_occ_date,
                    ROUND(EXTRACT(EPOCH FROM (event_restore_date - event_occ_date))/60) AS duration,
                    (event_restore_date - event_occ_date) AS diff,
                    TO_CHAR(event_occ_date,'yyyy-mm-dd') AS occurrence_date,
                    TO_CHAR(event_restore_date,'yyyy-mm-dd') AS restoration_date
                FROM unified_hes.uni_powerevents_analysis_2days_mv aa
                LEFT JOIN
                (
                    SELECT DISTINCT meterno
                    FROM unified_hes.uni_consumer_master
                    WHERE installation_date > '2000-01-01'
                      AND commissioning_date > '2000-01-01'
                      AND decommissioning_date IS NULL
                      AND meterno IN
                      (
                          SELECT DISTINCT meter_serial_number
                          FROM
                          (
                              SELECT DISTINCT
                                  meter_serial_number,
                                  CASE
                                      WHEN ticket_status='CLOSED'
                                           AND closed_time >= '{from_date_start}'
                                      THEN 'OPEN'
                                      ELSE ticket_status
                                  END AS status
                              FROM unified_hes.uni_non_operational_meters
                              WHERE responsibility='Utility'
                                AND non_operation_date <= '{to_date_end}'
                                AND meter_serial_number IN
                                (
                                    SELECT DISTINCT meterno
                                    FROM unified_hes.uni_consumer_master
                                    WHERE satc='SATC'
                                )
                          ) x
                          WHERE status='OPEN'
                      )
                ) bb
                    ON aa.meter_number = bb.meterno
                WHERE bb.meterno IS NULL
                  AND event_occ_date >= '{from_date_start}'
                  AND event_occ_date <= '{to_date_end}'
                  AND event_restore_date >= '{from_date_start}'
                  AND event_restore_date <= '{to_date_end}'
                  AND event_name = 'Power Failure'
                  AND meter_number IN
                  (
                      SELECT meterno
                      FROM unified_hes.uni_consumer_master
                      WHERE installation_date > '2000-01-01'
                        AND commissioning_date > '2000-01-01'
                        AND decommissioning_date IS NULL
                        AND meterno IN
                        (
                            SELECT DISTINCT meterno
                            FROM unified_hes.uni_consumer_master
                            WHERE satc='SATC'
                        )
                  )
            ) AA
        ) AAA
    ) BB
) CC
INNER JOIN
(
    SELECT
        consumer_id,
        meterno,
        sat_pkg,
        supported_ip
    FROM unified_hes.uni_consumer_master
) DD
ON CC.meter_number = DD.meterno
WHERE ((CAST(CC.date_difference AS NUMERIC) - 1) * DD.supported_ip) > 0
GROUP BY
    DD.consumer_id,
    DD.meterno,
    DD.sat_pkg,
    CC.hes_name
ORDER BY
    DD.meterno;
        """
        
        results = self.execute_query_with_retry(query, "Power OFF Query")
        
        if results:
            df = pd.DataFrame(results)
            self.log_message(f"Total records fetched: {len(df)}")
            return df
        else:
            self.log_message("Power OFF Query completed but no data found")
            return pd.DataFrame()
    
    def run_query_9_power_off_prior_month(self, from_date, to_date):
        """
        Run the prior-month power OFF query.
        This handles events that started before from_date and restored within
        [from_date, to_date + 1). The count is the overlap of the outage with
        the current month window.
        """
        from_date_start = f"{from_date} 00:00:00"
        to_date_end = f"{to_date} 23:59:59"
        
        # to_date + 1 day (exclusive upper bound for event_restore_date)
        to_date_dt = datetime.datetime.strptime(to_date, "%Y-%m-%d") + datetime.timedelta(days=1)
        query_to_date = to_date_dt.strftime("%Y-%m-%d")
        
        query = f"""
        SELECT
            uc.consumer_id,
            pf.meter_number,
            uc.sat_pkg,
            uc.hes,
            EXTRACT(
                DAY FROM (
                    (
                        pf.event_restore_date - TIMESTAMP '{from_date_start}'
                    ) * uc.supported_ip
                )
            )::INTEGER AS power_off_exceptions_count
        FROM unified_hes.uni_powerevents_analysis_2days_mv pf
        INNER JOIN unified_hes.uni_consumer_master uc
            ON pf.meter_number = uc.meterno
        WHERE pf.event_name = 'Power Failure'
          AND pf.meter_number in (
            SELECT DD.meterno
            FROM
            (
                SELECT
                    hes_name,
                    event_occ_date,
                    month_year,
                    meter_number,
                    date_difference
                FROM
                (
                    SELECT
                        hes_name,
                        month_year,
                        meter_number,
                        event_restore_date,
                        event_occ_date,
                        restoration_date,
                        occurrence_date,
                        date_difference
                    FROM
                    (
                        SELECT
                            hes_name,
                            month_year,
                            meter_number,
                            event_restore_date,
                            event_occ_date,
                            duration,
                            diff,
                            split_part(CAST(diff AS varchar), 'days', 1) AS diff_days,
                            restoration_date,
                            occurrence_date,
                            DATE(restoration_date) - DATE(occurrence_date) AS date_difference
                        FROM
                        (
                            SELECT
                                hes_name,
                                TO_CHAR(event_occ_date,'yyyymm') AS month_year,
                                meter_number,
                                event_restore_date,
                                event_occ_date,
                                ROUND(EXTRACT(EPOCH FROM (event_restore_date - event_occ_date))/60) AS duration,
                                (event_restore_date - event_occ_date) AS diff,
                                TO_CHAR(event_occ_date,'yyyy-mm-dd') AS occurrence_date,
                                TO_CHAR(event_restore_date,'yyyy-mm-dd') AS restoration_date
                            FROM unified_hes.uni_powerevents_analysis_2days_mv aa
                            LEFT JOIN
                            (
                                SELECT DISTINCT meterno
                                FROM unified_hes.uni_consumer_master
                                WHERE installation_date > '2000-01-01'
                                  AND commissioning_date > '2000-01-01'
                                  AND decommissioning_date IS NULL
                                  AND meterno IN
                                  (
                                      SELECT DISTINCT meter_serial_number
                                      FROM
                                      (
                                          SELECT DISTINCT
                                              meter_serial_number,
                                              CASE
                                                  WHEN ticket_status='CLOSED'
                                                       AND closed_time >= '{from_date_start}'
                                                  THEN 'OPEN'
                                                  ELSE ticket_status
                                              END AS status
                                          FROM unified_hes.uni_non_operational_meters
                                          WHERE responsibility='Utility'
                                            AND non_operation_date <= '{to_date_end}'
                                            AND meter_serial_number IN
                                            (
                                                SELECT DISTINCT meterno
                                                FROM unified_hes.uni_consumer_master
                                                WHERE satc='SATC'
                                            )
                                      ) x
                                      WHERE status='OPEN'
                                  )
                            ) bb
                                ON aa.meter_number = bb.meterno
                            WHERE bb.meterno IS NULL
                              AND event_occ_date < '{from_date_start}'
                              AND event_restore_date >= '{from_date_start}'
                              AND event_restore_date < '{query_to_date}'
                              AND event_name = 'Power Failure'
                              AND meter_number IN
                              (
                                  SELECT meterno
                                  FROM unified_hes.uni_consumer_master
                                  WHERE installation_date > '2000-01-01'
                                    AND commissioning_date > '2000-01-01'
                                    AND decommissioning_date IS NULL
                                    AND meterno IN
                                    (
                                        SELECT DISTINCT meterno
                                        FROM unified_hes.uni_consumer_master
                                        WHERE satc='SATC'
                                    )
                              )
                        ) AA
                    ) AAA
                ) BB
            ) CC
            INNER JOIN
            (
                SELECT
                    consumer_id,
                    meterno,
                    sat_pkg,
                    supported_ip
                FROM unified_hes.uni_consumer_master
            ) DD
            ON CC.meter_number = DD.meterno
            WHERE ((CAST(CC.date_difference AS NUMERIC) - 1) * DD.supported_ip) > 0
            GROUP BY
                DD.consumer_id,
                DD.meterno,
                DD.sat_pkg,
                CC.hes_name
            ORDER BY
                DD.meterno
          )
          AND pf.event_occ_date < TIMESTAMP '{from_date_start}'
          AND pf.event_restore_date >= TIMESTAMP '{from_date_start}'
          AND pf.event_restore_date < TIMESTAMP '{query_to_date}';
        """
        
        results = self.execute_query_with_retry(query, "Power OFF Prior Month Query")
        
        if results:
            df = pd.DataFrame(results)
            self.log_message(f"Total prior-month power OFF records fetched: {len(df)}")
            return df
        else:
            self.log_message("Power OFF Prior Month Query completed but no data found")
            return pd.DataFrame()
    
    def merge_power_off_results(self, df_current, df_prior):
        """
        Merge current-month power OFF data with prior-month power OFF data.
        - Match on meter_number
        - Sum power_off_exceptions_count for the same meter
        - Keep one row per meter
        - Include meters that only appear in the prior-month result
        """
        if df_current is None or df_current.empty:
            if df_prior is None or df_prior.empty:
                return pd.DataFrame()
            # Normalize prior-only result
            df_prior = df_prior.copy()
            df_prior.columns = ['consumer_id', 'meter_number', 'sat_pkg', 'hes',
                                'power_off_exceptions_count']
            return df_prior
        
        df_current = df_current.copy()
        df_current.columns = ['consumer_id', 'meter_number', 'sat_pkg', 'hes',
                              'power_off_exceptions_count']
        
        if df_prior is None or df_prior.empty:
            return df_current
        
        df_prior = df_prior.copy()
        df_prior.columns = ['consumer_id', 'meter_number', 'sat_pkg', 'hes',
                            'power_off_exceptions_count']
        
        # Ensure numeric counts
        df_current['power_off_exceptions_count'] = pd.to_numeric(
            df_current['power_off_exceptions_count'], errors='coerce').fillna(0)
        df_prior['power_off_exceptions_count'] = pd.to_numeric(
            df_prior['power_off_exceptions_count'], errors='coerce').fillna(0)
        
        # Concatenate and aggregate per meter_number
        combined = pd.concat([df_current, df_prior], ignore_index=True)
        
        # For non-numeric descriptive columns, keep the first non-null value
        combined['meter_number'] = combined['meter_number'].astype(str)
        
        agg = combined.groupby('meter_number', as_index=False).agg({
            'consumer_id': 'first',
            'sat_pkg': 'first',
            'hes': 'first',
            'power_off_exceptions_count': 'sum',
        })
        
        # Restore original column order
        agg = agg[['consumer_id', 'meter_number', 'sat_pkg', 'hes',
                   'power_off_exceptions_count']]
        
        return agg
    
    def run_query_8_rc_dc(self, from_date, to_date):
        """Run eighth query for RC_DC data using separate database connection - UPDATED with new query"""
        to_date_dt = datetime.datetime.strptime(to_date, "%Y-%m-%d") + datetime.timedelta(days=1)
        query_to_date = to_date_dt.strftime("%Y-%m-%d")
        
        query = f"""
        WITH ranked_data AS (
            SELECT rd.sdocode,mm.village AS sat_no,rd.mtrno,rd.hes_type,rd.req_type,p.request_time,rd.success_time AS hes_response_time,rd.job_id,p.rms_req_id,
        (rd.success_time::timestamp - p.request_time::timestamp) AS duration_minutes,rd.try_count,
                CASE
                    WHEN EXTRACT(EPOCH FROM (
                        rd.success_time::timestamp - p.request_time::timestamp
                    )) / 60 <= 15 THEN 1
                    ELSE 0
                END AS "15_min",
                CASE
                    WHEN EXTRACT(EPOCH FROM (
                        rd.success_time::timestamp - p.request_time::timestamp
                    )) / 60 <= 360 THEN 1
                    ELSE 0
                END AS "360_min",
                ROW_NUMBER() OVER (
                    PARTITION BY p.rms_req_id
                    ORDER BY p.request_time ASC, rd.first_try_time ASC
                ) AS rn
          FROM ami_master.rc_dc_log rd
            INNER JOIN ami_master.master_main mm
                ON rd.mtrno = mm.mtrno
            INNER JOIN ami_master.post_rcdc_tracker p
                ON rd.req_id = p.rms_req_id
                AND rd.mtrno = p.mtrno
            WHERE mm.mla = 'SAT'
              AND mm.comm IN ('0','1')
              AND rd.first_try_time >= '{from_date}'
              AND rd.first_try_time < '{query_to_date}'
              AND p.request_time >= '{from_date}'
              AND p.request_time < '{query_to_date}'
              AND p.success_flag = 't'
              AND rd.status = 'S'
              AND rd.success_time IS NOT NULL
              AND rd.first_try_time IS NOT NULL
              AND p.request_time IS NOT NULL
        )
        SELECT sdocode,sat_no,mtrno,hes_type,req_type,request_time,hes_response_time,job_id,rms_req_id,duration_minutes,try_count,"15_min","360_min"
        FROM ranked_data
        WHERE rn = 1
        ORDER BY mtrno, req_type, request_time;
        """
        
        results = self.execute_rc_dc_query_with_retry(query, "RC_DC Query")
        
        if results:
            df = pd.DataFrame(results)
            self.log_message(f"Total records fetched: {len(df)}")
            return df
        else:
            self.log_message("RC_DC Query completed but no data found")
            return pd.DataFrame()
    
    def run_power_events_query_for_rc_dc(self, meter_numbers, from_date, to_date):
        """Run power events query for RC_DC meters to determine exclusion flags"""
        if not meter_numbers:
            self.log_message("No meter numbers provided for power events query")
            return pd.DataFrame()
        
        # Format meter numbers for IN clause
        meter_list = "','".join(meter_numbers)
        
        # from_date - 1 month for event_occ_date
        from_date_obj = datetime.datetime.strptime(from_date, "%Y-%m-%d")
        if from_date_obj.month == 1:
            event_occ_date = from_date_obj.replace(year=from_date_obj.year - 1, month=12)
        else:
            event_occ_date = from_date_obj.replace(month=from_date_obj.month - 1)
        event_occ_date_str = event_occ_date.strftime("%Y-%m-%d")
        
        # to_date + 1 day for event_restore_date <
        to_date_dt = datetime.datetime.strptime(to_date, "%Y-%m-%d") + datetime.timedelta(days=1)
        query_to_date = to_date_dt.strftime("%Y-%m-%d")
        
        query = f"""
        select meter_number,event_occ_date,event_restore_date from unified_hes.uni_powerevents_analysis 
        where meter_number in ('{meter_list}')
        and event_occ_date >= '{event_occ_date_str}'
        and event_restore_date < '{query_to_date}'
        """
        
        results = self.execute_query_with_retry(query, "Power Events for RC_DC Query")
        
        if results:
            df = pd.DataFrame(results)
            self.log_message(f"Total power events fetched for RC_DC meters: {len(df)}")
            return df
        else:
            self.log_message("Power Events for RC_DC Query completed but no data found")
            return pd.DataFrame()
    
    def run_process(self):
        """Main process to run all SLA queries"""
        if self.fetch_var.get() == "YES":
            if not self.vpn_connected:
                self.log_message("VPN not connected. Please check VPN first.")
                messagebox.showwarning("VPN Not Connected", "Please check VPN connection before fetching data.")
                self.cleanup_after_process()
                return
            
            from_date = self.from_date_entry.get()
            to_date = self.to_date_entry.get()
            
            self.log_message("=" * 60)
            self.log_message("Starting SLA Data Extraction Process...")
            self.log_message(f"From Date: {from_date}")
            self.log_message(f"To Date: {to_date}")
            self.log_message(f"Exclusion: {'ENABLED' if self.exclusion_enabled else 'DISABLED'}")
            self.log_message("=" * 60)
            
            month_year = self.get_month_year_from_date(from_date)
            current_date = datetime.datetime.now().strftime("%Y%m%d")
            
            # Query 0: SATC Master Data
            if not self.stop_processing:
                self.log_message("\n[0/8] Fetching SATC Master data...")
                df_satc = self.run_query_satc()
                if not df_satc.empty:
                    self.save_excel_by_hes(df_satc, f"SATC_RAW_{month_year}")
            
            # Query 1: BLP Data
            if not self.stop_processing:
                self.log_message("\n[1/8] Fetching BLP data...")
                df_blp = self.run_query_1_blp(from_date, to_date)
                if not df_blp.empty:
                    self.save_excel_by_hes(df_blp, f"BLP_RAW_SLA_{month_year}")
            
            # Query 2: DLP Data
            if not self.stop_processing:
                self.log_message("\n[2/8] Fetching DLP data...")
                df_dlp = self.run_query_2_dlp(from_date, to_date)
                if not df_dlp.empty:
                    self.save_excel_by_hes(df_dlp, f"DLP_RAW_SLA_{month_year}")
            
            # Query 3: Billing Data
            if not self.stop_processing:
                self.log_message("\n[3/8] Fetching Billing data...")
                df_billing = self.run_query_3_billing(from_date)
                if not df_billing.empty:
                    self.save_excel_by_hes(df_billing, f"BILLING_RAW_SLA_{month_year}")
            
            # Power OFF query - ALWAYS run regardless of exclusion setting
            if not self.stop_processing:
                self.log_message("\n[4/8] Fetching Power OFF data...")
                df_power_off = self.run_query_7_power_off(from_date, to_date)
                
                # NEW: Immediate second power OFF query (prior-month carry-over)
                if not self.stop_processing:
                    self.log_message("  Fetching prior-month Power OFF data (carry-over events)...")
                    df_power_off_prior = self.run_query_9_power_off_prior_month(from_date, to_date)
                    
                    # Merge both power OFF datasets (sum counts per meter_number)
                    self.log_message("  Merging current + prior-month Power OFF counts...")
                    df_power_off = self.merge_power_off_results(df_power_off, df_power_off_prior)
                    self.log_message(f"  Merged Power OFF records: {len(df_power_off)}")
                
                if not df_power_off.empty:
                    self.save_dataframe(df_power_off, f"Power_OFF_RAW_{month_year}")
            
            # RC_DC query - ALWAYS run regardless of exclusion setting (NEW)
            if not self.stop_processing:
                self.log_message("\n[5/8] Fetching RC_DC data...")
                df_rc_dc = self.run_query_8_rc_dc(from_date, to_date)
                if not df_rc_dc.empty:
                    # Map exclusions HES-type-wise (batched power events fetch)
                    df_rc_dc = self.map_rc_dc_exclusions_by_hes_type(df_rc_dc, from_date, to_date)
                    self.save_dataframe(df_rc_dc, f"RC_DC_RAW_{month_year}")
            
            # SM to SM query - ALWAYS run regardless of exclusion setting (NEW)
            if not self.stop_processing:
                self.log_message("\n[6/8] Fetching SM to SM data...")
                df_sm = self.run_query_5_sm_to_sm(from_date)
                if not df_sm.empty:
                    self.save_dataframe(df_sm, f"SM_to_SM_RAW_{month_year}")
            
            # Only run other exclusion-related queries if Exclusion is YES
            if self.exclusion_enabled:
                # Query 4: PDC Data
                if not self.stop_processing:
                    self.log_message("\n[7/8] Fetching PDC data...")
                    df_pdc = self.run_query_4_pdc()
                    if not df_pdc.empty:
                        self.save_dataframe(df_pdc, f"PDC_RAW_{month_year}")
                
                # Query 6: Non-Operational Data
                if not self.stop_processing:
                    self.log_message("\n[8/8] Fetching Non-Operational data...")
                    df_non_opp = self.run_query_6_non_opp(from_date)
                    if not df_non_opp.empty:
                        self.save_dataframe(df_non_opp, f"NON_OPP_RAW_{month_year}")
            else:
                self.log_message("\n⚠ Exclusion is DISABLED. Skipping PDC and Non-Operational queries.")
                self.log_message("✓ Power OFF, RC_DC and SM to SM queries have been executed.")
            
            # Process completion
            if not self.stop_processing:
                self.log_message("\nDo you want final Cumulative Summary..?")
                self.summary_frame.grid(row=6, column=0, columnspan=3, pady=10, sticky=tk.EW)
            else:
                self.log_message("\n✗ Process stopped by user.")
        
        self.cleanup_after_process()
        self.close_db_connection()
    
    def _map_exclusions_for_subset(self, df_subset, df_power_events):
        """
        Map exclusion flags for a subset of RC_DC data based on power events.
        Returns the subset with Excl_flag, Occ_datetime, Res_datetime populated.
        """
        df_subset = df_subset.copy()
        df_subset['request_time_dt'] = pd.to_datetime(df_subset['request_time'], errors='coerce')
        
        df_subset['Excl_flag'] = '-'
        df_subset['Occ_datetime'] = ''
        df_subset['Res_datetime'] = ''
        
        if df_power_events is None or df_power_events.empty:
            df_subset.drop(columns=['request_time_dt'], inplace=True, errors='ignore')
            return df_subset
        
        df_power_events = df_power_events.copy()
        df_power_events['event_occ_date_dt'] = pd.to_datetime(df_power_events['event_occ_date'], errors='coerce')
        df_power_events['event_restore_date_dt'] = pd.to_datetime(df_power_events['event_restore_date'], errors='coerce')
        
        meter_events = {}
        for _, row in df_power_events.iterrows():
            mtr = str(row['meter_number'])
            occ = row['event_occ_date_dt']
            res = row['event_restore_date_dt']
            if pd.notna(occ) and pd.notna(res):
                meter_events.setdefault(mtr, []).append((occ, res))
        
        for idx, row in df_subset.iterrows():
            mtr = str(row['mtrno'])
            req_time = row['request_time_dt']
            if mtr in meter_events and pd.notna(req_time):
                for occ, res in meter_events[mtr]:
                    if occ <= req_time <= res:
                        df_subset.at[idx, 'Excl_flag'] = 'Power_off'
                        df_subset.at[idx, 'Occ_datetime'] = occ.strftime('%Y-%m-%d %H:%M:%S')
                        df_subset.at[idx, 'Res_datetime'] = res.strftime('%Y-%m-%d %H:%M:%S')
                        break
        
        power_off_mask = df_subset['Excl_flag'] == 'Power_off'
        df_subset.loc[power_off_mask, '15_min'] = 1
        df_subset.loc[power_off_mask, '360_min'] = 1
        
        df_subset.drop(columns=['request_time_dt'], inplace=True, errors='ignore')
        return df_subset
    
    def map_rc_dc_exclusions_by_hes_type(self, df_rc_dc, from_date, to_date):
        """
        Map exclusion flags for RC_DC data HES-type-wise (batched) to avoid
        loading the entire huge uni_powerevents_analysis table at once.
        """
        self.log_message("  Mapping exclusion flags for RC_DC data (HES-type-wise batching)...")
        
        if 'hes_type' not in df_rc_dc.columns:
            self.log_message("  WARNING: 'hes_type' column not found. Falling back to single batch.")
            unique_meters = df_rc_dc['mtrno'].dropna().astype(str).unique().tolist()
            df_power_events = self.run_power_events_query_for_rc_dc(unique_meters, from_date, to_date)
            return self._map_exclusions_for_subset(df_rc_dc, df_power_events)
        
        hes_types = df_rc_dc['hes_type'].dropna().astype(str).unique().tolist()
        self.log_message(f"  Found HES types: {hes_types}")
        
        mapped_parts = []
        
        for hes_type in hes_types:
            if self.stop_processing:
                break
            
            mask = df_rc_dc['hes_type'].astype(str) == hes_type
            df_subset = df_rc_dc[mask].copy()
            
            if df_subset.empty:
                continue
            
            unique_meters = df_subset['mtrno'].dropna().astype(str).unique().tolist()
            self.log_message(f"    [{hes_type}] Processing {len(df_subset)} RC_DC rows, {len(unique_meters)} unique meters...")
            
            df_power_events = self.run_power_events_query_for_rc_dc(unique_meters, from_date, to_date)
            mapped_subset = self._map_exclusions_for_subset(df_subset, df_power_events)
            mapped_parts.append(mapped_subset)
            
            po_count = (mapped_subset['Excl_flag'] == 'Power_off').sum()
            self.log_message(f"    [{hes_type}] Power_off exclusions mapped: {po_count}")
        
        if not mapped_parts:
            self.log_message("  No RC_DC data to map.")
            return df_rc_dc
        
        df_result = pd.concat(mapped_parts, ignore_index=True)
        
        # Reorder columns: duration_minutes, Excl_flag, Occ_datetime, Res_datetime, try_count
        cols = df_result.columns.tolist()
        try:
            insert_cols = ['Excl_flag', 'Occ_datetime', 'Res_datetime']
            for c in insert_cols:
                if c in cols:
                    cols.remove(c)
            dur_idx = cols.index('duration_minutes')
            for i, c in enumerate(insert_cols):
                cols.insert(dur_idx + 1 + i, c)
            df_result = df_result[cols]
        except ValueError:
            pass
        
        total_po = (df_result['Excl_flag'] == 'Power_off').sum()
        self.log_message(f"  Total RC_DC records with Power_off exclusion: {total_po}")
        
        return df_result
    
    def map_rc_dc_exclusions(self, df_rc_dc, df_power_events):
        """
        Backwards-compatible wrapper: map exclusions for full dataframe using
        a single power events dataframe.
        """
        self.log_message("  Mapping exclusion flags for RC_DC data...")
        return self._map_exclusions_for_subset(df_rc_dc, df_power_events)
    
    def cleanup_after_process(self):
        """Clean up after process completes or stops"""
        self.processing = False
        self.run_button.grid(row=5, column=0, columnspan=3, pady=20, sticky=tk.EW)
        self.stop_button.grid_remove()
    
    def select_satc_file(self):
        """Select SATC RAW file for final summary"""
        file_path = filedialog.askopenfilename(
            title="Select SATC RAW File",
            filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
        )
        if file_path:
            self.satc_file_path = file_path
            self.month_year_from_filename = self.extract_month_year_from_filename(os.path.basename(file_path))
            self.log_message(f"SATC RAW file selected: {file_path}")
            if self.month_year_from_filename:
                self.log_message(f"Month-Year : {self.month_year_from_filename}")
    
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
    
    def select_pdrc_file(self):
        """Select PDRC RAW file for final summary"""
        file_path = filedialog.askopenfilename(
            title="Select PDRC RAW File",
            filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
        )
        if file_path:
            self.pdrc_file_path = file_path
            self.log_message(f"PDRC RAW file selected: {file_path}")
    
    def select_sm_to_sm_file(self):
        """Select SM to SM RAW file for final summary"""
        file_path = filedialog.askopenfilename(
            title="Select SM to SM RAW File",
            filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
        )
        if file_path:
            self.sm_to_sm_file_path = file_path
            self.log_message(f"SM to SM RAW file selected: {file_path}")
    
    def select_non_opp_file(self):
        """Select Non Opp RAW file for final summary"""
        file_path = filedialog.askopenfilename(
            title="Select Non Opp RAW File",
            filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
        )
        if file_path:
            self.non_opp_file_path = file_path
            self.log_message(f"Non Opp RAW file selected: {file_path}")
    
    def select_power_off_file(self):
        """Select Power OFF RAW file for final summary"""
        file_path = filedialog.askopenfilename(
            title="Select Power OFF RAW File",
            filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
        )
        if file_path:
            self.power_off_file_path = file_path
            self.log_message(f"Power OFF RAW file selected: {file_path}")
    
    def select_rc_dc_file(self):
        """Select RC_DC RAW file for final summary"""
        file_path = filedialog.askopenfilename(
            title="Select RC_DC RAW File",
            filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
        )
        if file_path:
            self.rc_dc_file_path = file_path
            self.log_message(f"RC_DC RAW file selected: {file_path}")
    
    def calculate_nod(self, from_date):
        """Calculate Number of Days (NOD) based on from_date month"""
        date_obj = datetime.datetime.strptime(from_date, "%Y-%m-%d")
        month = date_obj.month
        year = date_obj.year
        
        if month in [1, 3, 5, 7, 8, 10, 12]:
            return 31
        elif month in [4, 6, 9, 11]:
            return 30
        else:
            if (year % 4 == 0 and year % 100 != 0) or (year % 400 == 0):
                return 29
            else:
                return 28
    
    def _first_record_map(self, df, key_col):
        """{str(key): first matching row as dict}."""
        out = {}
        if df is None or df.empty or key_col not in df.columns:
            return out
        keys = _to_str(df[key_col]).tolist()
        for k, rec in zip(keys, df.to_dict('records')):
            if k not in out:
                out[k] = rec
        return out
    
    def _build_exclusion_lookup(self, meter_list, pdc_df, sm_to_sm_df, non_opp_df, power_off_df, from_date):
        """Classifies every unique meter ONCE with the exact same rules and priority."""
        power_off_first = self._first_record_map(power_off_df, 'meterno')
        result = {}
        
        if not self.exclusion_enabled:
            for m in meter_list:
                rec = power_off_first.get(m)
                result[m] = ('Power_OFF', rec) if rec is not None else ('-', None)
            return result
        
        pdc_first = self._first_record_map(pdc_df, 'meterno')
        sm_first = self._first_record_map(sm_to_sm_df, 'meter_serial_number')
        non_opp_first = self._first_record_map(non_opp_df, 'meter_serial_number')
        
        from_date_obj = datetime.datetime.strptime(from_date, "%Y-%m-%d")
        if from_date_obj.month == 12:
            month_end = datetime.datetime(from_date_obj.year + 1, 1, 1) - datetime.timedelta(days=1)
        else:
            month_end = datetime.datetime(from_date_obj.year, from_date_obj.month + 1, 1) - datetime.timedelta(days=1)
        month_end_with_time = month_end.replace(hour=23, minute=59, second=59)
        month_start_with_time = from_date_obj.replace(hour=0, minute=0, second=0)
        
        for m in meter_list:
            rec = pdc_first.get(m)
            if rec is not None:
                result[m] = ('PD', rec)
                continue
            
            rec = sm_first.get(m)
            if rec is not None:
                ticket_status = rec.get('ticket_status', '')
                if ticket_status == 'CLOSED':
                    closed_time = rec.get('closed_time', '')
                    if closed_time and pd.notna(closed_time):
                        try:
                            closed_time_dt = pd.to_datetime(closed_time)
                            if month_start_with_time <= closed_time_dt <= month_end_with_time:
                                result[m] = ('SM_SM', rec)
                                continue
                        except Exception:
                            pass
                elif ticket_status == 'OPEN':
                    result[m] = ('SM_SM', rec)
                    continue
            
            rec = non_opp_first.get(m)
            if rec is not None:
                ticket_status = rec.get('ticket_status', '')
                if ticket_status == 'CLOSED':
                    closed_time = rec.get('closed_time', '')
                    if closed_time and pd.notna(closed_time):
                        try:
                            closed_time_dt = pd.to_datetime(closed_time)
                            if month_start_with_time <= closed_time_dt <= month_end_with_time:
                                result[m] = ('Non_Opp', rec)
                                continue
                        except Exception:
                            pass
                elif ticket_status == 'OPEN':
                    result[m] = ('Non_Opp', rec)
                    continue
            
            rec = power_off_first.get(m)
            if rec is not None:
                result[m] = ('Power_OFF', rec)
            else:
                result[m] = ('-', None)
        
        return result
    
    def _power_off_count_lookup(self, power_off_data):
        """{meter: power_off_count}"""
        out = {}
        if power_off_data is None or power_off_data.empty or \
           'meterno' not in power_off_data.columns or 'power_off_exceptions_count' not in power_off_data.columns:
            return out
        keys = _to_str(power_off_data['meterno']).tolist()
        counts = power_off_data['power_off_exceptions_count'].astype(object).tolist()
        for k, c in zip(keys, counts):
            if k:
                if c is not None and pd.notna(c):
                    out[k] = float(c) if isinstance(c, (int, float)) else 0
        return out
    
    def _last_value_map(self, df, key_col, value_col):
        """{meter: value from the LAST row of that meter}"""
        out = {}
        if df is None or df.empty or key_col not in df.columns or value_col not in df.columns:
            return out
        keys = _to_str(df[key_col]).tolist()
        values = df[value_col].astype(object).tolist()
        for k, v in zip(keys, values):
            if k:
                out[k] = v
        return out
    
    def get_exception_date(self, exclusion_type, match_record):
        """Get exception date based on exclusion type"""
        if exclusion_type == 'PD':
            return match_record.get('decommissioning_date', '') if match_record is not None else ''
        elif exclusion_type in ['SM_SM', 'Non_Opp']:
            if match_record is not None:
                ticket_status = match_record.get('ticket_status', '')
                if ticket_status == 'CLOSED':
                    return match_record.get('closed_time', '')
                elif ticket_status == 'OPEN':
                    return match_record.get('non_operation_date', '')
        elif exclusion_type == 'Power_OFF':
            return '-'
        return ''
    
    def get_exception_text(self, exclusion_type, match_record):
        """Get exception text based on exclusion type"""
        if exclusion_type == 'PD':
            return 'PD'
        elif exclusion_type == 'SM_SM':
            if match_record is not None:
                return match_record.get('remarks', 'SM_SM')
            return 'SM_SM'
        elif exclusion_type == 'Non_Opp':
            if match_record is not None:
                return match_record.get('remarks', 'Non_Opp')
            return 'Non_Opp'
        elif exclusion_type == 'Power_OFF':
            return '-'
        else:
            return '-'
    
    def get_ami_loc_data(self, bu_code):
        """Get zone, circle, division, subdivision from ami_loc based on BU code"""
        if self.ami_loc_df is None:
            return '', '', '', ''
        
        bu_str = str(bu_code)
        if bu_str.isdigit():
            bu_str = bu_str.zfill(4)
        
        match = self.ami_loc_df[self.ami_loc_df['sitecode'].astype(str).str.zfill(4) == bu_str]
        
        if not match.empty:
            row = match.iloc[0]
            zone = row.get('zone', '')
            circle = row.get('circle', '')
            division = row.get('division', '')
            subdivision = row.get('subdivision', '')
            return str(zone), str(circle), str(division), str(subdivision)
        
        return '', '', '', ''
    
    def create_hes_wise_raw_data(self, hes_name, satc_data, blp_data, dlp_data, billing_data, 
                              pdc_data, sm_to_sm_data, non_opp_data, power_off_data, from_date, nod):
        """Create HES-wise raw data sheet."""
        if 'hes' in satc_data.columns:
            hes_satc = satc_data[satc_data['hes'] == hes_name]
        else:
            hes_satc = satc_data
        
        if hes_satc.empty:
            self.log_message(f"No data found for HES: {hes_name}")
            return pd.DataFrame()
        
        hes_satc = hes_satc.reset_index(drop=True)
        cols = hes_satc.columns
        
        meter_s = _to_str(hes_satc['meterno']) if 'meterno' in cols else pd.Series([''] * len(hes_satc), dtype=object)
        keep = (meter_s != '').to_numpy()
        hes_satc = hes_satc.loc[keep].reset_index(drop=True)
        meter_s = meter_s[keep].reset_index(drop=True)
        if hes_satc.empty:
            return pd.DataFrame()
        n = len(hes_satc)
        
        def sat_col(name, default=''):
            if name in hes_satc.columns:
                return hes_satc[name]
            return pd.Series([default] * n, dtype=object)
        
        codes, meter_uniques = pd.factorize(meter_s)
        meter_list = meter_uniques.tolist()
        excl_info = self._build_exclusion_lookup(meter_list, pdc_data, sm_to_sm_data,
                                                 non_opp_data, power_off_data, from_date)
        
        po_map = self._power_off_count_lookup(power_off_data)
        blp_h8_map = self._last_value_map(blp_data, 'meter_number', 'h8')
        blp_h12_map = self._last_value_map(blp_data, 'meter_number', 'h12')
        blp_h24_map = self._last_value_map(blp_data, 'meter_number', 'h24')
        dlp_h24_map = self._last_value_map(dlp_data, 'meter_number', 'h24')
        bill_72_map = self._last_value_map(billing_data, 'meter_number', '72hrs')
        bill_168_map = self._last_value_map(billing_data, 'meter_number', '168hrs')
        
        def per_row(lookup, default=0):
            vals = pd.Series([lookup.get(m, default) for m in meter_list], dtype=object).infer_objects().to_numpy()
            return vals[codes]
        
        typ_u = _obj_array([excl_info[m][0] for m in meter_list])
        exc_text_u = _obj_array([self.get_exception_text(*excl_info[m]) for m in meter_list])
        exc_date_u = _obj_array([self.get_exception_date(*excl_info[m]) for m in meter_list])
        typ = typ_u[codes]
        exc_text = exc_text_u[codes]
        exc_date = exc_date_u[codes]
        
        is_po = typ == 'Power_OFF'
        inc = (typ == '-') | is_po
        excl = (typ == 'PD') | (typ == 'SM_SM') | (typ == 'Non_Opp')
        
        loc = _to_str(sat_col('location_code'))
        loc = loc.where(~loc.str.isdigit(), loc.str.zfill(4))
        cons = _to_str(sat_col('consumer_id'))
        loc_codes, loc_uniques = pd.factorize(loc)
        ami_u = [self.get_ami_loc_data(u) for u in loc_uniques.tolist()]
        zone = _obj_array([a[0] for a in ami_u])[loc_codes]
        circle = _obj_array([a[1] for a in ami_u])[loc_codes]
        division = _obj_array([a[2] for a in ami_u])[loc_codes]
        subdivision = _obj_array([a[3] for a in ami_u])[loc_codes]
        
        ip = np.array([float(x) if x else 0 for x in sat_col('supported_ip', 0).tolist()], dtype=float)
        
        sat_raw = sat_col('sat_pkg').astype(object).to_numpy()
        uid = (_to_str(sat_col('sat_pkg')) + '_' + loc).to_numpy()
        
        poc = per_row(po_map, 0).astype(float)
        
        with np.errstate(divide='ignore', invalid='ignore'):
            base_blp = ip * nod
            x = base_blp - poc
            blp_pof = np.where(0 > x, 0.0, x)
            y = np.where(ip > 0, nod - poc / ip, nod)
            dlp_pof = np.where(0 > y, 0.0, y)
        pof = is_po & (poc > 0)
        blp_tg = np.where(excl, 0.0, np.where(pof, blp_pof, base_blp))
        dlp_tg = np.where(excl, 0.0, np.where(pof, dlp_pof, float(nod)))
        
        def sla(lookup):
            return np.where(inc, per_row(lookup), 0)
        
        result_df = pd.DataFrame({
            'BU': loc.to_numpy(),
            'Zone': zone,
            'Circle': circle,
            'Division': division,
            'Subdivision': subdivision,
            'KNO': cons.to_numpy(),
            'meter_number': meter_s.to_numpy(),
            'HES': [hes_name] * n,
            'SAT_No': sat_raw,
            'UID': uid,
            'Exclusion_Type': typ,
            'Power_off_count': np.where(is_po, poc, 0),
            'IP': ip,
            'NOD': nod,
            'Exception': exc_text,
            'Exception_Date': exc_date,
            'Main_Ex_Flag': np.where(inc, '0', '1'),
            'BLP_Tg': blp_tg,
            'BLP_H8_SLA': sla(blp_h8_map),
            'BLP_H12_SLA': sla(blp_h12_map),
            'BLP_H24_SLA': sla(blp_h24_map),
            'DLP_Tg': dlp_tg,
            'DLP_H24_SLA': sla(dlp_h24_map),
            'Bill_Tg': np.where(inc, 1, 0),
            'Bill_h72_SLA': sla(bill_72_map),
            'Bill_h168_SLA': sla(bill_168_map),
        }, columns=['BU', 'Zone', 'Circle', 'Division', 'Subdivision', 'KNO', 'meter_number', 'HES',
                    'SAT_No', 'UID', 'Exclusion_Type', 'Power_off_count', 'IP', 'NOD', 'Exception',
                    'Exception_Date', 'Main_Ex_Flag', 'BLP_Tg', 'BLP_H8_SLA', 'BLP_H12_SLA', 'BLP_H24_SLA',
                    'DLP_Tg', 'DLP_H24_SLA', 'Bill_Tg', 'Bill_h72_SLA', 'Bill_h168_SLA'])
        return result_df
    
    def create_sdo_wise_summary(self, hes_wise_data, month_year, sat_numbers):
        """Create SDO_Wise_Summary sheet with ALL UID combinations"""
        self.log_message("  Creating SDO_Wise_Summary sheet...")
        
        if sat_numbers is None or not sat_numbers:
            sat_numbers = []
            for hes_name, hes_df in hes_wise_data.items():
                if 'SAT_No' in hes_df.columns:
                    sat_list = hes_df['SAT_No'].unique().tolist()
                    for sat in sat_list:
                        if sat and str(sat) not in sat_numbers:
                            sat_numbers.append(str(sat))
            sat_numbers = sorted(sat_numbers)
        
        all_bu_codes = []
        if self.ami_loc_df is not None:
            if 'sitecode' in self.ami_loc_df.columns:
                all_bu_codes = self.ami_loc_df['sitecode'].astype(str).str.zfill(4).unique().tolist()
                all_bu_codes = sorted(all_bu_codes)
        
        self.log_message(f"  Found {len(sat_numbers)} SAT numbers: {sat_numbers}")
        self.log_message(f"  Found {len(all_bu_codes)} BU codes from ami_loc")
        
        all_possible_uids = []
        for sat in sat_numbers:
            for bu in all_bu_codes:
                uid = f"{sat}_{bu}"
                all_possible_uids.append({
                    'sat_number': sat,
                    'bu_code': bu,
                    'uid': uid
                })
        
        self.log_message(f"  Total possible UIDs: {len(all_possible_uids)}")
        
        metric_cols = ['BLP_Tg', 'BLP_H8_SLA', 'BLP_H12_SLA', 'BLP_H24_SLA',
                       'DLP_Tg', 'DLP_H24_SLA', 'Bill_Tg', 'Bill_h72_SLA', 'Bill_h168_SLA']
        required_cols = ['UID'] + metric_cols
        
        uid_parts = []
        metric_parts = {c: [] for c in metric_cols}
        for hes_name, hes_df in hes_wise_data.items():
            missing_cols = [col for col in required_cols if col not in hes_df.columns]
            if missing_cols:
                self.log_message(f"  Warning: Missing columns in {hes_name}: {missing_cols}")
                continue
            uid_parts.append(_to_str(hes_df['UID']).to_numpy())
            for c in metric_cols:
                metric_parts[c].append(hes_df[c].astype(float).to_numpy())
        
        actual_data = {}
        if uid_parts:
            uid_all = np.concatenate(uid_parts)
            keep = uid_all != ''
            uid_codes, uid_uniques = pd.factorize(uid_all[keep])
            n_uid = len(uid_uniques)
            sums = {c: np.bincount(uid_codes, weights=np.concatenate(metric_parts[c])[keep], minlength=n_uid)
                    for c in metric_cols}
            counts = np.bincount(uid_codes, minlength=n_uid)
            uid_index = {u: i for i, u in enumerate(uid_uniques.tolist())}
        else:
            sums, counts, uid_index = {}, None, {}
        
        self.log_message(f"  Found data for {len(uid_index)} UIDs")
        
        sdo_rows = []
        
        for combo in all_possible_uids:
            uid = combo['uid']
            sat_number = combo['sat_number']
            bu_code = combo['bu_code']
            
            zone, circle, division, subdivision = self.get_ami_loc_data(bu_code)
            
            if uid in uid_index:
                i = uid_index[uid]
                blp_tg = sums['BLP_Tg'][i]
                blp_h8 = sums['BLP_H8_SLA'][i]
                blp_h12 = sums['BLP_H12_SLA'][i]
                blp_h24 = sums['BLP_H24_SLA'][i]
                dlp_tg = sums['DLP_Tg'][i]
                dlp_h24 = sums['DLP_H24_SLA'][i]
                bill_tg = sums['Bill_Tg'][i]
                bill_h72 = sums['Bill_h72_SLA'][i]
                bill_h168 = sums['Bill_h168_SLA'][i]
                meter_no_count = int(counts[i])
            else:
                blp_tg = blp_h8 = blp_h12 = blp_h24 = 0
                dlp_tg = dlp_h24 = bill_tg = bill_h72 = bill_h168 = 0
                meter_no_count = 0
            
            billing_month = month_year[:2] if month_year else ''
            billing_year = f"20{month_year[2:]}" if month_year else ''
            
            row = {
                'UID': uid,
                'zone': zone,
                'Circle': circle,
                'division': division,
                'subdivision': subdivision,
                'BU': bu_code,
                'amisp_name': 'MCL',
                'pkg_name': 'NAGPUR',
                'billing_month': billing_month,
                'billing_year': billing_year,
                'sat_number': sat_number,
                'expected_interval_count': blp_tg,
                'interval_data_08_hr': blp_h8,
                'interval_data_12_hr': blp_h12,
                'interval_data_24_hr': blp_h24,
                'expected_daily_count': dlp_tg,
                'daily_data_24_hr': dlp_h24,
                'expected_billing_count': bill_tg,
                'billing_data_72_hr': bill_h72,
                'billing_data_168_hr': bill_h168,
                'meter_no_count': meter_no_count
            }
            sdo_rows.append(row)
        
        sdo_df = pd.DataFrame(sdo_rows)
        self.log_message(f"  Created SDO_Wise_Summary with {len(sdo_df)} records (all UID combinations)")
        return sdo_df
    
    def create_final_excel(self):
        """Create the final Excel with all required sheets"""
        self.log_message("Creating Final Summary Excel...")
        
        try:
            if self.fetch_var.get() == "YES":
                if self.from_date_selected is None:
                    self.from_date_selected = self.from_date_entry.get()
                self.log_message(f"Using from_date from entry: {self.from_date_selected}")
            else:
                if self.month_year_from_filename:
                    from_date_from_filename = self.get_from_date_from_mmyy(self.month_year_from_filename)
                    if from_date_from_filename:
                        self.from_date_selected = from_date_from_filename
                        self.log_message(f"NOD (MMYY {self.month_year_from_filename})")
                    else:
                        self.from_date_selected = self.from_date_entry.get()
                        self.log_message(f"Could not extract date from filename, using from_date from entry: {self.from_date_selected}")
                else:
                    self.from_date_selected = self.from_date_entry.get()
                    self.log_message(f"Using from_date from entry: {self.from_date_selected}")
            
            self.log_message("\nLoading raw data files...")
            self.log_message(f"Excel reader: {EXCEL_ENGINE}")
            
            self.log_message("Loading SATC, BLP, DLP and Billing Data...")
            with ThreadPoolExecutor(max_workers=4) as pool:
                f_satc = pool.submit(read_all_sheets, self.satc_file_path)
                f_blp = pool.submit(read_all_sheets, self.blp_file_path)
                f_dlp = pool.submit(read_all_sheets, self.dlp_file_path)
                f_billing = pool.submit(read_all_sheets, self.billing_file_path)
                satc_data = f_satc.result()
                blp_data = f_blp.result()
                dlp_data = f_dlp.result()
                billing_data = f_billing.result()
            
            self.log_message("Loading Power OFF Data...")
            power_off_data = pd.read_excel(self.power_off_file_path, engine=EXCEL_ENGINE)
            
            # Always load RC_DC and SM to SM data (regardless of exclusion setting)
            self.log_message("Loading RC_DC Data...")
            rc_dc_data = pd.read_excel(self.rc_dc_file_path, engine=EXCEL_ENGINE)
            
            self.log_message("Loading SM to SM Data...")
            sm_to_sm_data = pd.read_excel(self.sm_to_sm_file_path, engine=EXCEL_ENGINE)
            
            # Only load other exclusion-related files if Exclusion is YES
            if self.exclusion_enabled:
                self.log_message("Loading PDRC Data...")
                pdrc_data = pd.read_excel(self.pdrc_file_path, engine=EXCEL_ENGINE)
                
                self.log_message("Loading Non Opp Data...")
                non_opp_data = pd.read_excel(self.non_opp_file_path, engine=EXCEL_ENGINE)
            else:
                pdrc_data = pd.DataFrame()
                non_opp_data = pd.DataFrame()
                self.log_message("⚠ Exclusion is DISABLED. Exclusion-related files (PDC, Non Opp) will be skipped.")
                self.log_message("✓ Power OFF, RC_DC and SM to SM data will be used.")
            
            if self.month_year_from_filename:
                nod = self.get_nod_from_month_year(self.month_year_from_filename)
                self.log_message(f"\nNumber of Days (NOD) : {nod}")
            else:
                nod = self.calculate_nod(self.from_date_selected)
                self.log_message(f"\nNumber of Days (NOD) : {nod}")
            
            satc_frames = []
            for sheet_name, sheet_data in satc_data.items():
                sheet_copy = sheet_data.copy()
                if 'location_code' in sheet_copy.columns:
                    sheet_copy['location_code'] = sheet_copy['location_code'].astype(str).str.zfill(4)
                if 'consumer_id' in sheet_copy.columns:
                    sheet_copy['consumer_id'] = sheet_copy['consumer_id'].astype(str)
                satc_frames.append(sheet_copy)
            all_satc_data = pd.concat(satc_frames, ignore_index=True) if satc_frames else pd.DataFrame()
            del satc_data, satc_frames
            
            blp_by_hes = {}
            for sheet_name, sheet_data in blp_data.items():
                blp_by_hes[sheet_name] = sheet_data
            
            dlp_by_hes = {}
            for sheet_name, sheet_data in dlp_data.items():
                dlp_by_hes[sheet_name] = sheet_data
            
            billing_by_hes = {}
            for sheet_name, sheet_data in billing_data.items():
                billing_by_hes[sheet_name] = sheet_data
            
            hes_list = ['BCITS', 'Cyan', 'Crystal']
            hes_wise_data = {}
            
            for hes_name in hes_list:
                self.log_message(f"\nCreating SLA_{hes_name}_Raw sheet...")
                
                blp_hes_data = blp_by_hes.get(hes_name, pd.DataFrame())
                dlp_hes_data = dlp_by_hes.get(hes_name, pd.DataFrame())
                billing_hes_data = billing_by_hes.get(hes_name, pd.DataFrame())
                
                hes_data = self.create_hes_wise_raw_data(
                    hes_name, all_satc_data, blp_hes_data, dlp_hes_data, billing_hes_data,
                    pdrc_data, sm_to_sm_data, non_opp_data, power_off_data, self.from_date_selected, nod
                )
                
                if not hes_data.empty:
                    if 'BU' in hes_data.columns:
                        hes_data['BU'] = hes_data['BU'].astype(str).str.zfill(4)
                    if 'KNO' in hes_data.columns:
                        hes_data['KNO'] = hes_data['KNO'].astype(str)
                    hes_wise_data[f'SLA_{hes_name}_Raw'] = hes_data
                    self.log_message(f"  Created SLA_{hes_name}_Raw with {len(hes_data)} records")
            
            sat_numbers = []
            for hes_name, hes_df in hes_wise_data.items():
                if 'SAT_No' in hes_df.columns:
                    sat_list = hes_df['SAT_No'].unique().tolist()
                    for sat in sat_list:
                        if sat and sat not in sat_numbers:
                            sat_numbers.append(str(sat))
            sat_numbers = sorted(sat_numbers)
            self.log_message(f"\n  SAT Numbers found: {sat_numbers}")
            
            month_year = self.get_month_year_from_date(self.from_date_selected)
            current_date = datetime.datetime.now().strftime("%d%m%y")
            
            if self.exclusion_enabled:
                output_filename = os.path.join(self.script_dir, f"SLA_HES_Wise_Final_{month_year}_{current_date}.xlsx")
            else:
                output_filename = os.path.join(self.script_dir, f"SLA_Cumm_Worksheet_Final_{month_year}_{current_date}.xlsx")
            
            self.log_message(f"\nSaving final Excel...!!")
            self.log_message(f"Filename: {os.path.basename(output_filename)}")
            
            with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
                hes_summary = []
                
                for hes_name, hes_df in hes_wise_data.items():
                    hes_display_name = hes_name.replace('SLA_', '').replace('_Raw', '')
                    total_meters = len(hes_df)
                    
                    if self.exclusion_enabled:
                        pdc_count = len(hes_df[hes_df['Exclusion_Type'] == 'PD'])
                        non_opp_count = len(hes_df[hes_df['Exclusion_Type'] == 'Non_Opp'])
                        sm_sm_count = len(hes_df[hes_df['Exclusion_Type'] == 'SM_SM'])
                        power_off_count = len(hes_df[hes_df['Exclusion_Type'] == 'Power_OFF'])
                        eligible_meters = total_meters - (pdc_count + non_opp_count + sm_sm_count)
                    else:
                        pdc_count = 0
                        non_opp_count = 0
                        sm_sm_count = 0
                        power_off_count = len(hes_df[hes_df['Exclusion_Type'] == 'Power_OFF'])
                        eligible_meters = total_meters
                    
                    if self.exclusion_enabled:
                        active_df = hes_df[hes_df['Main_Ex_Flag'] == '0']
                    else:
                        active_df = hes_df
                    
                    blp_tg = active_df['BLP_Tg'].sum() if not active_df.empty else 0
                    blp_8 = active_df['BLP_H8_SLA'].sum() if not active_df.empty else 0
                    blp_12 = active_df['BLP_H12_SLA'].sum() if not active_df.empty else 0
                    blp_24 = active_df['BLP_H24_SLA'].sum() if not active_df.empty else 0
                    dlp_tg = active_df['DLP_Tg'].sum() if not active_df.empty else 0
                    dlp_24 = active_df['DLP_H24_SLA'].sum() if not active_df.empty else 0
                    bill_tg = active_df['Bill_Tg'].sum() if not active_df.empty else 0
                    bill_3 = active_df['Bill_h72_SLA'].sum() if not active_df.empty else 0
                    bill_7 = active_df['Bill_h168_SLA'].sum() if not active_df.empty else 0
                    
                    pct_blp_8 = (blp_8 / blp_tg * 100) if blp_tg > 0 else 0
                    pct_blp_12 = (blp_12 / blp_tg * 100) if blp_tg > 0 else 0
                    pct_blp_24 = (blp_24 / blp_tg * 100) if blp_tg > 0 else 0
                    pct_dlp_24 = (dlp_24 / dlp_tg * 100) if dlp_tg > 0 else 0
                    pct_bill_3 = (bill_3 / bill_tg * 100) if bill_tg > 0 else 0
                    pct_bill_7 = (bill_7 / bill_tg * 100) if bill_tg > 0 else 0
                    
                    if self.exclusion_enabled:
                        summary_row = {
                            'HES': hes_display_name,
                            'Mtr_Qty': total_meters,
                            'PDC': pdc_count,
                            'Non_Opp': non_opp_count,
                            'SM_SM': sm_sm_count,
                            'Power_OFF': power_off_count,
                            'Eligible_meters': eligible_meters,
                            'BLP_Tg': blp_tg,
                            'BLP_8': blp_8,
                            'BLP_12': blp_12,
                            'BLP_24': blp_24,
                            'DLP_Tg': dlp_tg,
                            'DLP_24': dlp_24,
                            'Bill_Tg': bill_tg,
                            'Bill_3': bill_3,
                            'Bill_7': bill_7,
                            '%BLP_8': pct_blp_8,
                            '%BLP_12': pct_blp_12,
                            '%BLP_24': pct_blp_24,
                            '%DLP_24': pct_dlp_24,
                            '%Bill_3': pct_bill_3,
                            '%Bill_7': pct_bill_7
                        }
                    else:
                        summary_row = {
                            'HES': hes_display_name,
                            'Mtr_Qty': total_meters,
                            'Power_OFF': power_off_count,
                            'Eligible_meters': eligible_meters,
                            'BLP_Tg': blp_tg,
                            'BLP_8': blp_8,
                            'BLP_12': blp_12,
                            'BLP_24': blp_24,
                            'DLP_Tg': dlp_tg,
                            'DLP_24': dlp_24,
                            'Bill_Tg': bill_tg,
                            'Bill_3': bill_3,
                            'Bill_7': bill_7,
                            '%BLP_8': pct_blp_8,
                            '%BLP_12': pct_blp_12,
                            '%BLP_24': pct_blp_24,
                            '%DLP_24': pct_dlp_24,
                            '%Bill_3': pct_bill_3,
                            '%Bill_7': pct_bill_7
                        }
                    
                    hes_summary.append(summary_row)
                
                if hes_summary:
                    if self.exclusion_enabled:
                        all_hes_row = {
                            'HES': 'All HES -',
                            'Mtr_Qty': sum(row['Mtr_Qty'] for row in hes_summary),
                            'PDC': sum(row['PDC'] for row in hes_summary),
                            'Non_Opp': sum(row['Non_Opp'] for row in hes_summary),
                            'SM_SM': sum(row['SM_SM'] for row in hes_summary),
                            'Power_OFF': sum(row['Power_OFF'] for row in hes_summary),
                            'Eligible_meters': sum(row['Eligible_meters'] for row in hes_summary),
                            'BLP_Tg': sum(row['BLP_Tg'] for row in hes_summary),
                            'BLP_8': sum(row['BLP_8'] for row in hes_summary),
                            'BLP_12': sum(row['BLP_12'] for row in hes_summary),
                            'BLP_24': sum(row['BLP_24'] for row in hes_summary),
                            'DLP_Tg': sum(row['DLP_Tg'] for row in hes_summary),
                            'DLP_24': sum(row['DLP_24'] for row in hes_summary),
                            'Bill_Tg': sum(row['Bill_Tg'] for row in hes_summary),
                            'Bill_3': sum(row['Bill_3'] for row in hes_summary),
                            'Bill_7': sum(row['Bill_7'] for row in hes_summary),
                        }
                    else:
                        all_hes_row = {
                            'HES': 'All HES -',
                            'Mtr_Qty': sum(row['Mtr_Qty'] for row in hes_summary),
                            'Power_OFF': sum(row['Power_OFF'] for row in hes_summary),
                            'Eligible_meters': sum(row['Eligible_meters'] for row in hes_summary),
                            'BLP_Tg': sum(row['BLP_Tg'] for row in hes_summary),
                            'BLP_8': sum(row['BLP_8'] for row in hes_summary),
                            'BLP_12': sum(row['BLP_12'] for row in hes_summary),
                            'BLP_24': sum(row['BLP_24'] for row in hes_summary),
                            'DLP_Tg': sum(row['DLP_Tg'] for row in hes_summary),
                            'DLP_24': sum(row['DLP_24'] for row in hes_summary),
                            'Bill_Tg': sum(row['Bill_Tg'] for row in hes_summary),
                            'Bill_3': sum(row['Bill_3'] for row in hes_summary),
                            'Bill_7': sum(row['Bill_7'] for row in hes_summary),
                        }
                    
                    all_hes_row['%BLP_8'] = (all_hes_row['BLP_8'] / all_hes_row['BLP_Tg'] * 100) if all_hes_row['BLP_Tg'] > 0 else 0
                    all_hes_row['%BLP_12'] = (all_hes_row['BLP_12'] / all_hes_row['BLP_Tg'] * 100) if all_hes_row['BLP_Tg'] > 0 else 0
                    all_hes_row['%BLP_24'] = (all_hes_row['BLP_24'] / all_hes_row['BLP_Tg'] * 100) if all_hes_row['BLP_Tg'] > 0 else 0
                    all_hes_row['%DLP_24'] = (all_hes_row['DLP_24'] / all_hes_row['DLP_Tg'] * 100) if all_hes_row['DLP_Tg'] > 0 else 0
                    all_hes_row['%Bill_3'] = (all_hes_row['Bill_3'] / all_hes_row['Bill_Tg'] * 100) if all_hes_row['Bill_Tg'] > 0 else 0
                    all_hes_row['%Bill_7'] = (all_hes_row['Bill_7'] / all_hes_row['Bill_Tg'] * 100) if all_hes_row['Bill_Tg'] > 0 else 0
                    
                    hes_summary.append(all_hes_row)
                
                hes_summary_df = pd.DataFrame(hes_summary)
                
                pct_columns = ['%BLP_8', '%BLP_12', '%BLP_24', '%DLP_24', '%Bill_3', '%Bill_7']
                for col in pct_columns:
                    if col in hes_summary_df.columns:
                        hes_summary_df[col] = hes_summary_df[col].apply(lambda x: f"{x:.3f}%" if pd.notna(x) else "0.00%")
                
                if not hes_summary_df.empty:
                    hes_summary_df.to_excel(writer, sheet_name='Hes_Wise_Summary', index=False)
                    self.log_message("  Created Hes_Wise_Summary sheet")
                
                # Only create SDO_Wise_Summary if Exclusion is YES
                if self.exclusion_enabled:
                    sdo_summary_df = self.create_sdo_wise_summary(hes_wise_data, month_year, sat_numbers)
                    if not sdo_summary_df.empty:
                        sdo_summary_df.to_excel(writer, sheet_name='SDO_Wise_Summary', index=False)
                        self.log_message("  Created SDO_Wise_Summary sheet")
                
                # Write RC_DC sheet (always included)
                if not rc_dc_data.empty:
                    rc_dc_data.to_excel(writer, sheet_name='RC_DC', index=False)
                    self.log_message("  Created RC_DC sheet")
                
                # Write SM_to_SM sheet (always included)
                if not sm_to_sm_data.empty:
                    sm_to_sm_data.to_excel(writer, sheet_name='SM_to_SM', index=False)
                    self.log_message("  Created SM_to_SM sheet")
                
                # Write SLA_BCITS_Raw, SLA_Cyan_Raw, SLA_Crystal_Raw
                for sheet_name, sheet_data in hes_wise_data.items():
                    if 'BU' in sheet_data.columns:
                        sheet_data['BU'] = sheet_data['BU'].astype(str).str.zfill(4)
                    if 'KNO' in sheet_data.columns:
                        sheet_data['KNO'] = sheet_data['KNO'].astype(str)
                    sheet_data.to_excel(writer, sheet_name=sheet_name, index=False)
                
                # Write Power_OFF_RAW (always included)
                if not power_off_data.empty:
                    power_off_data.to_excel(writer, sheet_name='Power_OFF_RAW', index=False)
                    self.log_message("  Created Power_OFF_RAW sheet")
                
                # Only write other exclusion-related sheets if Exclusion is YES
                if self.exclusion_enabled:
                    if not pdrc_data.empty:
                        pdrc_data.to_excel(writer, sheet_name='PDC_RAW', index=False)
                        self.log_message("  Created PDC_RAW sheet")
                    
                    if not non_opp_data.empty:
                        non_opp_data.to_excel(writer, sheet_name='Non_Opp_Raw', index=False)
                        self.log_message("  Created Non_Opp_Raw sheet")
            
            self.log_message("\n" + "=" * 60)
            self.log_message(f"✓ Final Summary Process completed successfully!")
            self.log_message(f"Output saved to: {output_filename}")
            self.log_message("=" * 60)
            
            messagebox.showinfo("Success", f"Final Summary Process completed successfully!\n\nOutput saved to:\n{output_filename}")
            
        except Exception as e:
            self.log_message(f"Error in final summary process: {str(e)}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"Failed to run final summary process:\n{str(e)}")
    
    def run_final_process(self):
        """Run the final summary process"""
        if self.final_summary_var.get() == "NO":
            self.log_message("Final summary process cancelled.")
            return
        
        # Check required files based on exclusion setting
        if self.exclusion_enabled:
            if not all([self.satc_file_path, self.blp_file_path, self.dlp_file_path, self.billing_file_path, 
                       self.pdrc_file_path, self.sm_to_sm_file_path, self.non_opp_file_path, self.power_off_file_path, 
                       self.rc_dc_file_path]):
                self.log_message("Error: Please select all required files for the final summary.")
                messagebox.showerror("Error", "Please select all required files for the final summary.\n\nRequired files:\n- SATC RAW\n- BLP RAW\n- DLP RAW\n- BILLING RAW\n- PDRC RAW\n- SM to SM RAW\n- Non Opp RAW\n- Power OFF RAW\n- RC_DC RAW")
                return
        else:
            if not all([self.satc_file_path, self.blp_file_path, self.dlp_file_path, self.billing_file_path, 
                       self.power_off_file_path, self.rc_dc_file_path, self.sm_to_sm_file_path]):
                self.log_message("Error: Please select all required files for the final summary.")
                messagebox.showerror("Error", "Please select all required files for the final summary.\n\nRequired files:\n- SATC RAW\n- BLP RAW\n- DLP RAW\n- BILLING RAW\n- Power OFF RAW\n- RC_DC RAW\n- SM to SM RAW")
                return
        
        process_thread = threading.Thread(target=self.create_final_excel)
        process_thread.daemon = True
        process_thread.start()