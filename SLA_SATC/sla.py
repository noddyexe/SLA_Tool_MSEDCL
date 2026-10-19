import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import psycopg2
import pandas as pd
import os
import datetime
import socket
import sys
import time
import threading
import re
from config_loader import get_config
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))



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

        # NOTE: this is the SAME self.button_frame used everywhere else below —
        # nothing else in the file needs to change.
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
        self.from_date_selected = None  # Store from_date for final summary
        self.to_date_selected = None    # Store to_date for final summary
        self.month_year_from_filename = None  # Store month_year from filename
        self.exclusion_enabled = True  # Default to Yes (Exclusion enabled)
        
        # File paths for final summary
        self.satc_file_path = ""
        self.blp_file_path = ""
        self.dlp_file_path = ""
        self.billing_file_path = ""
        self.pdrc_file_path = ""
        self.sm_to_sm_file_path = ""
        self.non_opp_file_path = ""
        self.power_off_file_path = ""
        self.rc_dc_file_path = ""  # New file path for RC_DC
        
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
        
        # New RC_DC button
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
            self.log_message("Note: PDC, SM to SM, Non-Operational, and RC_DC queries will be skipped")
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
                # Hide exclusion-related buttons (except Power OFF)
                self.pdrc_button.grid_remove()
                self.sm_to_sm_button.grid_remove()
                self.non_opp_button.grid_remove()
                self.rc_dc_button.grid_remove()
                
                # Show only required buttons + Power OFF
                self.satc_button.grid(row=0, column=0, columnspan=3, pady=5, sticky=tk.EW)
                self.blp_button.grid(row=1, column=0, columnspan=3, pady=5, sticky=tk.EW)
                self.dlp_button.grid(row=2, column=0, columnspan=3, pady=5, sticky=tk.EW)
                self.billing_button.grid(row=3, column=0, columnspan=3, pady=5, sticky=tk.EW)
                # Power OFF is already at row=7
                
                # Enable only visible buttons
                self.satc_button.config(state=tk.NORMAL)
                self.blp_button.config(state=tk.NORMAL)
                self.dlp_button.config(state=tk.NORMAL)
                self.billing_button.config(state=tk.NORMAL)
                self.power_off_button.config(state=tk.NORMAL)
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
        
        # Close RC_DC connection as well
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
                
                # Fetch results in batches
                results = []
                while True:
                    rows = cur.fetchmany(self.BATCH_SIZE)
                    if not rows:
                        break
                    
                    # Get column names
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
                
                # Fetch results in batches
                results = []
                while True:
                    rows = cur.fetchmany(self.BATCH_SIZE)
                    if not rows:
                        break
                    
                    # Get column names
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
                
                # Close RC_DC connection on failure
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
    
    def get_nod_from_month_year(self, month_year_str):
        """Calculate NOD from MMYY format"""
        try:
            month = int(month_year_str[:2])
            year = int(month_year_str[2:])
            
            if month in [1, 3, 5, 7, 8, 10, 12]:
                return 31
            elif month in [4, 6, 9, 11]:
                return 30
            else:  # February
                # Check for leap year
                if (year % 4 == 0 and year % 100 != 0) or (year % 400 == 0):
                    return 29
                else:
                    return 28
        except:
            # Fallback to date-based calculation
            return 30
    
    def extract_month_year_from_filename(self, filename):
        """Extract MMYY from filename like SATC_RAW_0526_20260703"""
        try:
            # Look for pattern like _0526_ in filename
            match = re.search(r'_(\d{4})_', filename)
            if match:
                return match.group(1)
            
            # Alternative: look for MMYY at end
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
            
            # Add 2000 to year if it's 2 digits
            if year < 100:
                year = 2000 + year
            
            # Return first day of the month
            return f"{year}-{month:02d}-01"
        except:
            return None
    
    def save_excel_by_hes(self, df, base_filename):
        """Save Excel file with separate sheets for each HES name (no All_Data sheet)"""
        if df is None or df.empty:
            self.log_message(f"No data to save for {base_filename}")
            return
        
        current_date = datetime.datetime.now().strftime("%Y%m%d")
        filename = os.path.join(self.save_directory, f"{base_filename}_{current_date}.xlsx")
        
        try:
            # Check if hes column exists (for SATC data it's 'hes', for others it's 'hes_name')
            hes_column = None
            if 'hes_name' in df.columns:
                hes_column = 'hes_name'
            elif 'hes' in df.columns:
                hes_column = 'hes'
            
            if hes_column:
                hes_names = df[hes_column].unique()
                
                with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
                    # Write data for each HES to separate sheet only
                    for hes in hes_names:
                        if hes and str(hes) != '0' and str(hes) != 'nan':
                            hes_data = df[df[hes_column] == hes].copy()
                            sheet_name = str(hes)[:31]  # Excel sheet name max 31 chars
                            # Clean sheet name for Excel (remove invalid characters)
                            sheet_name = sheet_name.replace('/', '_').replace('\\', '_').replace('*', '_').replace('?', '_').replace(':', '_')
                            
                            # Format location_code and consumer_id as text to preserve leading zeros
                            if 'location_code' in hes_data.columns:
                                hes_data['location_code'] = hes_data['location_code'].astype(str).str.zfill(4)
                            if 'consumer_id' in hes_data.columns:
                                hes_data['consumer_id'] = hes_data['consumer_id'].astype(str)
                            
                            hes_data.to_excel(writer, sheet_name=sheet_name, index=False)
                            self.log_message(f"'{sheet_name}' with {len(hes_data)} records")
                    
                self.log_message(f"Data saved to: {filename}")
            else:
                # If no hes column, save normally with formatting
                df_copy = df.copy()
                if 'location_code' in df_copy.columns:
                    df_copy['location_code'] = df_copy['location_code'].astype(str).str.zfill(4)
                if 'consumer_id' in df_copy.columns:
                    df_copy['consumer_id'] = df_copy['consumer_id'].astype(str)
                df_copy.to_excel(filename, index=False)
                self.log_message(f"Data saved to: {filename}")
                
        except Exception as e:
            self.log_message(f"Error saving Excel file: {str(e)}")
            # Fallback to CSV if Excel fails
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
        WHERE satc   = 'SATC'
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
        """Run first query for BLP data - UPDATED"""
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
        """Run second query for DLP data - UPDATED"""
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
       # Parse the from_date to get year and month
       from_date_obj = datetime.datetime.strptime(from_date, "%Y-%m-%d")
       year = from_date_obj.year
       month = from_date_obj.month
       
       # Calculate date1_start as 3 days before the 1st of the month
       # First day of the month
       first_day = datetime.datetime(year, month, 1)
       # Subtract 3 days
       date1_start_dt = first_day - datetime.timedelta(days=3)
       date1_start = date1_start_dt.strftime("%Y-%m-%d %H:%M:%S")
       
       # Calculate the date ranges based on from_date
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
    
    def run_query_5_sm_to_sm(self):
        """Run fifth query for SM to SM data"""
        query = """
        SELECT * FROM unified_hes.uni_non_operational_meters 
        WHERE exception LIKE 'SM%'
        """
        
        results = self.execute_query_with_retry(query, "SM to SM Query")
        
        if results:
            df = pd.DataFrame(results)
            self.log_message(f"Total records fetched: {len(df)}")
            return df
        else:
            self.log_message("SM to SM Query completed but no data found")
            return pd.DataFrame()
    
    def run_query_6_non_opp(self):
        """Run sixth query for Non-Operational data"""
        query = """
        SELECT * FROM unified_hes.uni_non_operational_meters 
        WHERE exception NOT LIKE 'SM%'
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
        # Format dates for query
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
    
    def run_query_8_rc_dc(self, from_date, to_date):
        """Run eighth query for RC_DC data using separate database connection"""
        # Format dates for query
        from_date_start = f"{from_date} 00:00:00"
        to_date_end = f"{to_date} 23:59:59"
        
        query = f"""
        WITH eligible_meters AS (
            SELECT
                m.mtrno,
                m.sdocode,
                m.village AS sat_number
            FROM ami_master.master_main m
            WHERE m.comm IN ('0', '1')
              and m.mla = 'SAT'
        ),
        rc_dc_data AS (
            SELECT
                em.sdocode,
                em.sat_number,
                em.mtrno,
                r.id,
                r.request_time,
                r.success_time,
                EXTRACT(EPOCH FROM (r.first_try_time::timestamp - r.request_time::timestamp)) / 60
                           AS duration_minutes
            FROM eligible_meters em
            INNER JOIN ami_master.rc_dc_log r
                ON  em.mtrno  = r.mtrno
                AND r.status  = 'S'
                AND r.first_try_time IS NOT NULL
                AND r.request_time IS NOT NULL
                AND r.request_time >= '{from_date_start}'
                AND r.request_time <= '{to_date_end}'
        )
        SELECT
            rd.sdocode                                                        AS "SDO Code",
            rd.sat_number                                                     AS "SAT Number",
            COUNT(rd.id)                                                      AS "RC DC Count",
            COUNT(CASE WHEN rd.duration_minutes <= 15  THEN rd.id END)       AS "DISCONN_RECONN_15MIN",
            COUNT(CASE WHEN rd.duration_minutes <= 360 THEN rd.id END)       AS "DISCONN_RECONN_360MIN"
        FROM rc_dc_data rd
        GROUP BY
            rd.sdocode,
            rd.sat_number
        ORDER BY
            rd.sdocode,
            rd.sat_number
        """
        
        results = self.execute_rc_dc_query_with_retry(query, "RC_DC Query")
        
        if results:
            df = pd.DataFrame(results)
            self.log_message(f"Total records fetched: {len(df)}")
            return df
        else:
            self.log_message("RC_DC Query completed but no data found")
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
                if not df_power_off.empty:
                    self.save_dataframe(df_power_off, f"Power_OFF_RAW_{month_year}")
            
            # Only run other exclusion-related queries if Exclusion is YES
            if self.exclusion_enabled:
                # Query 4: PDC Data (now 5th overall)
                if not self.stop_processing:
                    self.log_message("\n[5/8] Fetching PDC data...")
                    df_pdc = self.run_query_4_pdc()
                    if not df_pdc.empty:
                        self.save_dataframe(df_pdc, f"PDC_RAW_{month_year}")
                
                # Query 5: SM to SM Data (now 6th overall)
                if not self.stop_processing:
                    self.log_message("\n[6/8] Fetching SM to SM data...")
                    df_sm = self.run_query_5_sm_to_sm()
                    if not df_sm.empty:
                        self.save_dataframe(df_sm, f"SM_to_SM_RAW_{month_year}")
                
                # Query 6: Non-Operational Data (now 7th overall)
                if not self.stop_processing:
                    self.log_message("\n[7/8] Fetching Non-Operational data...")
                    df_non_opp = self.run_query_6_non_opp()
                    if not df_non_opp.empty:
                        self.save_dataframe(df_non_opp, f"NON_OPP_RAW_{month_year}")
                
                # Query 8: RC_DC Data (now 8th overall)
                if not self.stop_processing:
                    self.log_message("\n[8/8] Fetching RC_DC data...")
                    df_rc_dc = self.run_query_8_rc_dc(from_date, to_date)
                    if not df_rc_dc.empty:
                        self.save_dataframe(df_rc_dc, f"RC_DC_RAW_{month_year}")
            else:
                self.log_message("\n⚠ Exclusion is DISABLED. Skipping PDC, SM to SM, Non-Operational, and RC_DC queries.")
                self.log_message("✓ Power OFF query has been executed.")
            
            # Process completion
            if not self.stop_processing:
                self.log_message("\nDo you want final Cumulative Summary..?")
                self.summary_frame.grid(row=6, column=0, columnspan=3, pady=10, sticky=tk.EW)
            else:
                self.log_message("\n✗ Process stopped by user.")
        
        self.cleanup_after_process()
        self.close_db_connection()
    
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
            # Extract month_year from filename
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
        else:  # February
            # Check for leap year
            if (year % 4 == 0 and year % 100 != 0) or (year % 400 == 0):
                return 29
            else:
                return 28
    
    def check_exclusion_type(self, meter_number, pdc_df, sm_to_sm_df, non_opp_df, power_off_df, from_date):
        """Check exclusion type for a meter number"""
        # If exclusion is disabled, skip all checks except Power OFF
        if not self.exclusion_enabled:
            # Still check Power OFF even when exclusion is disabled
            if not power_off_df.empty and 'meterno' in power_off_df.columns:
                power_off_match = power_off_df[power_off_df['meterno'].astype(str) == str(meter_number)]
                if not power_off_match.empty:
                    return 'Power_OFF', power_off_match.iloc[0]
            return '-', None
        
        # Check in PDC RAW
        if not pdc_df.empty and 'meterno' in pdc_df.columns:
            pdc_match = pdc_df[pdc_df['meterno'].astype(str) == str(meter_number)]
            if not pdc_match.empty:
                return 'PD', pdc_match.iloc[0]
        
        # Parse date ranges for closed_time check
        from_date_obj = datetime.datetime.strptime(from_date, "%Y-%m-%d")
        
        # Calculate month end date
        if from_date_obj.month == 12:
            month_end = datetime.datetime(from_date_obj.year + 1, 1, 1) - datetime.timedelta(days=1)
        else:
            month_end = datetime.datetime(from_date_obj.year, from_date_obj.month + 1, 1) - datetime.timedelta(days=1)
        
        month_end_with_time = month_end.replace(hour=23, minute=59, second=59)
        month_start_with_time = from_date_obj.replace(hour=0, minute=0, second=0)
        
        # Check in SM to SM RAW
        if not sm_to_sm_df.empty and 'meter_serial_number' in sm_to_sm_df.columns:
            sm_match = sm_to_sm_df[sm_to_sm_df['meter_serial_number'].astype(str) == str(meter_number)]
            if not sm_match.empty:
                ticket_status = sm_match.iloc[0].get('ticket_status', '')
                if ticket_status == 'CLOSED':
                    closed_time = sm_match.iloc[0].get('closed_time', '')
                    if closed_time and pd.notna(closed_time):
                        try:
                            closed_time_dt = pd.to_datetime(closed_time)
                            if month_start_with_time <= closed_time_dt <= month_end_with_time:
                                return 'SM_SM', sm_match.iloc[0]
                        except:
                            pass
                elif ticket_status == 'OPEN':
                    return 'SM_SM', sm_match.iloc[0]
        
        # Check in NON OPP RAW
        if not non_opp_df.empty and 'meter_serial_number' in non_opp_df.columns:
            non_opp_match = non_opp_df[non_opp_df['meter_serial_number'].astype(str) == str(meter_number)]
            if not non_opp_match.empty:
                ticket_status = non_opp_match.iloc[0].get('ticket_status', '')
                if ticket_status == 'CLOSED':
                    closed_time = non_opp_match.iloc[0].get('closed_time', '')
                    if closed_time and pd.notna(closed_time):
                        try:
                            closed_time_dt = pd.to_datetime(closed_time)
                            if month_start_with_time <= closed_time_dt <= month_end_with_time:
                                return 'Non_Opp', non_opp_match.iloc[0]
                        except:
                            pass
                elif ticket_status == 'OPEN':
                    non_operation_date = non_opp_match.iloc[0].get('non_operation_date', '')
                    if non_operation_date and pd.notna(non_operation_date):
                        try:
                            non_op_dt = pd.to_datetime(non_operation_date)
                            if non_op_dt <= month_end_with_time:
                                return 'Non_Opp', non_opp_match.iloc[0]
                        except:
                            pass
        
        # Check in Power OFF RAW
        if not power_off_df.empty and 'meterno' in power_off_df.columns:
            power_off_match = power_off_df[power_off_df['meterno'].astype(str) == str(meter_number)]
            if not power_off_match.empty:
                return 'Power_OFF', power_off_match.iloc[0]
        
        return '-', None
    
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
        
        # Convert bu_code to string and pad with zeros to 4 digits if it's numeric
        bu_str = str(bu_code)
        if bu_str.isdigit():
            bu_str = bu_str.zfill(4)
        
        # Find matching row in ami_loc
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
        """Create HES-wise raw data sheet"""
        # Filter SATC data for specific HES
        if 'hes' in satc_data.columns:
            hes_satc = satc_data[satc_data['hes'] == hes_name].copy()
        else:
            hes_satc = satc_data.copy()
        
        if hes_satc.empty:
            self.log_message(f"No data found for HES: {hes_name}")
            return pd.DataFrame()
        
        # Prepare BLP lookup
        blp_lookup = {}
        if not blp_data.empty and 'meter_number' in blp_data.columns:
            for _, row in blp_data.iterrows():
                meter_no = str(row.get('meter_number', ''))
                if meter_no:
                    blp_lookup[meter_no] = row
        
        # Prepare DLP lookup
        dlp_lookup = {}
        if not dlp_data.empty and 'meter_number' in dlp_data.columns:
            for _, row in dlp_data.iterrows():
                meter_no = str(row.get('meter_number', ''))
                if meter_no:
                    dlp_lookup[meter_no] = row
        
        # Prepare Billing lookup
        billing_lookup = {}
        if not billing_data.empty and 'meter_number' in billing_data.columns:
            for _, row in billing_data.iterrows():
                meter_no = str(row.get('meter_number', ''))
                if meter_no:
                    billing_lookup[meter_no] = row
        
        # Create Power OFF lookup
        power_off_lookup = {}
        if not power_off_data.empty and 'meterno' in power_off_data.columns:
            for _, row in power_off_data.iterrows():
                meter_no = str(row.get('meterno', ''))
                if meter_no:
                    power_off_count = row.get('power_off_exceptions_count', 0)
                    if power_off_count is not None and pd.notna(power_off_count):
                        power_off_lookup[meter_no] = float(power_off_count) if isinstance(power_off_count, (int, float)) else 0
        
        # Create result dataframe
        result_rows = []
        
        for _, satc_row in hes_satc.iterrows():
            meter_number = str(satc_row.get('meterno', ''))
            if not meter_number:
                continue
            
            # Check exclusion type
            exclusion_type, match_record = self.check_exclusion_type(
                meter_number, pdc_data, sm_to_sm_data, non_opp_data, power_off_data, from_date
            )
            
            # Format location_code and consumer_id as text to preserve leading zeros
            location_code = satc_row.get('location_code', '')
            consumer_id = satc_row.get('consumer_id', '')
            
            # Convert to string and preserve formatting
            if location_code is not None and location_code != '':
                location_code = str(location_code)
                # If it's numeric, pad with zeros to 4 digits
                if location_code.isdigit():
                    location_code = location_code.zfill(4)
            else:
                location_code = ''
            
            if consumer_id is not None and consumer_id != '':
                consumer_id = str(consumer_id)
            else:
                consumer_id = ''
            
            # Get Zone, Circle, Division, Subdivision from ami_loc based on BU
            zone, circle, division, subdivision = self.get_ami_loc_data(location_code)
            
            # Get IP value
            ip_value = float(satc_row.get('supported_ip', 0)) if satc_row.get('supported_ip', 0) else 0
            
            # Get Power OFF count if available
            power_off_count = power_off_lookup.get(meter_number, 0)
            
            # Calculate BLP_Tg and DLP_Tg based on exclusion type
            if exclusion_type in ['PD', 'SM_SM', 'Non_Opp']:
                # Excluded meters should have no target
                blp_tg = 0
                dlp_tg = 0
            
            elif exclusion_type == 'Power_OFF' and power_off_count > 0:
                # Power OFF meters have reduced targets
                blp_tg = (ip_value * nod) - power_off_count
                blp_tg = max(blp_tg, 0)
            
                if ip_value > 0:
                    dlp_tg = nod - (power_off_count / ip_value)
                else:
                    dlp_tg = nod
            
                dlp_tg = max(dlp_tg, 0)
            
            else:
                # Normal eligible meters
                blp_tg = ip_value * nod
                dlp_tg = nod

            
            # Base row data with Zone, Circle, Division, Subdivision added between BU and KNO
            row_data = {
                'BU': location_code,
                'Zone': zone,
                'Circle': circle,
                'Division': division,
                'Subdivision': subdivision,
                'KNO': consumer_id,
                'meter_number': meter_number,
                'HES': hes_name,
                'SAT_No': satc_row.get('sat_pkg', ''),
                'UID': f"{satc_row.get('sat_pkg', '')}_{location_code}",
                'Exclusion_Type': exclusion_type,
                'Power_off_count': power_off_count if exclusion_type == 'Power_OFF' else 0,
                'IP': ip_value,
                'NOD': nod,
                'Exception': self.get_exception_text(exclusion_type, match_record),
                'Exception_Date': self.get_exception_date(exclusion_type, match_record),
                'Main_Ex_Flag': '1' if exclusion_type != '-' and exclusion_type != 'Power_OFF' else '0',
                'BLP_Tg': blp_tg,  # Use the calculated value
                'BLP_H8_SLA': 0,
                'BLP_H12_SLA': 0,
                'BLP_H24_SLA': 0,
                'DLP_Tg': dlp_tg,  # Use the calculated value
                'DLP_H24_SLA': 0,
                'Bill_Tg': 0,
                'Bill_h72_SLA': 0,
                'Bill_h168_SLA': 0
            }
            
            # Now get the actual SLA values for non-excluded records (including Power_OFF)
            if exclusion_type == '-' or exclusion_type == 'Power_OFF':
                # Get BLP values
                blp_row = blp_lookup.get(meter_number, {})
                row_data['BLP_H8_SLA'] = blp_row.get('h8', 0) if 'h8' in blp_row else 0
                row_data['BLP_H12_SLA'] = blp_row.get('h12', 0) if 'h12' in blp_row else 0
                row_data['BLP_H24_SLA'] = blp_row.get('h24', 0) if 'h24' in blp_row else 0
                
                # Get DLP values
                dlp_row = dlp_lookup.get(meter_number, {})
                row_data['DLP_H24_SLA'] = dlp_row.get('h24', 0) if 'h24' in dlp_row else 0
                
                # Get Billing values
                billing_row = billing_lookup.get(meter_number, {})
                row_data['Bill_Tg'] = 1
                row_data['Bill_h72_SLA'] = billing_row.get('72hrs', 0) if '72hrs' in billing_row else 0
                row_data['Bill_h168_SLA'] = billing_row.get('168hrs', 0) if '168hrs' in billing_row else 0
            
            result_rows.append(row_data)
        
        result_df = pd.DataFrame(result_rows)
        return result_df
    
    
    def create_sdo_wise_summary(self, hes_wise_data, month_year, sat_numbers):
        """Create SDO_Wise_Summary sheet with ALL UID combinations"""
        self.log_message("  Creating SDO_Wise_Summary sheet...")
        
        # Step 1: Get all SAT numbers from SATC data
        if sat_numbers is None or not sat_numbers:
            sat_numbers = []
            # Extract from all HES data
            for hes_name, hes_df in hes_wise_data.items():
                if 'SAT_No' in hes_df.columns:
                    sat_list = hes_df['SAT_No'].unique().tolist()
                    for sat in sat_list:
                        if sat and str(sat) not in sat_numbers:
                            sat_numbers.append(str(sat))
            sat_numbers = sorted(sat_numbers)
        
        # Step 2: Get all BU codes from ami_loc
        all_bu_codes = []
        if self.ami_loc_df is not None:
            # Get all unique sitecodes from ami_loc
            if 'sitecode' in self.ami_loc_df.columns:
                all_bu_codes = self.ami_loc_df['sitecode'].astype(str).str.zfill(4).unique().tolist()
                all_bu_codes = sorted(all_bu_codes)
        
        self.log_message(f"  Found {len(sat_numbers)} SAT numbers: {sat_numbers}")
        self.log_message(f"  Found {len(all_bu_codes)} BU codes from ami_loc")
        
        # Step 3: Create all possible UIDs
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
        
        # Step 4: Collect actual data from all HES sheets
        actual_data = {}
        
        for hes_name, hes_df in hes_wise_data.items():
            required_cols = ['UID', 'BLP_Tg', 'BLP_H8_SLA', 'BLP_H12_SLA', 'BLP_H24_SLA', 
                           'DLP_Tg', 'DLP_H24_SLA', 'Bill_Tg', 'Bill_h72_SLA', 'Bill_h168_SLA']
            
            missing_cols = [col for col in required_cols if col not in hes_df.columns]
            if missing_cols:
                self.log_message(f"  Warning: Missing columns in {hes_name}: {missing_cols}")
                continue
            
            # Group by UID and aggregate
            for _, row in hes_df.iterrows():
                uid = row.get('UID', '')
                if uid and uid != '':
                    if uid not in actual_data:
                        actual_data[uid] = {
                            'BLP_Tg': 0,
                            'BLP_H8_SLA': 0,
                            'BLP_H12_SLA': 0,
                            'BLP_H24_SLA': 0,
                            'DLP_Tg': 0,
                            'DLP_H24_SLA': 0,
                            'Bill_Tg': 0,
                            'Bill_h72_SLA': 0,
                            'Bill_h168_SLA': 0,
                            'meter_count': 0
                        }
                    
                    # Sum up values
                    actual_data[uid]['BLP_Tg'] += row.get('BLP_Tg', 0)
                    actual_data[uid]['BLP_H8_SLA'] += row.get('BLP_H8_SLA', 0)
                    actual_data[uid]['BLP_H12_SLA'] += row.get('BLP_H12_SLA', 0)
                    actual_data[uid]['BLP_H24_SLA'] += row.get('BLP_H24_SLA', 0)
                    actual_data[uid]['DLP_Tg'] += row.get('DLP_Tg', 0)
                    actual_data[uid]['DLP_H24_SLA'] += row.get('DLP_H24_SLA', 0)
                    actual_data[uid]['Bill_Tg'] += row.get('Bill_Tg', 0)
                    actual_data[uid]['Bill_h72_SLA'] += row.get('Bill_h72_SLA', 0)
                    actual_data[uid]['Bill_h168_SLA'] += row.get('Bill_h168_SLA', 0)
                    actual_data[uid]['meter_count'] += 1
        
        self.log_message(f"  Found data for {len(actual_data)} UIDs")
        
        # Step 5: Create SDO_Wise_Summary with ALL UIDs
        sdo_rows = []
        
        for combo in all_possible_uids:
            uid = combo['uid']
            sat_number = combo['sat_number']
            bu_code = combo['bu_code']
            
            # Get zone, circle, division, subdivision from ami_loc
            zone, circle, division, subdivision = self.get_ami_loc_data(bu_code)
            
            # Get data for this UID (or zeros if not found)
            if uid in actual_data:
                data = actual_data[uid]
                blp_tg = data['BLP_Tg']
                blp_h8 = data['BLP_H8_SLA']
                blp_h12 = data['BLP_H12_SLA']
                blp_h24 = data['BLP_H24_SLA']
                dlp_tg = data['DLP_Tg']
                dlp_h24 = data['DLP_H24_SLA']
                bill_tg = data['Bill_Tg']
                bill_h72 = data['Bill_h72_SLA']
                bill_h168 = data['Bill_h168_SLA']
                meter_no_count = data['meter_count']
            else:
                # All zeros for UIDs with no data
                blp_tg = 0
                blp_h8 = 0
                blp_h12 = 0
                blp_h24 = 0
                dlp_tg = 0
                dlp_h24 = 0
                bill_tg = 0
                bill_h72 = 0
                bill_h168 = 0
                meter_no_count = 0
            
            # Extract month and year from month_year
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
            # Determine from_date based on Sync Data selection
            if self.fetch_var.get() == "YES":
                # If Sync Data is YES, use the date from entry field
                if self.from_date_selected is None:
                    self.from_date_selected = self.from_date_entry.get()
                self.log_message(f"Using from_date from entry: {self.from_date_selected}")
            else:
                # If Sync Data is NO, extract from SATC filename
                if self.month_year_from_filename:
                    from_date_from_filename = self.get_from_date_from_mmyy(self.month_year_from_filename)
                    if from_date_from_filename:
                        self.from_date_selected = from_date_from_filename
                        self.log_message(f"NOD (MMYY {self.month_year_from_filename})")
                    else:
                        # Fallback to entry if extraction fails
                        self.from_date_selected = self.from_date_entry.get()
                        self.log_message(f"Could not extract date from filename, using from_date from entry: {self.from_date_selected}")
                else:
                    # If no month_year from filename, use entry
                    self.from_date_selected = self.from_date_entry.get()
                    self.log_message(f"Using from_date from entry: {self.from_date_selected}")
            
            # Load all raw data files
            self.log_message("\nLoading raw data files...")
            
            self.log_message("Loading SATC Data...")
            satc_data = pd.read_excel(self.satc_file_path, sheet_name=None)
            
            self.log_message("Loading BLP Data...")
            blp_data = pd.read_excel(self.blp_file_path, sheet_name=None)
            
            self.log_message("Loading DLP Data...")
            dlp_data = pd.read_excel(self.dlp_file_path, sheet_name=None)
            
            self.log_message("Loading Billing Data...")
            billing_data = pd.read_excel(self.billing_file_path, sheet_name=None)
            
            # Always load Power OFF data (regardless of exclusion setting)
            self.log_message("Loading Power OFF Data...")
            power_off_data = pd.read_excel(self.power_off_file_path)
            
            # Only load other exclusion-related files if Exclusion is YES
            if self.exclusion_enabled:
                self.log_message("Loading PDRC Data...")
                pdrc_data = pd.read_excel(self.pdrc_file_path)
                
                self.log_message("Loading SM to SM Data...")
                sm_to_sm_data = pd.read_excel(self.sm_to_sm_file_path)
                
                self.log_message("Loading Non Opp Data...")
                non_opp_data = pd.read_excel(self.non_opp_file_path)
                
                self.log_message("Loading RC_DC Data...")
                rc_dc_data = pd.read_excel(self.rc_dc_file_path)
            else:
                # Create empty dataframes for exclusion data
                pdrc_data = pd.DataFrame()
                sm_to_sm_data = pd.DataFrame()
                non_opp_data = pd.DataFrame()
                rc_dc_data = pd.DataFrame()
                self.log_message("⚠ Exclusion is DISABLED. Exclusion-related files (PDC, SM to SM, Non Opp, RC_DC) will be skipped.")
                self.log_message("✓ Power OFF data will be used.")
            
            # Calculate NOD from filename or from_date
            if self.month_year_from_filename:
                nod = self.get_nod_from_month_year(self.month_year_from_filename)
                self.log_message(f"\nNumber of Days (NOD) : {nod}")
            else:
                nod = self.calculate_nod(self.from_date_selected)
                self.log_message(f"\nNumber of Days (NOD) : {nod}")
            
            # Prepare consolidated SATC data (all sheets combined for lookup)
            all_satc_data = pd.DataFrame()
            for sheet_name, sheet_data in satc_data.items():
                # Format location_code and consumer_id as text
                sheet_copy = sheet_data.copy()
                if 'location_code' in sheet_copy.columns:
                    sheet_copy['location_code'] = sheet_copy['location_code'].astype(str).str.zfill(4)
                if 'consumer_id' in sheet_copy.columns:
                    sheet_copy['consumer_id'] = sheet_copy['consumer_id'].astype(str)
                all_satc_data = pd.concat([all_satc_data, sheet_copy], ignore_index=True)
            
            # Prepare BLP, DLP, Billing data by sheet (HES-wise)
            blp_by_hes = {}
            for sheet_name, sheet_data in blp_data.items():
                blp_by_hes[sheet_name] = sheet_data
            
            dlp_by_hes = {}
            for sheet_name, sheet_data in dlp_data.items():
                dlp_by_hes[sheet_name] = sheet_data
            
            billing_by_hes = {}
            for sheet_name, sheet_data in billing_data.items():
                billing_by_hes[sheet_name] = sheet_data
            
            # Create HES-wise raw data sheets
            hes_list = ['BCITS', 'Cyan', 'Crystal']
            hes_wise_data = {}
            
            for hes_name in hes_list:
                self.log_message(f"\nCreating SLA_{hes_name}_Raw sheet...")
                
                # Get HES-specific data
                blp_hes_data = blp_by_hes.get(hes_name, pd.DataFrame())
                dlp_hes_data = dlp_by_hes.get(hes_name, pd.DataFrame())
                billing_hes_data = billing_by_hes.get(hes_name, pd.DataFrame())
                
                hes_data = self.create_hes_wise_raw_data(
                    hes_name, all_satc_data, blp_hes_data, dlp_hes_data, billing_hes_data,
                    pdrc_data, sm_to_sm_data, non_opp_data, power_off_data, self.from_date_selected, nod
                )
                
                if not hes_data.empty:
                    # Format BU and KNO columns
                    if 'BU' in hes_data.columns:
                        hes_data['BU'] = hes_data['BU'].astype(str).str.zfill(4)
                    if 'KNO' in hes_data.columns:
                        hes_data['KNO'] = hes_data['KNO'].astype(str)
                    hes_wise_data[f'SLA_{hes_name}_Raw'] = hes_data
                    self.log_message(f"  Created SLA_{hes_name}_Raw with {len(hes_data)} records")
            
            # Get sat_numbers from all sheets
            sat_numbers = []
            for hes_name, hes_df in hes_wise_data.items():
                if 'SAT_No' in hes_df.columns:
                    sat_list = hes_df['SAT_No'].unique().tolist()
                    for sat in sat_list:
                        if sat and sat not in sat_numbers:
                            sat_numbers.append(str(sat))
            sat_numbers = sorted(sat_numbers)
            self.log_message(f"\n  SAT Numbers found: {sat_numbers}")
            
            # Get month-year for file naming
            month_year = self.get_month_year_from_date(self.from_date_selected)
            current_date = datetime.datetime.now().strftime("%d%m%y")
            
            # Determine filename based on exclusion setting
            if self.exclusion_enabled:
                output_filename = os.path.join(self.script_dir, f"SLA_HES_Wise_Final_{month_year}_{current_date}.xlsx")
            else:
                output_filename = os.path.join(self.script_dir, f"SLA_Cumm_Worksheet_Final_{month_year}_{current_date}.xlsx")
            
            # Save to Excel
            self.log_message(f"\nSaving final Excel...!!")
            self.log_message(f"Filename: {os.path.basename(output_filename)}")
            
            with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
                # Write Hes_Wise_Summary with updated columns
                hes_summary = []
                
                for hes_name, hes_df in hes_wise_data.items():
                    hes_display_name = hes_name.replace('SLA_', '').replace('_Raw', '')
                    
                    # Count by exclusion type
                    total_meters = len(hes_df)
                    
                    if self.exclusion_enabled:
                        # Full exclusion tracking
                        pdc_count = len(hes_df[hes_df['Exclusion_Type'] == 'PD'])
                        non_opp_count = len(hes_df[hes_df['Exclusion_Type'] == 'Non_Opp'])
                        sm_sm_count = len(hes_df[hes_df['Exclusion_Type'] == 'SM_SM'])
                        power_off_count = len(hes_df[hes_df['Exclusion_Type'] == 'Power_OFF'])
                        eligible_meters = total_meters - (pdc_count + non_opp_count + sm_sm_count)
                    else:
                        # When exclusion is disabled, only Power_OFF is tracked
                        pdc_count = 0
                        non_opp_count = 0
                        sm_sm_count = 0
                        power_off_count = len(hes_df[hes_df['Exclusion_Type'] == 'Power_OFF'])
                        eligible_meters = total_meters
                    
                    # Filter active records (Main_Ex_Flag == '0')
                    if self.exclusion_enabled:
                        active_df = hes_df[hes_df['Main_Ex_Flag'] == '0']
                    else:
                        # When exclusion is disabled, all records are active
                        active_df = hes_df
                    
                    # Calculate subtotals for active records
                    blp_tg = active_df['BLP_Tg'].sum() if not active_df.empty else 0
                    blp_8 = active_df['BLP_H8_SLA'].sum() if not active_df.empty else 0
                    blp_12 = active_df['BLP_H12_SLA'].sum() if not active_df.empty else 0
                    blp_24 = active_df['BLP_H24_SLA'].sum() if not active_df.empty else 0
                    dlp_tg = active_df['DLP_Tg'].sum() if not active_df.empty else 0
                    dlp_24 = active_df['DLP_H24_SLA'].sum() if not active_df.empty else 0
                    bill_tg = active_df['Bill_Tg'].sum() if not active_df.empty else 0
                    bill_3 = active_df['Bill_h72_SLA'].sum() if not active_df.empty else 0
                    bill_7 = active_df['Bill_h168_SLA'].sum() if not active_df.empty else 0
                    
                    # Calculate percentages
                    pct_blp_8 = (blp_8 / blp_tg * 100) if blp_tg > 0 else 0
                    pct_blp_12 = (blp_12 / blp_tg * 100) if blp_tg > 0 else 0
                    pct_blp_24 = (blp_24 / blp_tg * 100) if blp_tg > 0 else 0
                    pct_dlp_24 = (dlp_24 / dlp_tg * 100) if dlp_tg > 0 else 0
                    pct_bill_3 = (bill_3 / bill_tg * 100) if bill_tg > 0 else 0
                    pct_bill_7 = (bill_7 / bill_tg * 100) if bill_tg > 0 else 0
                    
                    # Build summary row based on exclusion setting
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
                        # Simplified summary with only Power_OFF exclusion
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
                
                # Add All HES row (cumulative)
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
                    
                    # Calculate cumulative percentages
                    all_hes_row['%BLP_8'] = (all_hes_row['BLP_8'] / all_hes_row['BLP_Tg'] * 100) if all_hes_row['BLP_Tg'] > 0 else 0
                    all_hes_row['%BLP_12'] = (all_hes_row['BLP_12'] / all_hes_row['BLP_Tg'] * 100) if all_hes_row['BLP_Tg'] > 0 else 0
                    all_hes_row['%BLP_24'] = (all_hes_row['BLP_24'] / all_hes_row['BLP_Tg'] * 100) if all_hes_row['BLP_Tg'] > 0 else 0
                    all_hes_row['%DLP_24'] = (all_hes_row['DLP_24'] / all_hes_row['DLP_Tg'] * 100) if all_hes_row['DLP_Tg'] > 0 else 0
                    all_hes_row['%Bill_3'] = (all_hes_row['Bill_3'] / all_hes_row['Bill_Tg'] * 100) if all_hes_row['Bill_Tg'] > 0 else 0
                    all_hes_row['%Bill_7'] = (all_hes_row['Bill_7'] / all_hes_row['Bill_Tg'] * 100) if all_hes_row['Bill_Tg'] > 0 else 0
                    
                    hes_summary.append(all_hes_row)
                
                hes_summary_df = pd.DataFrame(hes_summary)
                
                # Format percentage columns
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
                
                # Write RC_DC sheet only if Exclusion is YES
                if self.exclusion_enabled and not rc_dc_data.empty:
                    rc_dc_data.to_excel(writer, sheet_name='RC_DC', index=False)
                    self.log_message("  Created RC_DC sheet")
                
                # Write SLA_BCITS_Raw, SLA_Cyan_Raw, SLA_Crystal_Raw
                for sheet_name, sheet_data in hes_wise_data.items():
                    # Ensure BU and KNO are formatted as text
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
                    # Write PDC_RAW
                    if not pdrc_data.empty:
                        pdrc_data.to_excel(writer, sheet_name='PDC_RAW', index=False)
                        self.log_message("  Created PDC_RAW sheet")
                    
                    # Write Non_Opp_Raw
                    if not non_opp_data.empty:
                        non_opp_data.to_excel(writer, sheet_name='Non_Opp_Raw', index=False)
                        self.log_message("  Created Non_Opp_Raw sheet")
                    
                    # Write SM_to_SM_RAW
                    if not sm_to_sm_data.empty:
                        sm_to_sm_data.to_excel(writer, sheet_name='SM_to_SM_RAW', index=False)
                        self.log_message("  Created SM_to_SM_RAW sheet")
            
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
        
        # Check if all required files are selected based on exclusion setting
        if self.exclusion_enabled:
            # All files required
            if not all([self.satc_file_path, self.blp_file_path, self.dlp_file_path, self.billing_file_path, 
                       self.pdrc_file_path, self.sm_to_sm_file_path, self.non_opp_file_path, self.power_off_file_path, 
                       self.rc_dc_file_path]):
                self.log_message("Error: Please select all required files for the final summary.")
                messagebox.showerror("Error", "Please select all required files for the final summary.\n\nRequired files:\n- SATC RAW\n- BLP RAW\n- DLP RAW\n- BILLING RAW\n- PDRC RAW\n- SM to SM RAW\n- Non Opp RAW\n- Power OFF RAW\n- RC_DC RAW")
                return
        else:
            # Only 5 files required when exclusion is disabled (including Power OFF)
            if not all([self.satc_file_path, self.blp_file_path, self.dlp_file_path, self.billing_file_path, self.power_off_file_path]):
                self.log_message("Error: Please select all required files for the final summary.")
                messagebox.showerror("Error", "Please select all required files for the final summary.\n\nRequired files:\n- SATC RAW\n- BLP RAW\n- DLP RAW\n- BILLING RAW\n- Power OFF RAW")
                return
        
        # Run final process in separate thread
        process_thread = threading.Thread(target=self.create_final_excel)
        process_thread.daemon = True
        process_thread.start()