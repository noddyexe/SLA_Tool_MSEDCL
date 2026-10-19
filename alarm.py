import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import psycopg2
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


class AlarmReportApp:
    def __init__(self, root, go_back_callback, save_directory, is_exe):
        self.root = root
        self.go_back_callback = go_back_callback
        self.save_directory = save_directory
        self.is_exe = is_exe
        self.root.title("Alarm Report")
        self.root.geometry("1200x700")
        self.root.configure(bg='white')
        
        # Configuration for retries and batch processing
        self.MAX_RETRIES = 500000
        self.RETRY_INTERVAL_MS = 5000
        self.METER_BATCH_SIZE = 1000 
        self.DATA_BATCH_SIZE = 1000
        
        # Database connection cache
        self.db_connections = {}
        
        # Progress tracking file
        self.progress_file = os.path.join(save_directory, 'alarm_progress.json')
        
        # Style configuration
        self.style = ttk.Style()
        self.style.configure('TFrame', background='MediumAquamarine')
        self.style.configure('TLabel', background='white', font=('Arial', 10))
        
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
        self.vpn_connected = False
        self.processing = False
        self.stop_processing = False
        
        # # Load previous data if exists
        # self.load_previous_data()
    
    def log_message(self, message):
       """Add message to log with timestamp and print to terminal"""
       timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
       formatted_message = f"[{timestamp}] {message}\n"
       
       # Add to GUI log
       self.log_text.insert(tk.END, formatted_message)
       self.log_text.see(tk.END)
       self.root.update()
       
       # Print to terminal if available
       try:
           print(formatted_message, end="")
           if sys.stdout is not None:
               sys.stdout.flush()
       except:
           pass  # Silently ignore if stdout is not available
    
    def create_widgets(self):
        """Create all GUI widgets for Alarm Report"""
        # Go Back Button (top right)
        go_back_button = tk.Button(self.button_frame, text="← Go Back", 
                                 command=self.go_back_callback,
                                 bg='#f44336', fg='white',
                                 font=('Arial', 10, 'bold'),
                                 bd=0, padx=15, pady=5,
                                 activebackground='#d32f2f',
                                 activeforeground='white')
        go_back_button.grid(row=0, column=2, sticky=tk.E, pady=5)
        
        # Bind hover effects
        go_back_button.bind("<Enter>", lambda e: go_back_button.config(bg='#d32f2f'))
        go_back_button.bind("<Leave>", lambda e: go_back_button.config(bg='#f44336'))
        
        # Title
        title_label = ttk.Label(self.button_frame, text="Alarm Report", 
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
        
        # Bind hover effects
        vpn_button.bind("<Enter>", lambda e: vpn_button.config(bg='#42A5F5'))
        vpn_button.bind("<Leave>", lambda e: vpn_button.config(bg='#2196F3'))
        
        # Select File Button
        excel_button = tk.Button(self.button_frame, text="Select File", 
                               command=self.load_excel, 
                               bg='#2196F3', fg='white',
                               font=('Arial', 10, 'bold'),
                               bd=0, padx=20, pady=10,
                               activebackground='#42A5F5',
                               activeforeground='white')
        excel_button.grid(row=3, column=0, columnspan=2, pady=5, sticky=tk.EW)
        
        # Bind hover effects
        excel_button.bind("<Enter>", lambda e: excel_button.config(bg='#42A5F5'))
        excel_button.bind("<Leave>", lambda e: excel_button.config(bg='#2196F3'))
        
        # Alarm Type Dropdown
        self.alarm_type_var = tk.StringVar()
        alarm_type_label = ttk.Label(self.button_frame, text="Alarm Type:")
        alarm_type_label.grid(row=4, column=0, sticky=tk.W, pady=5)
        
        self.alarm_type_combo = ttk.Combobox(self.button_frame, textvariable=self.alarm_type_var, 
                                           values=["Last Gasp- Occurrence", "First Breath- Restoration", 
                                                  "Over Voltage", "Voltage unbalance"],
                                           state="readonly")
        self.alarm_type_combo.grid(row=4, column=1, columnspan=2, sticky=tk.EW, pady=5)
        self.alarm_type_combo.set("Last Gasp- Occurrence")
        
        # Server Time Selection
        server_time_label = ttk.Label(self.button_frame, text="Server Time:")
        server_time_label.grid(row=5, column=0, sticky=tk.W, pady=5)
        
        # Use DateEntry for calendar selection
        try:
            from tkcalendar import DateEntry
            self.server_time_entry = DateEntry(self.button_frame, width=12, background='darkblue',
                                             foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
            self.server_time_entry.grid(row=5, column=1, columnspan=2, sticky=tk.EW, pady=5)
        except ImportError:
            # Fallback to regular Entry if tkcalendar is not available
            self.server_time_entry = ttk.Entry(self.button_frame, width=12)
            self.server_time_entry.grid(row=5, column=1, columnspan=2, sticky=tk.EW, pady=5)
            self.server_time_entry.insert(0, datetime.datetime.now().strftime("%Y-%m-%d"))
        
        # Run Process Button
        self.run_button = tk.Button(self.button_frame, text="Run Process", 
                             command=self.start_process_thread, 
                             bg='#4CAF50', fg='white',
                             font=('Arial', 10, 'bold'),
                             bd=0, padx=20, pady=10,
                             activebackground='#66BB6A',
                             activeforeground='white')
        self.run_button.grid(row=6, column=0, columnspan=3, pady=20, sticky=tk.EW)
        
        # Stop Process Button (initially hidden)
        self.stop_button = tk.Button(self.button_frame, text="Stop Process", 
                             command=self.stop_process, 
                             bg='#f44336', fg='white',
                             font=('Arial', 10, 'bold'),
                             bd=0, padx=20, pady=10,
                             activebackground='#d32f2f',
                             activeforeground='white')
        
        # Resume from Progress Button
        self.resume_button = tk.Button(self.button_frame, text="Resume from Last Progress", 
                                     command=self.resume_process,
                                     bg='#FF9800', fg='white',
                                     font=('Arial', 10, 'bold'),
                                     bd=0, padx=20, pady=10,
                                     activebackground='#F57C00',
                                     activeforeground='white')
        
        # Check if progress file exists and show resume button
        if os.path.exists(self.progress_file):
            self.resume_button.grid(row=7, column=0, columnspan=3, pady=5, sticky=tk.EW)
        
        # Bind hover effects
        self.run_button.bind("<Enter>", lambda e: self.run_button.config(bg='#66BB6A'))
        self.run_button.bind("<Leave>", lambda e: self.run_button.config(bg='#4CAF50'))
        self.resume_button.bind("<Enter>", lambda e: self.resume_button.config(bg='#F57C00'))
        self.resume_button.bind("<Leave>", lambda e: self.resume_button.config(bg='#FF9800'))
    
    def save_progress(self, processed_batches, completed_batches, aggregated_results, total_batches):
        """Save current progress to file"""
        # Convert set to list for JSON serialization
        completed_batches_list = list(completed_batches) if completed_batches else []
        
        progress_data = {
            'processed_batches': processed_batches,
            'completed_batch_indices': completed_batches_list,  # Now it's a list, not a set
            'total_batches': total_batches,
            'alarm_type': self.alarm_type_var.get(),
            'server_time': self.server_time_entry.get(),
            'timestamp': datetime.datetime.now().isoformat(),
            'meter_numbers': self.get_meter_numbers_for_query()  # Save original meter list
        }
        
        # Save results so far
        if aggregated_results:
            temp_results_file = os.path.join(self.save_directory, 'temp_results.pkl')
            with open(temp_results_file, 'wb') as f:
                pickle.dump(aggregated_results, f)
            progress_data['temp_results_file'] = temp_results_file
        
        with open(self.progress_file, 'w') as f:
            json.dump(progress_data, f)
        
        self.log_message(f"Progress saved: {processed_batches}/{total_batches} batches completed")
    
    def load_progress(self):
        """Load saved progress"""
        if not os.path.exists(self.progress_file):
            return None, []
        
        try:
            with open(self.progress_file, 'r') as f:
                progress_data = json.load(f)
            
            # Convert completed_batch_indices back to set for internal use
            if 'completed_batch_indices' in progress_data:
                progress_data['completed_batch_indices'] = set(progress_data['completed_batch_indices'])
            
            # Load temporary results if they exist
            aggregated_results = []
            if 'temp_results_file' in progress_data and os.path.exists(progress_data['temp_results_file']):
                with open(progress_data['temp_results_file'], 'rb') as f:
                    aggregated_results = pickle.load(f)
            
            return progress_data, aggregated_results
        except Exception as e:
            self.log_message(f"Error loading progress: {str(e)}")
            return None, []
    
    def clear_progress(self):
        """Clear saved progress"""
        if os.path.exists(self.progress_file):
            os.remove(self.progress_file)
        
        temp_results_file = os.path.join(self.save_directory, 'temp_results.pkl')
        if os.path.exists(temp_results_file):
            os.remove(temp_results_file)
    
    def resume_process(self):
        """Resume process from saved progress"""
        progress_data, aggregated_results = self.load_progress()
        if not progress_data:
            self.log_message("No saved progress found.")
            return
        
        # Set the saved parameters
        self.alarm_type_var.set(progress_data['alarm_type'])
        self.server_time_entry.delete(0, tk.END)
        self.server_time_entry.insert(0, progress_data['server_time'])
        
        # Start the process with resume flag
        self.start_process_thread(resume=True, progress_data=progress_data, 
                                 existing_results=aggregated_results)
    
    def start_process_thread(self, resume=False, progress_data=None, existing_results=None):
        """Start the process in a separate thread to keep GUI responsive"""
        if not self.processing:
            self.processing = True
            self.stop_processing = False
            self.run_button.grid_remove()
            self.resume_button.grid_remove()
            self.stop_button.grid(row=6, column=0, columnspan=3, pady=20, sticky=tk.EW)
            
            # Store resume data
            self.resume_data = {
                'resume': resume,
                'progress_data': progress_data,
                'existing_results': existing_results
            }
            
            # Start the process in a separate thread
            process_thread = threading.Thread(target=self.run_process)
            process_thread.daemon = True
            process_thread.start()
    
    def stop_process(self):
        """Stop the current process"""
        self.stop_processing = True
        self.log_message("Stopping process...")


    
    def check_vpn(self):
        """Check VPN connection by testing database server connectivity"""
        self.log_message("Checking VPN connection...")
        
        host = get_config()['MDM']['ip']
        port = int(get_config()['MDM']['port'])
        
        try:
            # Create a socket connection
            s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            s.settimeout(5)  # 5-second timeout
            
            # Attempt connection
            s.connect((host, port))
            s.close()
            
            self.log_message("VPN connected successfully!")
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
        
        host = get_config()['MDM']['ip']
        port = int(get_config()['MDM']['port'])
 
        
        retries = 0
        while retries < max_retries and not self.stop_processing:
            try:
                # Create a socket connection
                s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                s.settimeout(5)  # 5-second timeout
                
                # Attempt connection
                s.connect((host, port))
                s.close()
                
                self.log_message("VPN connected successfully!")
                self.vpn_connected = True
                return True
                
            except Exception as e:
                self.log_message(f"VPN connection attempt {retries + 1}/{max_retries} failed: {str(e)}")
                self.vpn_connected = False
                retries += 1
                
                # Wait before retrying
                for i in range(retry_interval):
                    if self.stop_processing:
                        return False
                    time.sleep(1)
        
        return False
    
    def load_excel(self):
       """Load Excel file with meter_number column and save to local data_selve.dat"""
       file_path = filedialog.askopenfilename(
           title="Select Excel File",
           filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
       )
       
       # Check if user cancelled the dialog
       if not file_path:
           self.log_message("No file selected. Operation cancelled.")
           return
       
       try:
           self.log_message(f"Loading Excel file: {file_path}")
           self.excel_data = pd.read_excel(file_path)
           
           # Check if meter_number column exists
           if 'meter_number' not in self.excel_data.columns:
               self.log_message("Error: 'meter_number' column not found in Excel file")
               messagebox.showerror("Error", "'meter_number' column not found in Excel file")
               return
           
           # Save to local file
           data_file = os.path.join(self.save_directory, 'data_selve.dat')
           with open(data_file, 'wb') as f:
               pickle.dump(self.excel_data, f)
           
           self.log_message(f"Total meter numbers loaded: {len(self.excel_data)}")
           
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
    
    def get_meter_numbers_for_query(self):
        """Get meter numbers in list format for batch processing"""
        if self.excel_data is None:
            self.log_message("No data available. Please load Excel file first.")
            return []
        
        # Filter valid meter numbers
        meter_numbers = []
        for num in self.excel_data['meter_number'].unique():
            if pd.notna(num) and str(num).strip():
                meter_numbers.append(str(num).strip())
        
        return meter_numbers


    def connect_db(self):
        if 'alarm_db' in self.db_connections:
            conn, cur = self.db_connections['alarm_db']
            try:
                cur.execute("SELECT 1")
                return cur, conn
            except:
                try:
                    cur.close(); conn.close()
                except:
                    pass
                del self.db_connections['alarm_db']
    
        cfg = get_config()['MDM']          # alarm uses MDM host
        DBIP   = cfg['ip']
        DBPort = cfg['port']
        DBUID  = cfg['uid']
        DBPwd  = cfg['pwd']
        BaseDB = cfg['database']
    
        try:
            conn = psycopg2.connect(
                host=DBIP, port=DBPort, database=BaseDB,
                user=DBUID, password=DBPwd
            )
            cur = conn.cursor()
            self.db_connections['alarm_db'] = (conn, cur)
            return cur, conn
        except (Exception, psycopg2.DatabaseError) as error:
            self.log_message(f'Alarm DB Connection Failed. Error: {error}')
            return None, None


        
    
    def close_db_connection(self):
        """Close database connection"""
        if 'alarm_db' in self.db_connections:
            try:
                conn, cur = self.db_connections['alarm_db']
                cur.close()
                conn.close()
                del self.db_connections['alarm_db']
                # self.log_message("Database connection closed")
            except Exception as e:
                self.log_message(f"Error closing database connection: {str(e)}")
    
    def format_meter_list(self, meter_list):
        """Format meter list for SQL IN clause"""
        if not meter_list:
            return "('')"
        return "('" + "','".join(meter_list) + "')"
    
    def execute_single_query(self, cur, query, batch_num=None, aggregated_count=0):
       """Execute single query with batch fetching"""
       results = []
       start_time = time.time()
       
       try:
           # Use server-side cursor for batch fetching
           cur.execute(query)
           
           # Get column names
           columns = [desc[0] for desc in cur.description]
           
           # Fetch in batches
           fetch_batch_count = 0
           while True:
               batch_start = time.time()
               rows = cur.fetchmany(self.DATA_BATCH_SIZE)
               if not rows:
                   break
               
               fetch_batch_count += 1
               
               # Convert to list of dictionaries
               for row in rows:
                   results.append(dict(zip(columns, row)))
               
               # Calculate cumulative total
               current_cumulative = aggregated_count + len(results)
               
               # Log batch progress with cumulative total
               batch_time = time.time() - batch_start
               if batch_num is not None:
                   self.log_message(f"Batch {batch_num} - Fetched {len(rows)}, Total: {current_cumulative}")
               else:
                   self.log_message(f"Fetch batch {fetch_batch_count}: Fetched {len(rows)}, Total: {current_cumulative}")
               
               if self.stop_processing:
                   break
           
           total_time = time.time() - start_time
           
           return results
       except (Exception, psycopg2.DatabaseError) as error:
           self.log_message(f'Query execution failed. Error: {error}')
           raise
       
    
    def build_alarm_query(self, meter_list):
        """Build the alarm query based on selections"""
        alarm_type_full = self.alarm_type_var.get()
        
        # Extract the base alarm type (everything before the hyphen if it exists)
        if '-' in alarm_type_full:
            alarm_type_base = alarm_type_full.split('-')[0].strip()
        else:
            alarm_type_base = alarm_type_full.strip()
        
        try:
            server_time = self.server_time_entry.get()
        except:
            server_time = datetime.datetime.now().strftime("%Y-%m-%d")
        
        meter_in_clause = self.format_meter_list(meter_list)
        
        # Optimized query - using LIKE with the base alarm type
        query = f"""
        SELECT meter_number, alarm_active, COUNT(*) as alarmCount 
        FROM meter_data.meter_alarm 
        WHERE alarm_active LIKE '{alarm_type_base}%'
        AND server_time >= '{server_time}'
        AND meter_number IN {meter_in_clause}
        GROUP BY meter_number, alarm_active
        ORDER BY meter_number DESC
        """
        return query
    
    
    
    def execute_alarm_query_with_retry(self, meter_numbers, resume_from_batch=0, 
                                      completed_batches=None, existing_results=None):
        """Execute alarm SQL query with retry logic for VPN disconnections and resume capability"""
        if not meter_numbers:
            return pd.DataFrame()
        
        # Initialize tracking variables
        if completed_batches is None:
            completed_batches = set()
        if existing_results is None:
            existing_results = []
        
        self.log_message(f"Processing {len(meter_numbers)} meter numbers in batches")
        total_batches = ((len(meter_numbers) - 1) // self.METER_BATCH_SIZE) + 1
        
        if resume_from_batch > 0:
            self.log_message(f"Resuming from batch {resume_from_batch + 1}/{total_batches}")
            self.log_message(f"Already completed batches: {sorted(completed_batches)}")
        
        aggregated_results = existing_results.copy()
        retry_count = 0
        start_process_time = time.time()
        
        # Process meter numbers in batches
        for i in range(resume_from_batch * self.METER_BATCH_SIZE, len(meter_numbers), self.METER_BATCH_SIZE):
            if self.stop_processing:
                self.save_progress(i // self.METER_BATCH_SIZE, completed_batches, 
                                 aggregated_results, total_batches)
                break
            
            batch_num = (i // self.METER_BATCH_SIZE) + 1
            
            # Skip already completed batches
            if batch_num - 1 in completed_batches:
                self.log_message(f"Skipping already completed batch {batch_num}")
                continue
            
            batch_successful = False
            batch_retry_count = 0
            
            while not batch_successful and batch_retry_count < self.MAX_RETRIES and not self.stop_processing:
                try:
                    # Get database connection
                    cur, conn = self.connect_db()
                    if cur is None:
                        # Wait for VPN connection
                        self.log_message(f"Batch {batch_num}: Database connection failed. Waiting for VPN...")
                        if not self.wait_for_vpn(max_retries=10000, retry_interval=5):
                            break
                        cur, conn = self.connect_db()
                        if cur is None:
                            batch_retry_count += 1
                            continue
                    
                    batch_start_time = time.time()
                    end = min(len(meter_numbers), i + self.METER_BATCH_SIZE)
                    batch = meter_numbers[i:end]
                    
                    self.log_message(f"---Processing batch {batch_num}/{total_batches}---")
                    
                    # Build and execute query for this batch
                    query = self.build_alarm_query(batch)
                    if query:
                        batch_results = self.execute_single_query(cur, query, batch_num, len(aggregated_results))
                        aggregated_results.extend(batch_results)
                        
                        # Mark batch as completed
                        completed_batches.add(batch_num - 1)
                        batch_successful = True
                        
                        # Save progress after each successful batch
                        self.save_progress(batch_num, completed_batches, aggregated_results, total_batches)
                    
                    # Log progress
                    batch_time = time.time() - batch_start_time
                    
                    # Calculate estimated time remaining
                    if batch_num > 1:
                        avg_time_per_batch = (time.time() - start_process_time) / len(completed_batches)
                        remaining_batches = total_batches - len(completed_batches)
                        estimated_remaining = avg_time_per_batch * remaining_batches
                        self.log_message(f"Estimated time remaining: {estimated_remaining/60:.1f} minutes")
                    
                except Exception as e:
                    self.log_message(f"\n✗ Batch {batch_num} failed (Attempt {batch_retry_count + 1}/{self.MAX_RETRIES}): {str(e)}")
                    batch_retry_count += 1
                    
                    # Close connection if exists
                    self.close_db_connection()
                    
                    if batch_retry_count >= self.MAX_RETRIES:
                        self.log_message(f"✗ Max retries reached for batch {batch_num}")
                        break
                    
                    self.log_message(f"Waiting {self.RETRY_INTERVAL_MS/1000}s before retrying batch {batch_num}...")
                    for i in range(int(self.RETRY_INTERVAL_MS/1000)):
                        if self.stop_processing:
                            break
                        time.sleep(1)
            
            if not batch_successful and batch_retry_count >= self.MAX_RETRIES:
                self.log_message(f"Failed to process batch {batch_num} after {self.MAX_RETRIES} attempts")
                break
        
        # Close connection when all batches are processed
        self.close_db_connection()
        
        if aggregated_results:
            # Convert results to DataFrame
            df = pd.DataFrame(aggregated_results)
            total_process_time = time.time() - start_process_time
            self.log_message(f"✓ Total records fetched: {len(df)}")
            return df
        else:
            self.log_message("\n✓ Query completed successfully but no data found")
            return pd.DataFrame()  # Success but no data
    
    def save_alarm_results(self, df):
        """Save alarm results to Excel file and locally"""
        if df is None or df.empty:
            self.log_message("No alarm data to save.")
            return
        
        # Save to data_selve.dat first
        try:
            data_file = os.path.join(self.save_directory, 'data_selve.dat')
            with open(data_file, 'wb') as f:
                pickle.dump(df, f)
            # self.log_message("Alarm data saved locally to data_selve.dat")
        except Exception as e:
            self.log_message(f"Error saving alarm data locally: {str(e)}")
        
        # Get current date in yyyymmdd format
        current_date = datetime.datetime.now().strftime("%Y%m%d")
        
        # Get month and year from the server_time in MMYY format
        try:
            server_time = self.server_time_entry.get()
            month_year = datetime.datetime.strptime(server_time, "%Y-%m-%d").strftime("%m%y")
        except:
            month_year = datetime.datetime.now().strftime("%m%y")
        
        # Get alarm type for filename
        alarm_type = self.alarm_type_var.get()
        
        # Map alarm type to filename part
        if alarm_type == "Last Gasp- Occurrence":
            alarm_type_part = "Last Gasp"
        elif alarm_type == "First Breath- Restoration":
            alarm_type_part = "First Breath"
        elif alarm_type == "Over Voltage":
            alarm_type_part = "Over Voltage"
        elif alarm_type == "Voltage unbalance":
            alarm_type_part = "Voltage unbalance"
        else:
            alarm_type_part = "alarm"
        
        # Save to Excel file with alarm type and month_year included
        filename = os.path.join(self.save_directory, f"Final_{alarm_type_part}_alarm_report_{month_year}_{current_date}.xlsx")
        
        try:
            # Check if data exceeds Excel row limit
            if len(df) > 1048576:
                # Save as CSV
                csv_filename = os.path.join(self.save_directory, f"Final_{alarm_type_part}_alarm_report_{month_year}_{current_date}.csv")
                df.to_csv(csv_filename, index=False)
                self.log_message(f"Data too large for Excel ({len(df)} rows). Alarm results saved to CSV: {csv_filename}")
            else:
                # Save as Excel
                df.to_excel(filename, index=False)
                self.log_message(f"Alarm results saved to Excel: {filename}")
            
            # Clear progress after successful save
            self.clear_progress()
                
        except Exception as e:
            self.log_message(f"Error saving alarm results: {str(e)}")
    
    def run_process(self):
        """Main process to run based on selections"""
        global start_process_time
        start_process_time = time.time()
        
        # Check if we're resuming
        is_resume = hasattr(self, 'resume_data') and self.resume_data.get('resume', False)
        
        if not self.vpn_connected and not is_resume:
            self.log_message("VPN not connected. Please check VPN first.")
            messagebox.showwarning("VPN Not Connected", "Please check VPN connection before fetching data.")
            self.cleanup_after_process()
            return
        
        if self.excel_data is None and not is_resume:
            self.log_message("No Excel data loaded. Please select a file first.")
            self.cleanup_after_process()
            return
        
        self.log_message("Starting alarm report process...")
        
        if is_resume:
            # Resume from saved progress
            progress_data = self.resume_data['progress_data']
            existing_results = self.resume_data['existing_results']
            
            meter_numbers = progress_data['meter_numbers']
            completed_batches = set(progress_data['completed_batch_indices'])
            resume_from_batch = progress_data['processed_batches']
            
            self.log_message(f"Resuming process with {len(completed_batches)} completed batches")
            
            # Process meter numbers with resume
            final_result = self.execute_alarm_query_with_retry(
                meter_numbers, 
                resume_from_batch=resume_from_batch,
                completed_batches=completed_batches,
                existing_results=existing_results
            )
        else:
            # New process
            meter_numbers = self.get_meter_numbers_for_query()
            if not meter_numbers:
                self.log_message("No meter numbers found in the Excel file.")
                self.cleanup_after_process()
                return
            
            # Process meter numbers with retry logic
            final_result = self.execute_alarm_query_with_retry(meter_numbers)
        
        if not self.stop_processing:
            if not final_result.empty:
                self.save_alarm_results(final_result)
                total_time = time.time() - start_process_time
                self.log_message(f"\n✓ Alarm report completed successfully!")
                # self.log_message(f"✓ Total records: {len(final_result)}")
                # self.log_message(f"✓ Total time: {total_time/60:.1f} minutes")
            else:
                self.log_message("\n✗ No alarm data found for any meter numbers")
                self.clear_progress()  # Clear progress if no data found
        
        self.cleanup_after_process()
    
    def cleanup_after_process(self):
        """Clean up after process completes or stops"""
        self.processing = False
        self.run_button.grid(row=6, column=0, columnspan=3, pady=20, sticky=tk.EW)
        self.stop_button.grid_remove()
        
        # Show resume button if progress exists
        if os.path.exists(self.progress_file):
            self.resume_button.grid(row=7, column=0, columnspan=3, pady=5, sticky=tk.EW)
        else:
            self.resume_button.grid_remove()
        
        # Clear resume data
        if hasattr(self, 'resume_data'):
            delattr(self, 'resume_data')