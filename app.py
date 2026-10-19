import tkinter as tk
from tkinter import ttk, messagebox
import os
import sys
import datetime
import pickle
# import the backend modules
from sat import DBDataFetcher
from alarm import AlarmReportApp
from SLA_SATC.sla import SLAExtractionApp
from config_loader import get_config


#pyinstaller --onefile --name=SLA_ALARM_SAT_TOOL  app.py


class LoginApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Data Processing Application - Login")
        self.root.geometry("500x600")
        self.root.configure(bg='#2196F3')
        self.root.resizable(False, False)
        
        cfg = get_config()['LOGIN']
        self.valid_credentials = {
            "mcl": cfg['mcl'],
            "admin": cfg['admin'],
            "user": cfg['user'],
            "f": cfg['f'],
        }
        self.expiration_date = cfg['expiration_date']
        
        self.create_login_ui()

        
    
    def check_expiration(self):
        """Check if current date is past expiration date"""
        current_date = datetime.datetime.now()
        
        # Parse the expiration date string to a datetime object
        day = int(self.expiration_date[:2])
        month = int(self.expiration_date[2:4])
        year = int(self.expiration_date[4:])
        expiration_date_obj = datetime.datetime(year, month, day)
        
        return current_date > expiration_date_obj
    
    def create_login_ui(self):
        """Create login interface"""
        # Main container
        main_frame = tk.Frame(self.root, bg="#2196F3")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=30)
        
        # Logo/Header
        logo_frame = tk.Frame(main_frame, bg="#2196F3")
        logo_frame.pack(pady=(20, 30))
        
        # Application Title
        title_label = tk.Label(logo_frame, 
                              text="Data Processing\nApplication",
                              font=('Arial', 20, 'bold'),
                              fg="white",
                              bg="#2196F3",
                              justify=tk.CENTER)
        title_label.pack()
        
        # Subtitle
        subtitle_label = tk.Label(logo_frame,
                                 text="SAT Extraction & ALARM Report",
                                 font=('Arial', 12),
                                 fg="#E3F2FD",
                                 bg="#2196F3")
        subtitle_label.pack(pady=(10, 0))
        
        # Login Box
        login_box = tk.Frame(main_frame, bg="white", relief=tk.RAISED, bd=1)
        login_box.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
        
        # Login Title
        login_title = tk.Label(login_box,
                              text="USER LOGIN",
                              font=('Arial', 16, 'bold'),
                              bg="white",
                              fg="#2196F3")
        login_title.pack(pady=(25, 20))
        
        # Input Fields Frame
        input_frame = tk.Frame(login_box, bg="white", padx=30, pady=10)
        input_frame.pack(fill=tk.BOTH, expand=True)
        
        # Username
        username_frame = tk.Frame(input_frame, bg="white")
        username_frame.pack(fill=tk.X, pady=(0, 15))
        
        tk.Label(username_frame,
                text="Username:",
                font=('Arial', 11, 'bold'),
                bg="white",
                fg="#555").pack(anchor=tk.W)
        
        self.username_entry = tk.Entry(username_frame,
                                      font=('Arial', 12),
                                      relief=tk.GROOVE,
                                      bd=2,
                                      width=25)
        self.username_entry.pack(fill=tk.X, pady=(5, 0))
        self.username_entry.bind("<Return>", lambda e: self.password_entry.focus())
        
        # Password
        password_frame = tk.Frame(input_frame, bg="white")
        password_frame.pack(fill=tk.X, pady=(0, 25))
        
        tk.Label(password_frame,
                text="Password:",
                font=('Arial', 11, 'bold'),
                bg="white",
                fg="#555").pack(anchor=tk.W)
        
        self.password_entry = tk.Entry(password_frame,
                                      font=('Arial', 12),
                                      relief=tk.GROOVE,
                                      bd=2,
                                      width=25,
                                      show="•")
        self.password_entry.pack(fill=tk.X, pady=(5, 0))
        self.password_entry.bind("<Return>", lambda e: self.authenticate())
        
        # Login Button Frame
        button_frame = tk.Frame(login_box, bg="white", pady=10)
        button_frame.pack(fill=tk.X)
        
        self.login_button = tk.Button(button_frame,
                                     text="LOGIN",
                                     command=self.authenticate,
                                     font=('Arial', 12, 'bold'),
                                     bg="#2196F3",
                                     fg="white",
                                     relief=tk.GROOVE,
                                     bd=0,
                                     padx=40,
                                     pady=10,
                                     cursor="hand2")
        self.login_button.pack()
        
        # Set cursor style
        self.login_button.bind("<Enter>", lambda e: self.login_button.config(bg="#1976D2"))
        self.login_button.bind("<Leave>", lambda e: self.login_button.config(bg="#2196F3"))
        
        # Show/Hide Password Checkbox
        self.show_pass_var = tk.BooleanVar()
        show_pass_check = tk.Checkbutton(login_box,
                                        text="Show Password",
                                        variable=self.show_pass_var,
                                        command=self.toggle_password,
                                        font=('Arial', 9),
                                        bg="white",
                                        fg="#666",
                                        selectcolor="white")
        show_pass_check.pack(pady=(0, 20))
        
        # Footer
        footer_frame = tk.Frame(main_frame, bg="#2196F3")
        footer_frame.pack(fill=tk.X, pady=(10, 0))
        
        # Version info
        version_label = tk.Label(footer_frame,
                                text="Data Processing Application v1.0",
                                font=('Arial', 9),
                                fg="#E3F2FD",
                                bg="#2196F3")
        version_label.pack(side=tk.LEFT)
        
        # Date info
        date_label = tk.Label(footer_frame,
                             text=f"Valid until: {self.format_expiration_date()}",
                             font=('Arial', 9),
                             fg="#E3F2FD",
                             bg="#2196F3")
        date_label.pack(side=tk.RIGHT)
        
        # Copyright
        copyright_label = tk.Label(main_frame,
                                  text="© 2025 Data Processing System",
                                  font=('Arial', 9),
                                  fg="#BBDEFB",
                                  bg="#2196F3")
        copyright_label.pack(side=tk.BOTTOM, pady=(10, 0))
        
        # Set focus to username entry
        self.username_entry.focus_set()
    
    def format_expiration_date(self):
        """Format expiration date for display"""
        day = self.expiration_date[:2]
        month = self.expiration_date[2:4]
        year = self.expiration_date[4:]
        return f"{day}/{month}/{year}"
    
    def toggle_password(self):
        """Toggle password visibility"""
        if self.show_pass_var.get():
            self.password_entry.config(show="")
        else:
            self.password_entry.config(show="•")
    
    def authenticate(self):
        """Authenticate user credentials"""
        # First check if application has expired
        if self.check_expiration():
            messagebox.showerror("Access Denied", 
                               "User validity is over!!\n\nPlease contact the developer to enable the user...")
            return
        
        # Get raw values without stripping
        raw_username = self.username_entry.get()
        password = self.password_entry.get()
        
        # Trim whitespace for all except space username
        if raw_username != " ":
            username = raw_username.strip()
        else:
            username = raw_username
        
        if not username or not password:
            messagebox.showerror("Login Failed", "Please enter both username and password")
            return
        
        # Check credentials
        if username in self.valid_credentials and self.valid_credentials[username] == password:
            self.login_success()
        else:
            messagebox.showerror("Login Failed", "Invalid username or password")
            self.password_entry.delete(0, tk.END)
            self.password_entry.focus_set()
    
    def login_success(self):
        """Handle successful login"""
        # Disable login button
        self.login_button.config(state=tk.DISABLED, bg="#cccccc")
        
        # Show success message
        success_label = tk.Label(self.root,
                                text="Login Successful!",
                                font=('Arial', 12, 'bold'),
                                fg="green",
                                bg="#2196F3")
        success_label.place(relx=0.5, rely=0.9, anchor=tk.CENTER)
        
        # Open main application after delay
        self.root.after(1000, self.open_main_app)
    
    def open_main_app(self):
        """Open the main application"""
        self.root.destroy()
        
        # Create new root window for main application
        main_root = tk.Tk()
        app = MainApplication(main_root)
        main_root.mainloop()


class MainApplication:
    def __init__(self, root):
        self.root = root
        self.root.title("Data Processing Application")
        self.root.geometry("1200x700")
        self.root.configure(bg='white')
        
        # Determine if running as exe
        self.is_exe = getattr(sys, 'frozen', False)
        
        # Set the default save directory
        if self.is_exe:
            # When running as EXE, save files in the same directory as the EXE
            if getattr(sys, 'frozen', False):
                # Running as PyInstaller bundle
                application_path = sys._MEIPASS
                self.save_directory = os.path.dirname(sys.executable)
            else:
                # Normal Python script
                self.save_directory = os.getcwd()
        else:
            self.save_directory = os.getcwd()
        
        # Create save directory if it doesn't exist
        os.makedirs(self.save_directory, exist_ok=True)
        
        # Show front page initially
        self.show_front_page()
    
    def show_front_page(self):
        """Show the front page with two main options"""
        # Clear the window
        for widget in self.root.winfo_children():
            widget.destroy()
        
        # Create main frame
        main_frame = tk.Frame(self.root, bg='white')
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Header frame
        header_frame = tk.Frame(main_frame, bg='#2196F3', height=80)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        # Title in header
        title_label = tk.Label(header_frame, text="Data Processing Application", 
                              font=('Arial', 24, 'bold'), bg='#2196F3', fg='white')
        title_label.pack(expand=True)
        
        # Logout button in header
        logout_button = tk.Button(header_frame, text="Logout", 
                                 command=self.logout,
                                 font=('Arial', 10, 'bold'),
                                 bg='#ffffff',
                                 fg='#2196F3',
                                 bd=0,
                                 padx=15,
                                 pady=5)
        logout_button.place(relx=0.95, rely=0.5, anchor=tk.E)
        
        # Set cursor to hand when hovering over the logout button
        logout_button.bind("<Enter>", lambda e: logout_button.config(bg='#f0f0f0', cursor="hand2"))
        logout_button.bind("<Leave>", lambda e: logout_button.config(bg='#ffffff', cursor=""))
        
        # Content frame
        content_frame = tk.Frame(main_frame, bg='white')
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # Buttons frame
        buttons_frame = tk.Frame(content_frame, bg='white')
        buttons_frame.pack(expand=True)
        
        # SAT Extraction Button
        sat_button = tk.Button(buttons_frame, text="SAT Extraction", 
                             command=self.open_sat_extraction,
                             bg='#4CAF50', fg='white',
                             font=('Arial', 16, 'bold'),
                             bd=0, padx=40, pady=20,
                             activebackground='#66BB6A',
                             activeforeground='white')
        sat_button.pack(pady=20)
        
        # Bind hover effects
        sat_button.bind("<Enter>", lambda e: sat_button.config(bg='#66BB6A', cursor="hand2"))
        sat_button.bind("<Leave>", lambda e: sat_button.config(bg='#4CAF50', cursor=""))
        
        # Alarm Report Button
        alarm_button = tk.Button(buttons_frame, text="ALARM Report", 
                               command=self.open_alarm_report,
                               bg='#FF9800', fg='white',
                               font=('Arial', 16, 'bold'),
                               bd=0, padx=40, pady=20,
                               activebackground='#FFB74D',
                               activeforeground='white')
        alarm_button.pack(pady=20)
        
        # Bind hover effects
        alarm_button.bind("<Enter>", lambda e: alarm_button.config(bg='#FFB74D', cursor="hand2"))
        alarm_button.bind("<Leave>", lambda e: alarm_button.config(bg='#FF9800', cursor=""))
        
        # SLA Extraction Button
        sla_button = tk.Button(buttons_frame, text="SLA Extraction", 
                             command=self.open_sla_extraction,
                             bg='#9C27B0', fg='white',
                             font=('Arial', 16, 'bold'),
                             bd=0, padx=40, pady=20,
                             activebackground='#AB47BC',
                             activeforeground='white')
        sla_button.pack(pady=20)
        
        # Bind hover effects
        sla_button.bind("<Enter>", lambda e: sla_button.config(bg='#AB47BC', cursor="hand2"))
        sla_button.bind("<Leave>", lambda e: sla_button.config(bg='#9C27B0', cursor=""))
        
        
        
        # Footer
        footer_frame = tk.Frame(main_frame, bg='#2196F3', height=30)
        footer_frame.pack(fill=tk.X, side=tk.BOTTOM)
        footer_frame.pack_propagate(False)
        
        footer_label = tk.Label(footer_frame, text="© 2025 Data Processing System",
                              font=('Arial', 10), bg='#2196F3', fg='white')
        footer_label.pack(expand=True)
    
    def logout(self):
        """Logout and return to login screen"""
        self.root.destroy()
        
        # Create new login window
        login_root = tk.Tk()
        login_app = LoginApp(login_root)
        login_root.mainloop()
    
    def open_sat_extraction(self):
        """Open SAT Extraction module"""
        # Clear the window
        for widget in self.root.winfo_children():
            widget.destroy()
        
        # Create SAT Extraction instance
        self.sat_app = DBDataFetcher(self.root, self.show_front_page, self.save_directory, self.is_exe)
    
    def open_alarm_report(self):
        """Open Alarm Report module"""
        # Clear the window
        for widget in self.root.winfo_children():
            widget.destroy()
        
        # Create Alarm Report instance
        self.alarm_app = AlarmReportApp(self.root, self.show_front_page, self.save_directory, self.is_exe)  
        
    def open_sla_extraction(self):
        """Open SLA Extraction module"""
        # Clear the window
        for widget in self.root.winfo_children():
            widget.destroy()
        
        # Create SLA Extraction instance
        self.sla_app = SLAExtractionApp(self.root, self.show_front_page, self.save_directory, self.is_exe)    
    

if __name__ == "__main__":
    # Start with login screen
    login_root = tk.Tk()
    login_app = LoginApp(login_root)
    login_root.mainloop()