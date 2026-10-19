# build_exe.py
import PyInstaller.__main__
import os
import sys

def create_exe():
    # Define the main script and output directory
    script = "app.py"
    app_name = "DataProcessingTool"
    
    # Define PyInstaller arguments
    args = [
    script,
    '--name', app_name,
    '--onefile',
    '--windowed',
    '--icon', 'icon.ico',
    '--add-data', 'config.example.ini;.',
    '--add-data', 'ami_loc.xlsx;.',
    '--distpath', './dist',
    '--workpath', './build',
    '--specpath', './',
    '--clean',
    '--noconfirm',
]
    
    # Add hidden imports for all dependencies
    hidden_imports = [
        'psycopg2',
        'pandas',
        'openpyxl',
        'xlrd',
        'tkinter',
        'tkcalendar',
        'pickle',
        'datetime',
        'threading',
        'socket',
        'time'
    ]
    
    for imp in hidden_imports:
        args.extend(['--hidden-import', imp])
    
    # Run PyInstaller
    PyInstaller.__main__.run(args)

if __name__ == "__main__":
    create_exe()