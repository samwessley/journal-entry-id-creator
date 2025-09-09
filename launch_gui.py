#!/usr/bin/env python3
"""
Simple launcher for Journal Entry ID Creator GUI
Double-click this file to launch the application
"""

try:
    from journal_entry_gui import main
    main()
except ImportError as e:
    import tkinter as tk
    from tkinter import messagebox
    
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    
    messagebox.showerror(
        "Missing Dependencies", 
        f"Required modules are missing:\n\n{e}\n\n"
        "Please install the required dependencies:\n"
        "pip install pandas openpyxl"
    )
    
except Exception as e:
    import tkinter as tk
    from tkinter import messagebox
    
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    
    messagebox.showerror(
        "Application Error", 
        f"An error occurred while starting the application:\n\n{e}"
    )
