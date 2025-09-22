#!/usr/bin/env python3
"""
GUI for Journal Entry ID Creator
Provides a user-friendly interface to select input files, output locations, and run the journal entry creator.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
from pathlib import Path
from journal_entry_creator import JournalEntryCreator
import sys

class JournalEntryGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Journal Entry ID Creator")
        self.root.geometry("600x450")
        self.root.resizable(True, True)
        
        # Variables
        self.input_file_var = tk.StringVar()
        self.output_file_var = tk.StringVar()
        self.running = False
        
        self.setup_ui()
        self.center_window()
        
    def center_window(self):
        """Center the window on the screen"""
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (self.root.winfo_width() // 2)
        y = (self.root.winfo_screenheight() // 2) - (self.root.winfo_height() // 2)
        self.root.geometry(f"+{x}+{y}")
        
    def setup_ui(self):
        """Setup the user interface"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Journal Entry ID Creator", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Description
        desc_text = ("This tool reads journal lines from an Excel file and creates balanced journal entries\n"
                    "where each entry has equal debits and credits, grouped by date and other fields.")
        desc_label = ttk.Label(main_frame, text=desc_text, foreground="gray")
        desc_label.grid(row=1, column=0, columnspan=3, pady=(0, 20))
        
        # Input file selection
        ttk.Label(main_frame, text="Input Excel File:", font=("Arial", 10, "bold")).grid(
            row=2, column=0, sticky=tk.W, pady=(0, 5))
        
        input_frame = ttk.Frame(main_frame)
        input_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        input_frame.columnconfigure(0, weight=1)
        
        self.input_entry = ttk.Entry(input_frame, textvariable=self.input_file_var, 
                                    font=("Arial", 9))
        self.input_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        
        ttk.Button(input_frame, text="Browse...", 
                  command=self.browse_input_file).grid(row=0, column=1)
        
        # Output file selection
        ttk.Label(main_frame, text="Output Location:", font=("Arial", 10, "bold")).grid(
            row=4, column=0, sticky=tk.W, pady=(0, 5))
        
        output_frame = ttk.Frame(main_frame)
        output_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        output_frame.columnconfigure(0, weight=1)
        
        self.output_entry = ttk.Entry(output_frame, textvariable=self.output_file_var,
                                     font=("Arial", 9))
        self.output_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        
        ttk.Button(output_frame, text="Browse...", 
                  command=self.browse_output_file).grid(row=0, column=1)
        
        # Run button
        self.run_button = ttk.Button(main_frame, text="Create Journal Entries", 
                                    command=self.run_journal_creator,
                                    style="Accent.TButton")
        self.run_button.grid(row=6, column=0, columnspan=3, pady=20)
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Status label
        self.status_label = ttk.Label(main_frame, text="Ready to process journal entries", 
                                     foreground="green")
        self.status_label.grid(row=8, column=0, columnspan=3)
        
        # Results text area
        self.results_frame = ttk.LabelFrame(main_frame, text="Processing Results", padding="10")
        self.results_frame.grid(row=9, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), 
                               pady=(20, 0))
        self.results_frame.columnconfigure(0, weight=1)
        self.results_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(9, weight=1)
        
        # Text widget with scrollbar
        text_frame = ttk.Frame(self.results_frame)
        text_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        text_frame.columnconfigure(0, weight=1)
        text_frame.rowconfigure(0, weight=1)
        
        self.results_text = tk.Text(text_frame, height=8, wrap=tk.WORD, font=("Consolas", 9))
        scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=self.results_text.yview)
        self.results_text.configure(yscrollcommand=scrollbar.set)
        
        self.results_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Initially hide results frame
        self.results_frame.grid_remove()
        
    def browse_input_file(self):
        """Browse for input Excel file"""
        filename = filedialog.askopenfilename(
            title="Select Input Excel File",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("All files", "*.*")
            ],
            initialdir=os.getcwd()
        )
        if filename:
            self.input_file_var.set(filename)
            # Auto-generate output filename
            if not self.output_file_var.get():
                input_path = Path(filename)
                output_path = input_path.parent / f"{input_path.stem}_with_journal_ids{input_path.suffix}"
                self.output_file_var.set(str(output_path))
    
    def browse_output_file(self):
        """Browse for output Excel file location"""
        input_file = self.input_file_var.get()
        initial_name = ""
        initial_dir = os.getcwd()
        
        if input_file:
            input_path = Path(input_file)
            initial_name = f"{input_path.stem}_with_journal_ids{input_path.suffix}"
            initial_dir = str(input_path.parent)
        
        filename = filedialog.asksaveasfilename(
            title="Save Output Excel File As",
            filetypes=[
                ("Excel files", "*.xlsx"),
                ("All files", "*.*")
            ],
            initialdir=initial_dir,
            initialfile=initial_name,
            defaultextension=".xlsx"
        )
        if filename:
            self.output_file_var.set(filename)
    
    def validate_inputs(self):
        """Validate user inputs"""
        input_file = self.input_file_var.get().strip()
        output_file = self.output_file_var.get().strip()
        
        if not input_file:
            messagebox.showerror("Error", "Please select an input Excel file.")
            return False
            
        if not os.path.exists(input_file):
            messagebox.showerror("Error", f"Input file does not exist:\n{input_file}")
            return False
            
        if not output_file:
            messagebox.showerror("Error", "Please specify an output file location.")
            return False
            
        # Check if output directory exists
        output_dir = os.path.dirname(output_file)
        if output_dir and not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir)
            except Exception as e:
                messagebox.showerror("Error", f"Cannot create output directory:\n{e}")
                return False
                
        return True
    
    def update_status(self, message, color="black"):
        """Update status label"""
        self.status_label.config(text=message, foreground=color)
        self.root.update_idletasks()
    
    def append_results(self, text):
        """Append text to results area"""
        self.results_text.insert(tk.END, text + "\n")
        self.results_text.see(tk.END)
        self.root.update_idletasks()
    
    def clear_results(self):
        """Clear results text area"""
        self.results_text.delete(1.0, tk.END)
    
    def run_journal_creator(self):
        """Run the journal entry creator in a separate thread"""
        if self.running:
            return
            
        if not self.validate_inputs():
            return
            
        # Start processing in separate thread
        thread = threading.Thread(target=self._process_journal_entries)
        thread.daemon = True
        thread.start()
    
    def _process_journal_entries(self):
        """Process journal entries (runs in separate thread)"""
        try:
            self.running = True
            
            # Update UI
            self.root.after(0, lambda: self.run_button.config(state='disabled'))
            self.root.after(0, lambda: self.progress.start())
            self.root.after(0, lambda: self.update_status("Processing journal entries...", "blue"))
            self.root.after(0, lambda: self.results_frame.grid())
            self.root.after(0, self.clear_results)
            
            input_file = self.input_file_var.get().strip()
            output_file = self.output_file_var.get().strip()
            max_fields = 5  # Use default value
            
            # Create journal entry creator
            creator = JournalEntryCreator()
            
            # Redirect output to capture results
            import io
            import contextlib
            
            output_buffer = io.StringIO()
            
            with contextlib.redirect_stdout(output_buffer):
                # Load data
                self.root.after(0, lambda: self.append_results("Loading data from Excel file..."))
                if not creator.load_data(input_file):
                    raise Exception("Failed to load data from input file")
                
                # Validate balances before attempting to create entries
                self.root.after(0, lambda: self.append_results("Validating balances (overall, by date, by month if needed)..."))
                creator.validate_balances()
                
                # Create journal entries
                self.root.after(0, lambda: self.append_results("Creating journal entries..."))
                if not creator.create_journal_entries(max_fields):
                    raise Exception("Failed to create journal entries")
                
                # If there are unassigned lines, prompt to auto-balance with plug
                if len(creator.unassigned_lines) > 0:
                    self.root.after(0, lambda: self.append_results(
                        f"Found {len(creator.unassigned_lines)} unassigned lines after grouping."))
                    def ask_balance():
                        return messagebox.askyesno(
                            "Unbalanced Entries Detected",
                            "There are unassigned/unbalanced journal lines by date.\n\n"
                            "Would you like to automatically add plug lines using account 'Audit Sight Clearing' "
                            "to balance each affected date?"
                        )
                    user_choice = self.root.after(0, ask_balance)
                    # Run synchronously to capture user choice
                    choice = messagebox.askyesno(
                        "Unbalanced Entries Detected",
                        "There are unassigned/unbalanced journal lines by date.\n\n"
                        "Would you like to automatically add plug lines using account 'Audit Sight Clearing' "
                        "to balance each affected date?"
                    )
                    if choice:
                        balanced_dates = creator.balance_unassigned_with_plug()
                        self.root.after(0, lambda: self.append_results(
                            f"Added plug lines for {balanced_dates} posted date(s)."))
                    else:
                        raise Exception("User declined to auto-balance unassigned lines.")
                
                # Generate output
                self.root.after(0, lambda: self.append_results("Generating output file..."))
                if not creator.generate_output(input_file, output_file):
                    raise Exception("Failed to generate output file")
            
            # Get captured output
            captured_output = output_buffer.getvalue()
            
            # Update UI with results
            self.root.after(0, lambda: self.append_results("\n" + "="*50))
            self.root.after(0, lambda: self.append_results("PROCESSING COMPLETE!"))
            self.root.after(0, lambda: self.append_results("="*50))
            
            # Show summary from captured output
            lines = captured_output.split('\n')
            summary_started = False
            for line in lines:
                if "Summary:" in line:
                    summary_started = True
                if summary_started and line.strip():
                    self.root.after(0, lambda l=line: self.append_results(l))
                    if "Total lines unassigned:" in line:
                        break
            
            self.root.after(0, lambda: self.append_results(f"\nOutput file saved to:\n{output_file}"))
            
            # Success
            self.root.after(0, lambda: self.update_status("Processing completed successfully!", "green"))
            self.root.after(0, lambda: messagebox.showinfo("Success", 
                f"Journal entries created successfully!\n\nOutput saved to:\n{output_file}"))
            
        except Exception as e:
            error_msg = str(e)
            self.root.after(0, lambda: self.append_results(f"\nERROR: {error_msg}"))
            self.root.after(0, lambda: self.update_status(f"Error: {error_msg}", "red"))
            self.root.after(0, lambda: messagebox.showerror("Error", f"Processing failed:\n\n{error_msg}"))
            
        finally:
            self.running = False
            self.root.after(0, lambda: self.progress.stop())
            self.root.after(0, lambda: self.run_button.config(state='normal'))

def main():
    """Main function to run the GUI"""
    root = tk.Tk()
    
    # Set up styles
    style = ttk.Style()
    
    # Try to use a modern theme if available
    available_themes = style.theme_names()
    if 'clam' in available_themes:
        style.theme_use('clam')
    elif 'alt' in available_themes:
        style.theme_use('alt')
    
    # Create custom style for the run button
    style.configure("Accent.TButton", font=("Arial", 10, "bold"))
    
    # Create and run the application
    app = JournalEntryGUI(root)
    
    # Handle window closing
    def on_closing():
        if app.running:
            if messagebox.askokcancel("Quit", "Processing is still running. Do you want to quit?"):
                root.destroy()
        else:
            root.destroy()
    
    root.protocol("WM_DELETE_WINDOW", on_closing)
    
    try:
        root.mainloop()
    except KeyboardInterrupt:
        root.destroy()

if __name__ == "__main__":
    main()
