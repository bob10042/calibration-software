import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog
import subprocess
import sys
from datetime import datetime
import os
import pandas as pd

class VSCodeTheme:
    BG = "#1e1e1e"
    FG = "#d4d4d4"
    ACCENT_BLUE = "#0078d4"  # VS Code blue accent color
    SELECTED_BG = "#264f78"
    BTN_BG = "#2d2d2d"
    BTN_FG = "#cccccc"
    CONSOLE_BG = "#1e1e1e"
    CONSOLE_FG = "#d4d4d4"
    MENU_BG = "#333333"
    MENU_FG = "#cccccc"
    SIDEBAR_BG = "#252526"
    PANE_BG = "#1e1e1e"

class ScriptLauncher:
    def __init__(self, root):
        self.root = root
        self.root.title("AGX Python Script Launcher")
        
        # Configure main window
        self.root.geometry("1400x900")
        self.root.minsize(1200, 800)
        self.root.configure(bg=VSCodeTheme.BG)
        
        # Configure styles
        self.style = ttk.Style()
        self.configure_styles()
        
        # Create menu bar
        self.create_menu()
        
        # Create main container
        main_frame = ttk.Frame(root, style='Main.TFrame')
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=0, pady=0)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        
        # Create horizontal paned window for main split
        main_paned = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        main_paned.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(0, weight=1)
        
        # Create left sidebar (25% width)
        self.create_sidebar(main_paned)
        
        # Create right side container
        right_container = ttk.Frame(main_paned, style='Content.TFrame')
        
        # Create vertical paned window for right side
        right_paned = ttk.PanedWindow(right_container, orient=tk.VERTICAL)
        right_paned.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        right_container.columnconfigure(0, weight=1)
        right_container.rowconfigure(0, weight=1)
        
        # Create top panes container (75% height)
        top_container = ttk.Frame(right_paned, style='Content.TFrame')
        
        # Create horizontal paned window for top panes
        top_paned = ttk.PanedWindow(top_container, orient=tk.HORIZONTAL)
        top_paned.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        top_container.columnconfigure(0, weight=1)
        top_container.rowconfigure(0, weight=1)
        
        # Create calibration and results panes
        self.create_cal_pane(top_paned)
        self.create_res_pane(top_paned)
        
        # Create terminal (25% height)
        self.create_terminal(right_paned)
        
        # Add containers to paned windows
        main_paned.add(self.sidebar, weight=1)
        main_paned.add(right_container, weight=3)
        right_paned.add(top_container, weight=3)
        right_paned.add(self.terminal_frame, weight=1)
        
    def configure_styles(self):
        self.style.configure('Main.TFrame', background=VSCodeTheme.BG)
        self.style.configure('Content.TFrame', background=VSCodeTheme.BG)
        
        self.style.configure('Header.TFrame', 
                           background=VSCodeTheme.ACCENT_BLUE)
        
        self.style.configure('Treeview',
                           background=VSCodeTheme.SIDEBAR_BG,
                           foreground=VSCodeTheme.FG,
                           fieldbackground=VSCodeTheme.SIDEBAR_BG,
                           font=('Segoe UI', 9))
        self.style.map('Treeview',
                      background=[('selected', VSCodeTheme.SELECTED_BG)])
                      
    def create_menu(self):
        menubar = tk.Menu(self.root, bg=VSCodeTheme.MENU_BG, fg=VSCodeTheme.MENU_FG)
        self.root.config(menu=menubar)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0, bg=VSCodeTheme.MENU_BG, fg=VSCodeTheme.MENU_FG)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Open Calibration Sheet...", command=self.open_calibration)
        file_menu.add_command(label="Open Results Sheet...", command=self.open_results)
        
        # Terminal menu
        terminal_menu = tk.Menu(menubar, tearoff=0, bg=VSCodeTheme.MENU_BG, fg=VSCodeTheme.MENU_FG)
        menubar.add_cascade(label="Terminal", menu=terminal_menu)
        terminal_menu.add_command(label="New Terminal", command=lambda: self.launch_terminal("cmd"))
        terminal_menu.add_command(label="New PowerShell", command=lambda: self.launch_terminal("powershell"))
        
    def create_sidebar(self, parent):
        self.sidebar = ttk.Frame(parent, style='Content.TFrame')
        
        # Explorer header
        header = tk.Frame(self.sidebar, bg=VSCodeTheme.ACCENT_BLUE)
        header.grid(row=0, column=0, sticky='ew', columnspan=2)
        tk.Label(header, text="EXPLORER",
                bg=VSCodeTheme.ACCENT_BLUE,
                fg="white",
                font=('Segoe UI', 9, 'bold'),
                padx=10, pady=5).grid(row=0, column=0, sticky=tk.W)
        
        # Create treeview with scrollbars
        self.tree = ttk.Treeview(self.sidebar, style='Treeview', show='tree')
        self.tree.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Add vertical scrollbar
        scrollbar = ttk.Scrollbar(self.sidebar, orient="vertical", command=self.tree.yview)
        scrollbar.grid(row=1, column=1, sticky=(tk.N, tk.S))
        
        # Add horizontal scrollbar
        hsb = ttk.Scrollbar(self.sidebar, orient="horizontal", command=self.tree.xview)
        hsb.grid(row=2, column=0, sticky=(tk.W, tk.E))
        
        # Configure tree scrolling
        self.tree.configure(yscrollcommand=scrollbar.set, xscrollcommand=hsb.set)
        
        # Configure grid weights for full height
        self.sidebar.columnconfigure(0, weight=1)
        self.sidebar.rowconfigure(1, weight=1)
        
        # Populate tree
        self.populate_tree()
        
        # Bind double-click
        self.tree.bind('<Double-1>', self.on_tree_double_click)
        
    def create_cal_pane(self, parent):
        cal_frame = ttk.Frame(parent, style='Content.TFrame')
        
        # Header
        header = tk.Frame(cal_frame, bg=VSCodeTheme.ACCENT_BLUE)
        header.grid(row=0, column=0, sticky='ew')
        tk.Label(header, text="CALIBRATION SHEET",
                bg=VSCodeTheme.ACCENT_BLUE,
                fg="white",
                font=('Segoe UI', 9, 'bold'),
                padx=10, pady=5).grid(row=0, column=0, sticky=tk.W)
        
        # Text widget with scrollbars
        self.cal_text = tk.Text(cal_frame,
                              bg=VSCodeTheme.PANE_BG,
                              fg=VSCodeTheme.FG,
                              font=('Consolas', 9),
                              wrap=tk.NONE)
        self.cal_text.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Add vertical scrollbar
        cal_vsb = ttk.Scrollbar(cal_frame, orient="vertical", command=self.cal_text.yview)
        cal_vsb.grid(row=1, column=1, sticky=(tk.N, tk.S))
        
        # Add horizontal scrollbar
        cal_hsb = ttk.Scrollbar(cal_frame, orient="horizontal", command=self.cal_text.xview)
        cal_hsb.grid(row=2, column=0, sticky=(tk.W, tk.E))
        
        # Configure text widget scrolling
        self.cal_text.configure(yscrollcommand=cal_vsb.set, xscrollcommand=cal_hsb.set)
        
        cal_frame.columnconfigure(0, weight=1)
        cal_frame.rowconfigure(1, weight=1)
        
        parent.add(cal_frame, weight=1)
        
    def create_res_pane(self, parent):
        res_frame = ttk.Frame(parent, style='Content.TFrame')
        
        # Header
        header = tk.Frame(res_frame, bg=VSCodeTheme.ACCENT_BLUE)
        header.grid(row=0, column=0, sticky='ew')
        tk.Label(header, text="RESULTS SHEET",
                bg=VSCodeTheme.ACCENT_BLUE,
                fg="white",
                font=('Segoe UI', 9, 'bold'),
                padx=10, pady=5).grid(row=0, column=0, sticky=tk.W)
        
        # Text widget with scrollbars
        self.res_text = tk.Text(res_frame,
                              bg=VSCodeTheme.PANE_BG,
                              fg=VSCodeTheme.FG,
                              font=('Consolas', 9),
                              wrap=tk.NONE)
        self.res_text.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Add vertical scrollbar
        res_vsb = ttk.Scrollbar(res_frame, orient="vertical", command=self.res_text.yview)
        res_vsb.grid(row=1, column=1, sticky=(tk.N, tk.S))
        
        # Add horizontal scrollbar
        res_hsb = ttk.Scrollbar(res_frame, orient="horizontal", command=self.res_text.xview)
        res_hsb.grid(row=2, column=0, sticky=(tk.W, tk.E))
        
        # Configure text widget scrolling
        self.res_text.configure(yscrollcommand=res_vsb.set, xscrollcommand=res_hsb.set)
        
        res_frame.columnconfigure(0, weight=1)
        res_frame.rowconfigure(1, weight=1)
        
        parent.add(res_frame, weight=1)
        
    def create_terminal(self, parent):
        self.terminal_frame = ttk.Frame(parent, style='Content.TFrame')
        
        # Header
        header = tk.Frame(self.terminal_frame, bg=VSCodeTheme.ACCENT_BLUE)
        header.grid(row=0, column=0, sticky='ew')
        tk.Label(header, text="TERMINAL",
                bg=VSCodeTheme.ACCENT_BLUE,
                fg="white",
                font=('Segoe UI', 9, 'bold'),
                padx=10, pady=5).grid(row=0, column=0, sticky=tk.W)
        
        # Terminal output
        self.terminal = scrolledtext.ScrolledText(
            self.terminal_frame,
            bg=VSCodeTheme.CONSOLE_BG,
            fg=VSCodeTheme.CONSOLE_FG,
            insertbackground=VSCodeTheme.CONSOLE_FG,
            font=('Consolas', 9)
        )
        self.terminal.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.terminal_frame.columnconfigure(0, weight=1)
        self.terminal_frame.rowconfigure(1, weight=1)
        
    def populate_tree(self):
        scripts = self.tree.insert('', 'end', text='Test Scripts', open=True)
        for script in [f for f in os.listdir() if f.endswith('.py')]:
            self.tree.insert(scripts, 'end', text=script, values=(script,))
            
    def on_tree_double_click(self, event):
        item = self.tree.selection()[0]
        script = self.tree.item(item)['values']
        if script:  # If it's a script, not a folder
            self.launch_script(script[0])
            
    def open_calibration(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if file_path:
            try:
                df = pd.read_csv(file_path)
                self.cal_text.delete('1.0', tk.END)
                self.cal_text.insert('1.0', df.to_string())
            except Exception as e:
                self.terminal.insert(tk.END, f"Error loading calibration file: {str(e)}\n")
                
    def open_results(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if file_path:
            try:
                df = pd.read_csv(file_path)
                self.res_text.delete('1.0', tk.END)
                self.res_text.insert('1.0', df.to_string())
            except Exception as e:
                self.terminal.insert(tk.END, f"Error loading results file: {str(e)}\n")
                
    def launch_terminal(self, terminal_type):
        try:
            if terminal_type == "powershell":
                subprocess.Popen(["powershell"])
            else:
                subprocess.Popen(["cmd"])
        except Exception as e:
            self.terminal.insert(tk.END, f"Error launching {terminal_type}: {str(e)}\n")
            
    def launch_script(self, script):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.terminal.insert(tk.END, f"\n[{timestamp}] Launching {script}...\n")
        self.terminal.see(tk.END)
        
        try:
            process = subprocess.Popen(
                [sys.executable, script],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True
            )
            
            self.terminal.insert(tk.END, f"Process started with PID: {process.pid}\n")
            self.terminal.see(tk.END)
            
            self.root.after(100, self.check_process, process, script)
            
        except Exception as e:
            self.terminal.insert(tk.END, f"Error: {str(e)}\n")
            self.terminal.see(tk.END)
            
    def check_process(self, process, script):
        if process.poll() is None:
            self.root.after(100, self.check_process, process, script)
        else:
            stdout, stderr = process.communicate()
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            if stdout:
                self.terminal.insert(tk.END, f"Output:\n{stdout}\n")
            if stderr:
                self.terminal.insert(tk.END, f"Errors:\n{stderr}\n")
                
            self.terminal.insert(tk.END, f"[{timestamp}] {script} finished with exit code: {process.returncode}\n")
            self.terminal.see(tk.END)

def main():
    root = tk.Tk()
    app = ScriptLauncher(root)
    root.mainloop()

if __name__ == "__main__":
    main()
