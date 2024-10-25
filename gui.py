import json
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.font import Font
import webbrowser
from PIL import Image, ImageTk
from excel_io import read_departments_from_excel
from main import main as run_scheduler
import os
import logging
from utils import resource_path

class Themes:
    def __init__(self):
        self.dark_mode = {
            'bg': '#000000',
            'fg': '#FFFFFF',
            'button_bg': '#FF7900',
            'button_hover': '#E67E22',
            'entry_bg': '#333333',
            'label_bg': '#000000',
            'tooltip_bg': '#333333',
            'tooltip_fg': '#FFFFFF'
        }

        self.light_mode = {
            'bg': '#FFFFFF',
            'fg': '#333333',
            'button_bg': '#FF7900',
            'button_hover': '#E67E22',
            'entry_bg': '#FFFFFF',
            'label_bg': '#FFFFFF',
            'tooltip_bg': '#ffffe0',
            'tooltip_fg': '#333333'
        }

        self.current_theme = self.dark_mode

    def toggle_theme(self):
        if self.current_theme == self.dark_mode:
            self.current_theme = self.light_mode
        else:
            self.current_theme = self.dark_mode

class SchedulerGUI:
    def __init__(self, master):
        self.master = master
        self.themes = Themes()
        self.create_titlebar()
        self.create_main_frame()
        self.create_widgets()
        self.adjust_window_size()
        self.dependency_rules_window = None

    def create_titlebar(self):
        self.titlebar = tk.Frame(self.master, bg=self.themes.current_theme['bg'], height=30)
        self.titlebar.pack(fill=tk.X)

        # Load icons for dark and light mode
        self.dark_icon_image = Image.open(resource_path('light_mode_icon.png'))
        self.dark_icon_image = self.dark_icon_image.resize((20, 20), Image.Resampling.LANCZOS)
        self.dark_icon = ImageTk.PhotoImage(self.dark_icon_image)

        self.light_icon_image = Image.open(resource_path('dark_mode_icon.png'))
        self.light_icon_image = self.light_icon_image.resize((20, 20), Image.Resampling.LANCZOS)
        self.light_icon = ImageTk.PhotoImage(self.light_icon_image)

        # Initialize with the correct icon based on the current theme
        initial_icon = self.dark_icon if self.themes.current_theme == self.themes.dark_mode else self.light_icon
        self.mode_button = ttk.Button(self.titlebar, image=initial_icon, command=self.toggle_mode, style="TitlebarButton.TButton")
        self.mode_button.pack(side=tk.RIGHT, padx=10)

        self.master.title("Scheduling Application")
        self.master.iconphoto(False, tk.PhotoImage(file=resource_path('Orange_small_logo.png')))

    def create_main_frame(self):
        self.configure_styles()
        self.main_frame = ttk.Frame(self.master, padding="15", style="MainFrame.TFrame")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        self.master.columnconfigure(0, weight=1)
        self.master.rowconfigure(1, weight=1)
        
    def configure_styles(self):
        self.style = ttk.Style()
        self.style.theme_use('clam')

        self.style.configure("TitlebarButton.TButton", font=("Helvetica Neue", 12), padding=5, borderwidth=0, background=self.themes.current_theme['button_bg'], foreground=self.themes.current_theme['fg'])
        self.style.map("TitlebarButton.TButton", background=[("active", self.themes.current_theme['button_hover'])])

        self.style.configure("MainFrame.TFrame", background=self.themes.current_theme['bg'])

        self.style.configure("TButton", font=("Helvetica Neue", 12), padding=10, borderwidth=0, background=self.themes.current_theme['button_bg'], foreground=self.themes.current_theme['fg'])
        self.style.map("TButton", background=[("active", self.themes.current_theme['button_hover'])])

        self.style.configure("TLabel", font=("Helvetica Neue", 12), background=self.themes.current_theme['label_bg'], foreground=self.themes.current_theme['fg'])

        self.style.configure("TEntry", font=("Helvetica Neue", 12), fieldbackground=self.themes.current_theme['entry_bg'], foreground=self.themes.current_theme['fg'], padding=5, borderwidth=0)

    def toggle_mode(self, event=None):
        self.themes.toggle_theme()
        self.configure_styles()
        self.update_widgets()
        self.update_tooltips()
        
        # Update the icon based on the current theme
        if self.themes.current_theme == self.themes.dark_mode:
            self.mode_button.config(image=self.dark_icon)
        else:
            self.mode_button.config(image=self.light_icon)

        # If DependencyRulesWindow is open, update it
        if self.dependency_rules_window and self.dependency_rules_window.window.winfo_exists():
            self.dependency_rules_window.update_theme()
        else:
            self.dependency_rules_window = None  # Reset the window instance



    def create_widgets(self):
        self.create_tutorial_link()
        self.create_logo()
        self.create_input_tasks_file()
        self.create_output_schedule_file()
        self.create_run_button()
        self.create_dependency_rules_button()
        self.create_tooltips()

    def create_tutorial_link(self):
        # Tutorial Link
        self.tutorial_link = ttk.Label(self.main_frame, text="How to use this application", foreground="#FF7900", cursor="hand2", font=Font(family="Helvetica Neue", size=14, weight="bold"))
        self.tutorial_link.grid(row=1, column=0, columnspan=3, pady=5)
        self.tutorial_link.bind("<Button-1>", self.open_tutorial)
  
    def create_logo(self):
        logo_image = Image.open(resource_path('53876.png'))
        logo_image = logo_image.resize((100, 100), Image.Resampling.LANCZOS)
        self.logo = ImageTk.PhotoImage(logo_image)

        self.logo_label = tk.Label(self.main_frame, image=self.logo,bg=self.themes.current_theme['bg'])  # Keep this as a tk.Label because it uses images
        self.logo_label.grid(row=0, column=0, columnspan=3, pady=5)


    def create_input_tasks_file(self):
        self.tasks_label = ttk.Label(self.main_frame, text="Input Excel File:", style="TLabel")
        self.tasks_label.grid(row=3, column=0, sticky=tk.E, padx=5, pady=10)

        self.tasks_entry = ttk.Entry(self.main_frame, width=40, style="TEntry")
        self.tasks_entry.grid(row=3, column=1, padx=5, pady=10)
        self.tasks_entry.insert(0, "input.xlsx")

        self.tasks_browse_button = ttk.Button(self.main_frame, text="Browse", command=lambda: self.browse_file(self.tasks_entry), style="TButton")
        self.tasks_browse_button.grid(row=3, column=2, padx=5, pady=10)

    def create_output_schedule_file(self):
        self.output_label = ttk.Label(self.main_frame, text="Output Schedule File:", style="TLabel")
        self.output_label.grid(row=4, column=0, sticky=tk.E, padx=5, pady=10)

        self.output_entry = ttk.Entry(self.main_frame, width=40, style="TEntry")
        self.output_entry.grid(row=4, column=1, padx=5, pady=10)
        self.output_entry.insert(0, "output.xlsx")

        self.output_browse_button = ttk.Button(self.main_frame, text="Browse", command=lambda: self.save_file(self.output_entry), style="TButton")
        self.output_browse_button.grid(row=4, column=2, padx=5, pady=10)

    def create_run_button(self):
        self.run_button = ttk.Button(self.main_frame, text="Run Scheduler", command=self.run_scheduler, style="TButton")
        self.run_button.grid(row=5, column=1, pady=20)
    
    def create_dependency_rules_button(self):
        self.dependency_rules_button = ttk.Button(
            self.main_frame, 
            text="Set Dependency Rules", 
            command=self.open_dependency_rules_window,
            style="TButton"
        )
        self.dependency_rules_button.grid(row=6, column=1, pady=10)
    
    def create_tooltips(self):
        self.tasks_entry_tooltip = self.create_tooltip(self.tasks_entry, "Select the Excel file containing your tasks and resources")
        self.output_entry_tooltip = self.create_tooltip(self.output_entry, "Choose where to save the generated schedule")

    def create_tooltip(self, widget, text):
        def enter(event):
            self.tooltip = tk.Toplevel(self.master)
            self.tooltip.wm_overrideredirect(True)
            self.tooltip.wm_geometry(f"+{event.x_root+15}+{event.y_root+10}")
            label = tk.Label(self.tooltip, text=text, background=self.themes.current_theme['tooltip_bg'], relief="solid", borderwidth=1, foreground=self.themes.current_theme['tooltip_fg'])
            label.pack()

        def leave(event):
            if hasattr(self, 'tooltip'):
                self.tooltip.destroy()

        widget.bind('<Enter>', enter)
        widget.bind('<Leave>', leave)
        return widget

    def update_tooltips(self):
        self.tasks_entry_tooltip.unbind('<Enter>')
        self.tasks_entry_tooltip.unbind('<Leave>')
        self.output_entry_tooltip.unbind('<Enter>')
        self.output_entry_tooltip.unbind('<Leave>')

        self.tasks_entry_tooltip = self.create_tooltip(self.tasks_entry, "Select the Excel file containing your tasks and resources")
        self.output_entry_tooltip = self.create_tooltip(self.output_entry, "Choose where to save the generated schedule")

    def update_widgets(self):
        self.titlebar.config(bg=self.themes.current_theme['bg'])
        self.mode_button.configure(style="TitlebarButton.TButton")

        self.main_frame.configure(style="MainFrame.TFrame")
        self.tutorial_link.configure(style="TLabel")
        self.logo_label.config(bg=self.themes.current_theme['bg'])  # This label uses image, so `bg` is set here
        self.tasks_label.configure(style="TLabel")
        self.tasks_entry.configure(style="TEntry")
        self.tasks_browse_button.configure(style="TButton")
        self.output_label.configure(style="TLabel")
        self.output_entry.configure(style="TEntry")
        self.output_browse_button.configure(style="TButton")
        self.run_button.configure(style="TButton")
        self.dependency_rules_button.configure(style="TButton")



    def adjust_window_size(self):
        # Prevent window resizing
        self.master.resizable(False, False)

        # Get the window size
        self.master.update_idletasks()  # Ensure the window size is calculated
        width = self.master.winfo_reqwidth()
        height = self.master.winfo_reqheight()

        # Get the screen size
        screen_width = self.master.winfo_screenwidth()
        screen_height = self.master.winfo_screenheight()

        # Calculate the position to center the window
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)

        # Set the window position and size
        self.master.geometry(f"{width}x{height}+{x}+{y}")

        # Optionally, set a minimum size if you want to enforce one
        self.master.minsize(width, height)


    def browse_file(self, entry):
        filename = filedialog.askopenfilename(initialfile="input.xlsx", filetypes=[("Excel files", "*.xlsx")])
        entry.delete(0, tk.END)
        entry.insert(0, filename)

    def save_file(self, entry):
        filename = filedialog.asksaveasfilename(initialfile="output.xlsx", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        entry.delete(0, tk.END)
        entry.insert(0, filename)

    def run_scheduler(self):
        input_file = self.tasks_entry.get()
        output_file = self.output_entry.get()

        if not input_file or not output_file:
            messagebox.showerror("Error", "Please select all required files.")
            return

        try:
            self.master.after(100, self._run_scheduler_thread, input_file, output_file)
        except Exception as e:
            logging.exception("An error occurred while running the scheduler")
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    def _run_scheduler_thread(self, input_file, output_file):
        try:
            run_scheduler(input_file, output_file)
            messagebox.showinfo("Success", "Scheduling completed successfully!")
        except Exception as e:
            logging.exception("An error occurred while running the scheduler")
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    def open_tutorial(self, event):
        tutorial_path = resource_path('instructions.html')
        webbrowser.open_new("file://" + os.path.abspath(tutorial_path))
    
    def open_dependency_rules_window(self):
        input_file = self.tasks_entry.get()
        if not input_file:
            messagebox.showerror("Error", "Please select an input file first.")
            return

        if not self.dependency_rules_window or not self.dependency_rules_window.window.winfo_exists():
            self.dependency_rules_window = DependencyRulesWindow(self.master, input_file, self.themes)
        else:
            self.dependency_rules_window.window.lift()

class DependencyRulesWindow:
    def __init__(self, parent, input_file, themes):
        self.parent = parent
        self.input_file = input_file
        self.themes = themes
        self.window = tk.Toplevel(parent)
        self.window.title("Department Dependency Rules")
        self.window.geometry("600x400")
        self.configure_styles()
        self.create_widgets()
        self.window.iconphoto(False, tk.PhotoImage(file=resource_path('Orange_small_logo.png')))

    def configure_styles(self):
        self.style = ttk.Style()
        self.style.theme_use('clam')

        # Configure styles using current theme
        self.style.configure("DepartmentFrame.TFrame", background=self.themes.current_theme['bg'])
        self.style.configure("ScrollableFrame.TFrame", background=self.themes.current_theme['bg'])
        self.style.configure("TButton", font=("Helvetica Neue", 12), padding=10, borderwidth=0, background=self.themes.current_theme['button_bg'], foreground=self.themes.current_theme['fg'])
        self.style.map("TButton", background=[("active", self.themes.current_theme['button_hover'])])
        self.style.configure("TLabel", font=("Helvetica Neue", 12), background=self.themes.current_theme['bg'], foreground=self.themes.current_theme['fg'])
        self.style.configure("TCheckbutton", font=("Helvetica Neue", 12), background=self.themes.current_theme['bg'], foreground=self.themes.current_theme['fg'])
        self.style.map("TCheckbutton", background=[("hover", self.themes.current_theme['button_hover'])])
        
        self.window.configure(bg=self.themes.current_theme['bg'])

    def create_widgets(self):
        self.create_department_frame()
        self.create_save_button()

    def create_department_frame(self):
        self.dept_frame = ttk.Frame(self.window, style="DepartmentFrame.TFrame")
        self.dept_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.canvas = tk.Canvas(self.dept_frame, bg=self.themes.current_theme['bg'])
        self.scrollbar = ttk.Scrollbar(self.dept_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas, style="ScrollableFrame.TFrame")

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.dept_frame.pack(fill=tk.BOTH, expand=True)
        self.canvas.pack(side="left", fill=tk.BOTH, expand=True)
        self.scrollbar.pack(side="right", fill="y")

        self.load_dependencies()

    def create_save_button(self):
        save_button = ttk.Button(self.window, text="Save Dependencies", command=self.save_dependencies, style="TButton")
        save_button.pack(pady=10)

    def load_dependencies(self):
        try:
            with open('dependency_rules.json', 'r') as f:
                self.dependencies = json.load(f)
        except FileNotFoundError:
            self.dependencies = {}

        self.departments = read_departments_from_excel(self.input_file)
        self.dependency_vars = {}

        # Create header
        ttk.Label(self.scrollable_frame, text="Department", font=("Helvetica", 10, "bold"), style="TLabel").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        ttk.Label(self.scrollable_frame, text="Must be completed before", font=("Helvetica", 10, "bold"), style="TLabel").grid(row=0, column=1, sticky="w", padx=5, pady=2)

        for i, dept in enumerate(self.departments, start=1):
            label = ttk.Label(self.scrollable_frame, text=dept, style="TLabel")
            label.grid(row=i, column=0, sticky="w", padx=5, pady=2)
            
            self.dependency_vars[dept] = {}
            dep_frame = ttk.Frame(self.scrollable_frame, style="ScrollableFrame.TFrame")
            dep_frame.grid(row=i, column=1, sticky="w", padx=5, pady=2)

            for j, dep_dept in enumerate(self.departments):
                if dept != dep_dept:
                    var = tk.BooleanVar(value=dep_dept in self.dependencies.get(dept, []))
                    self.dependency_vars[dept][dep_dept] = var
                    check = ttk.Checkbutton(dep_frame, text=dep_dept, variable=var, style="TCheckbutton")
                    check.grid(row=0, column=j, sticky="w", padx=2, pady=2)

    def save_dependencies(self):
        new_dependencies = {}
        for dept, deps in self.dependency_vars.items():
            new_dependencies[dept] = [d for d, v in deps.items() if v.get()]
        
        with open('dependency_rules.json', 'w') as f:
            json.dump(new_dependencies, f, indent=4)
        
        self.window.destroy()
        messagebox.showinfo("Success", "Dependency rules saved successfully!")

    def update_theme(self):
        """Update theme styles and widget configurations."""
        self.configure_styles()
        
        # Update the background and styles of all existing widgets
        self.window.configure(bg=self.themes.current_theme['bg'])
        self.dept_frame.configure(style="DepartmentFrame.TFrame")
        self.canvas.configure(bg=self.themes.current_theme['bg'])
        self.scrollable_frame.configure(style="ScrollableFrame.TFrame")
        
        for child in self.scrollable_frame.winfo_children():
            if isinstance(child, ttk.Label):
                child.configure(style="TLabel")
            elif isinstance(child, ttk.Checkbutton):
                child.configure(style="TCheckbutton")

        for child in self.window.winfo_children():
            if isinstance(child, ttk.Button):
                child.configure(style="TButton")




def main():
    root = tk.Tk()
    app = SchedulerGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()