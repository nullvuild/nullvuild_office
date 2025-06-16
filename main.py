import tkinter as tk
from tkinter import filedialog, messagebox
import os

class OfficeAutomationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("NV Offie Automation")

        self.input_path = tk.StringVar()
        self.output_folder_path = tk.StringVar()
        self.output_postfix = tk.StringVar()
        self.is_file = False
        self.is_folder = False
        self.modules = []  # List to store available modules

        self.create_widgets()
        self.load_modules()

    def create_widgets(self):
        """Create UI components"""
        tk.Label(self.root, text="Input Path:").grid(row=0, column=0, padx=10, pady=10)
        tk.Entry(self.root, textvariable=self.input_path, width=50).grid(row=0, column=1, padx=10, pady=10)
        tk.Button(self.root, text="Select File", command=self.browse_input_file).grid(row=0, column=2, padx=5, pady=10)
        tk.Button(self.root, text="Select Folder", command=self.browse_input_folder).grid(row=0, column=3, padx=5, pady=10)

        tk.Label(self.root, text="Output Folder:").grid(row=1, column=0, padx=10, pady=10)
        tk.Entry(self.root, textvariable=self.output_folder_path, width=50).grid(row=1, column=1, padx=10, pady=10)
        tk.Button(self.root, text="Browse", command=self.browse_output_folder).grid(row=1, column=2, padx=10, pady=10)

        tk.Label(self.root, text="Output Postfix:").grid(row=2, column=0, padx=10, pady=10)
        tk.Entry(self.root, textvariable=self.output_postfix, width=20).grid(row=2, column=1, padx=10, pady=10)

        self.selected_module = tk.StringVar()
        tk.Label(self.root, text="Select Module:").grid(row=3, column=0, padx=10, pady=10)
        self.module_menu = tk.OptionMenu(self.root, self.selected_module, "")
        self.module_menu.grid(row=3, column=1, padx=10, pady=10)

        tk.Button(self.root, text="Run", command=self.run_automation).grid(row=4, column=1, pady=20)

    def browse_input_file(self):
        """Handle file selection"""
        file_selected = filedialog.askopenfilename()
        if file_selected:
            self.input_path.set(file_selected)
            self.is_file = True
            self.is_folder = False

    def browse_input_folder(self):
        """Handle folder selection"""
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.input_path.set(folder_selected)
            self.is_folder = True
            self.is_file = False

    def browse_output_folder(self):
        """Select the output folder"""
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.output_folder_path.set(folder_selected)

    def load_modules(self):
        """Load available processing modules"""
        self.module_menu['menu'].delete(0, 'end')
        self.modules = ["Common.nv_common_001"]

        for module_name in self.modules:
            self.module_menu['menu'].add_command(
                label=module_name, command=tk._setit(self.selected_module, module_name)
            )

        if self.modules:
            self.selected_module.set(self.modules[0])

    def run_automation(self):
        """Execute the selected module"""
        input_path = self.input_path.get()
        output_folder = self.output_folder_path.get()
        postfix = self.output_postfix.get()

        if not input_path or not output_folder:
            messagebox.showerror("Error", "Both input and output folders must be selected.")
            return

        selected_module = self.selected_module.get()
        if not selected_module:
            messagebox.showerror("Error", "A module must be selected.")
            return

        try:
            module = __import__(selected_module, fromlist=['process'])
            if hasattr(module, 'process'):
                print(f"Executing process from {selected_module}")
                module.process(input_path, output_folder, postfix, self.is_file, self.is_folder)
            else:
                messagebox.showerror("Error", f"The selected module {selected_module} does not have a process function.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to execute {selected_module}: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = OfficeAutomationApp(root)
    root.mainloop()
