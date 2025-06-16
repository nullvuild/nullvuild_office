import tkinter as tk
from tkinter import filedialog, messagebox
import os

class OfficeAutomationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Office Automation")

        self.input_folder_path = tk.StringVar()
        self.output_folder_path = tk.StringVar()
        self.modules = []  # 모듈 리스트 저장

        self.create_widgets()
        self.load_modules()  # 앱 실행 시 자동으로 모듈 로드

    def create_widgets(self):
        tk.Label(self.root, text="Input File:").grid(row=0, column=0, padx=10, pady=10)
        tk.Entry(self.root, textvariable=self.input_folder_path, width=50).grid(row=0, column=1, padx=10, pady=10)
        tk.Button(self.root, text="Browse", command=self.browse_input_folder).grid(row=0, column=2, padx=10, pady=10)

        tk.Label(self.root, text="Output Folder:").grid(row=1, column=0, padx=10, pady=10)
        tk.Entry(self.root, textvariable=self.output_folder_path, width=50).grid(row=1, column=1, padx=10, pady=10)
        tk.Button(self.root, text="Browse", command=self.browse_output_folder).grid(row=1, column=2, padx=10, pady=10)

        self.selected_module = tk.StringVar()
        tk.Label(self.root, text="Select Module:").grid(row=2, column=0, padx=10, pady=10)
        self.module_menu = tk.OptionMenu(self.root, self.selected_module, "")
        self.module_menu.grid(row=2, column=1, padx=10, pady=10)

        tk.Button(self.root, text="Run", command=self.run_automation).grid(row=3, column=1, pady=20)

    def browse_input_folder(self):
        file_selected = filedialog.askopenfilename()
        if file_selected:
            self.input_folder_path.set(file_selected)

    def browse_output_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.output_folder_path.set(folder_selected)

    def load_modules(self):
        """앱 시작 시 자동으로 모듈 리스트를 로드"""
        self.module_menu['menu'].delete(0, 'end')
        self.modules = [  
            "Common.nv_common_001",
            "Common.nv_common_002",
        ]

        # UI 업데이트
        for module_name in self.modules:
            self.module_menu['menu'].add_command(
                label=module_name, command=tk._setit(self.selected_module, module_name)
            )

        # 기본값 설정
        if self.modules:
            self.selected_module.set(self.modules[0])

    def run_automation(self):
        """선택한 모듈을 실행"""
        input_folder = self.input_folder_path.get()
        output_folder = self.output_folder_path.get()

        if not input_folder or not output_folder:
            messagebox.showerror("Error", "Both input and output folders must be selected.")
            return

        selected_module = self.selected_module.get()
        if not selected_module:
            messagebox.showerror("Error", "A module must be selected.")
            return

        try:
            # 동적으로 모듈 로드
            module = __import__(selected_module, fromlist=['process'])
            if hasattr(module, 'process'):
                print(f"Executing process from {selected_module}")
                module.process(input_folder, output_folder)
            else:
                messagebox.showerror("Error", f"The selected module {selected_module} does not have a process function.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to execute {selected_module}: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = OfficeAutomationApp(root)
    root.mainloop()
