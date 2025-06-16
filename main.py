import os
import subprocess
import sys
import tkinter as tk
from tkinter import messagebox, scrolledtext
from pathlib import Path
from tkinter import simpledialog
import importlib.util

MODULES_DIR = "modules"
proc = None

def get_module_list(search_text=""):
    modules = []
    if os.path.exists(MODULES_DIR):
        for name in os.listdir(MODULES_DIR):
            path = os.path.join(MODULES_DIR, name)
            main_py = os.path.join(path, "main.py")
            if os.path.isdir(path) and os.path.isfile(main_py):
                if search_text.lower() in name.lower():
                    modules.append(name)
    return modules

def show_modules():
    search_text = search_var.get()
    modules = get_module_list(search_text)
    listbox.delete(0, tk.END)
    if modules:
        for m in modules:
            listbox.insert(tk.END, m)
    else:
        listbox.insert(tk.END, "모듈이 없습니다.")

def on_search(*args):
    show_modules()

def style_markdown(text_widget, content):
    text_widget.config(state='normal')
    text_widget.delete(1.0, tk.END)
    lines = content.splitlines()
    for line in lines:
        if line.startswith("# "):
            text_widget.insert(tk.END, line[2:] + "\n", "h1")
        elif line.startswith("## "):
            text_widget.insert(tk.END, line[3:] + "\n", "h2")
        elif line.startswith("### "):
            text_widget.insert(tk.END, line[4:] + "\n", "h3")
        elif line.startswith("- ") or line.startswith("* "):
            text_widget.insert(tk.END, "• " + line[2:] + "\n", "li")
        elif line.strip() == "":
            text_widget.insert(tk.END, "\n")
        else:
            text_widget.insert(tk.END, line + "\n", "normal")
    text_widget.config(state='disabled')

def on_select_module(event):
    selection = listbox.curselection()
    if not selection:
        style_markdown(desc_text, "모듈을 선택하면 설명이 여기에 표시됩니다.")
        return
    module_name = listbox.get(selection[0])
    md_path = os.path.join(MODULES_DIR, module_name, "introduce.md")
    if os.path.exists(md_path):
        with open(md_path, encoding="utf-8") as f:
            content = f.read()
        style_markdown(desc_text, content)
    else:
        style_markdown(desc_text, "설명 파일이 없습니다.")

def run_selected_module():
    selection = listbox.curselection()
    if not selection:
        messagebox.showinfo("알림", "실행할 모듈을 선택하세요.")
        return
    module_name = listbox.get(selection[0])
    module_path = os.path.join(MODULES_DIR, module_name, "main.py")
    if not os.path.exists(module_path):
        messagebox.showerror("오류", f"{module_name} 모듈에 main.py가 없습니다.")
        return
    try:
        # exe로 빌드된 경우 python.exe 경로를 직접 지정
        if getattr(sys, 'frozen', False):
            # pyinstaller로 빌드된 경우
            python_exe = os.path.join(os.path.dirname(sys.executable), "python.exe")
            if not os.path.exists(python_exe):
                # 시스템 PATH에서 python 찾기
                python_exe = "python"
        else:
            python_exe = sys.executable

        if new_window_var.get():
            subprocess.Popen(
                [python_exe, module_path],
                creationflags=subprocess.CREATE_NEW_CONSOLE
            )
        else:
            subprocess.Popen([python_exe, module_path])
    except Exception as e:
        messagebox.showerror("실행 오류", str(e))

root = tk.Tk()
root.title("모듈 리스트 및 설명")

# 프레임 분할
main_frame = tk.Frame(root)
main_frame.pack(fill="both", expand=True, padx=10, pady=10)

left_frame = tk.Frame(main_frame)
left_frame.pack(side="left", fill="y")

right_frame = tk.Frame(main_frame)
right_frame.pack(side="right", fill="both", expand=True)

# 왼쪽: 모듈 리스트
search_var = tk.StringVar()
search_var.trace("w", on_search)
search_entry = tk.Entry(left_frame, textvariable=search_var)
search_entry.pack(pady=5)
search_entry.insert(0, "")

listbox = tk.Listbox(left_frame, width=30, height=20)
listbox.pack(pady=5, fill="y")
listbox.bind("<<ListboxSelect>>", on_select_module)

new_window_var = tk.BooleanVar(value=True)
new_window_check = tk.Checkbutton(left_frame, text="새 커맨드 창에서 실행", variable=new_window_var)
new_window_check.pack(pady=2)

run_btn = tk.Button(left_frame, text="실행", command=run_selected_module)
run_btn.pack(pady=5)

# 오른쪽: 설명란
desc_text = scrolledtext.ScrolledText(right_frame, width=60, height=25, state='disabled')
desc_text.pack(fill="both", expand=True, padx=5, pady=5)

# 마크다운 스타일 태그 정의
desc_text.tag_configure("h1", font=("맑은 고딕", 16, "bold"), foreground="#003366", spacing3=8)
desc_text.tag_configure("h2", font=("맑은 고딕", 13, "bold"), foreground="#005599", spacing3=6)
desc_text.tag_configure("h3", font=("맑은 고딕", 11, "bold"), foreground="#0077aa", spacing3=4)
desc_text.tag_configure("li", font=("맑은 고딕", 10), lmargin1=20, lmargin2=40)
desc_text.tag_configure("normal", font=("맑은 고딕", 10))

style_markdown(desc_text, "모듈을 선택하면 설명이 여기에 표시됩니다.")

show_modules()
root.mainloop()
