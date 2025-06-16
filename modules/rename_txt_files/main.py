import os
import tkinter as tk
from tkinter import filedialog, messagebox

def rename_txt_files(directory, prefix):
    """
    지정한 디렉터리의 모든 .txt 파일 이름을 prefix를 붙여서 변경합니다.
    """
    count = 0
    for filename in os.listdir(directory):
        if filename.endswith(".txt"):
            old_path = os.path.join(directory, filename)
            new_filename = prefix + filename
            new_path = os.path.join(directory, new_filename)
            os.rename(old_path, new_path)
            count += 1
    return count

def select_directory():
    dir_path = filedialog.askdirectory()
    if dir_path:
        dir_entry.delete(0, tk.END)
        dir_entry.insert(0, dir_path)

def run_rename():
    directory = dir_entry.get()
    prefix = prefix_entry.get() or "renamed_"
    if not directory or not os.path.isdir(directory):
        messagebox.showerror("오류", "유효한 디렉터리 경로를 입력하세요.")
        return
    count = rename_txt_files(directory, prefix)
    messagebox.showinfo("완료", f"{count}개의 파일 이름이 변경되었습니다.")

root = tk.Tk()
root.title("TXT 파일 이름 변경")

tk.Label(root, text="디렉터리 경로:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
dir_entry = tk.Entry(root, width=40)
dir_entry.grid(row=0, column=1, padx=5, pady=5)
tk.Button(root, text="찾아보기", command=select_directory).grid(row=0, column=2, padx=5, pady=5)

tk.Label(root, text="접두사:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
prefix_entry = tk.Entry(root, width=40)
prefix_entry.insert(0, "renamed_")
prefix_entry.grid(row=1, column=1, padx=5, pady=5)

tk.Button(root, text="이름 변경 실행", command=run_rename).grid(row=2, column=0, columnspan=3, pady=10)

root.mainloop()