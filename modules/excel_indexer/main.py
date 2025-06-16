import tkinter as tk
from tkinter import messagebox
import win32com.client

def get_open_workbooks():
    excel = win32com.client.Dispatch("Excel.Application")
    wb_list = []
    for wb in excel.Workbooks:
        wb_list.append(wb.Name)
    return wb_list

def create_index_sheet(wb_name):
    excel = win32com.client.Dispatch("Excel.Application")
    wb = None
    for w in excel.Workbooks:
        if w.Name == wb_name:
            wb = w
            break
    if wb is None:
        return "선택한 파일을 찾을 수 없습니다."

    # Index 시트 생성 또는 초기화
    try:
        idx_ws = wb.Worksheets("Index")
        idx_ws.Cells.Clear()
    except:
        idx_ws = wb.Worksheets.Add(Before=wb.Worksheets(1))
        idx_ws.Name = "Index"

    idx_ws.Cells(1, 1).Value = "시트 목록"
    row = 2
    for ws in wb.Worksheets:
        if ws.Name != "Index":
            display_text = f"▶ {ws.Name}"
            hl = idx_ws.Hyperlinks.Add(
                Anchor=idx_ws.Cells(row, 1),
                Address="",
                SubAddress=f"'{ws.Name}'!A1",
                TextToDisplay=display_text
            )
            idx_ws.Cells(row, 1).Value = display_text  # 추가로 한 번 더 지정
            row += 1

    return f"Index 시트가 생성되었습니다. ({row-2}개 시트 링크)"

def run_indexer():
    selection = listbox.curselection()
    if not selection:
        messagebox.showerror("오류", "엑셀 파일을 선택하세요.")
        return
    wb_name = listbox.get(selection[0])
    result = create_index_sheet(wb_name)
    messagebox.showinfo("결과", result)

root = tk.Tk()
root.title("열려있는 엑셀 파일 Index 시트 생성기")

tk.Label(root, text="열려있는 엑셀 파일 목록:").pack(pady=5)

wb_names = get_open_workbooks()
listbox = tk.Listbox(root, width=50, height=10)
for name in wb_names:
    listbox.insert(tk.END, name)
listbox.pack(pady=5)

tk.Button(root, text="Index 시트 생성", command=run_indexer).pack(pady=10)

root.mainloop()

# Excel Index 시트 생성기

# 이 프로그램은 **열려있는 엑셀 파일**의 시트 목록을 자동으로 정리해주는 `Index` 시트를 생성합니다.  
# 각 시트 이름에 하이퍼링크가 걸려 있어 클릭하면 해당 시트로 바로 이동할 수 있습니다.

# ## 주요 기능

# - 현재 열려있는 모든 엑셀 파일을 자동으로 탐색
# - 선택한 엑셀 파일에 `Index` 시트 생성
# - 기존에 `Index` 시트가 있으면 초기화 후 재생성
# - 각 시트 이름에 하이퍼링크 자동 추가

# ## 사용 방법

# 1. **엑셀 파일을 열어둡니다.**
# 2. 프로그램을 실행합니다.
# 3. 프로그램 창에 열려있는 엑셀 파일 목록이 표시됩니다.
# 4. 목록에서 Index 시트를 만들 파일을 선택합니다.
# 5. `Index 시트 생성` 버튼을 클릭합니다.
# 6. 선택한 엑셀 파일에 `Index` 시트가 생성되고, 각 시트로 이동할 수 있는 하이퍼링크가 추가됩니다.

# ## 실행 환경

# - Windows
# - Python 3.x
# - `tkinter` (GUI)
# - `pywin32` (win32com)

# ## 설치 방법

# ```bash
# pip install pywin32
# ```

# ## 코드 주요 부분

# - **get_open_workbooks**: 현재 열려있는 엑셀 파일 목록을 가져옵니다.
# - **create_index_sheet**: 선택한 엑셀 파일에 Index 시트를 생성합니다.
# - **run_indexer**: GUI에서 선택된 파일에 대해 Index 시트 생성을 실행합니다.

# ## 참고

# - 반드시 엑셀 파일을 **미리 열어둔 상태**에서 사용하세요.
# - Index 시트는 항상 첫 번째 시트로 생성됩니다.