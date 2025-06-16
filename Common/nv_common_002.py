import os
import tkinter as tk
from tkinter import filedialog, messagebox
import win32com.client
import shutil

def process_office_file(file_path, app_name):
    """Office 파일(Excel, Word, PowerPoint)을 처리하는 함수"""
    try:
        # 파일 경로와 확장자 분리
        f_path_name = os.path.splitext(file_path)[0]
        f_ext = os.path.splitext(file_path)[1]

        # 새 파일명 생성 (_dcrypt 접미사 포함)
        temp_file = f_path_name + '.temp'
        dcrypt_file = f_path_name + '_dcrypt' + f_ext

        # Office 애플리케이션 객체 생성
        app = win32com.client.Dispatch(f'{app_name}.Application')
        app.Visible = False

        if app_name == 'Excel':
            # Excel 파일 처리
            workbook = app.Workbooks.Open(file_path)
            workbook.SaveCopyAs(temp_file)
            workbook.Close(False)
        elif app_name == 'Word':
            # Word 파일 처리
            doc = app.Documents.Open(file_path)
            doc.SaveAs2(temp_file)
            doc.Close(False)
        elif app_name == 'PowerPoint':
            # PowerPoint 파일 처리
            presentation = app.Presentations.Open(file_path)
            presentation.SaveCopyAs(temp_file)
            presentation.Close()

        app.Quit()

        # 기존 파일이 있으면 삭제 후 이름 변경
        if os.path.isfile(dcrypt_file):
            os.remove(dcrypt_file)
        os.rename(temp_file, dcrypt_file)

        return dcrypt_file
    except Exception as e:
        raise Exception(f"{app_name} 파일 처리 중 오류 발생: {str(e)}")

def process_text_file(file_path):
    """일반 텍스트 파일을 처리하는 함수"""
    try:
        # 파일 경로와 확장자 분리
        f_path_name = os.path.splitext(file_path)[0]
        f_ext = os.path.splitext(file_path)[1]

        # 새 파일명 생성 (_dcrypt 접미사 포함)
        dcrypt_file = f_path_name + '_dcrypt' + f_ext

        # 파일 복사
        shutil.copy2(file_path, dcrypt_file)

        return dcrypt_file
    except Exception as e:
        raise Exception(f"텍스트 파일 처리 중 오류 발생: {str(e)}")

def process_file(file_path, input_folder, output_folder):
    """파일 형식에 따라 적절한 처리 함수를 호출"""
    file_ext = os.path.splitext(file_path)[1].lower()

    # 파일 형식에 따른 처리
    if file_ext in ['.xlsx', '.xls', '.xlsm']:
        return process_office_file(file_path, 'Excel')
    elif file_ext in ['.docx', '.doc']:
        return process_office_file(file_path, 'Word')
    elif file_ext in ['.pptx', '.ppt']:
        return process_office_file(file_path, 'PowerPoint')
    elif file_ext in ['.txt', '.csv', '.log', '.ini', '.md']:
        return process_text_file(file_path)
    else:
        # 알 수 없는 파일 형식은 단순 복사
        return process_text_file(file_path)

def select_and_process_file(input_folder, output_folder):
    """파일 선택 대화상자를 열고 선택한 파일을 처리"""
    # 파일 선택 대화상자
    file_path = input_folder

    if file_path:
        try:
            # 파일 확장자 확인
            file_ext = os.path.splitext(file_path)[1].lower()
            file_type = "알 수 없는 형식"

            if file_ext in ['.xlsx', '.xls', '.xlsm']:
                file_type = "Excel"
            elif file_ext in ['.docx', '.doc']:
                file_type = "Word"
            elif file_ext in ['.pptx', '.ppt']:
                file_type = "PowerPoint"
            elif file_ext in ['.txt', '.csv', '.log', '.ini', '.md']:
                file_type = "텍스트"

            # 파일 처리 함수 호출
            result_file = process_file(file_path, input_folder, output_folder)

            # 성공 메시지
            messagebox.showinfo(
                "완료",
                f"{file_type} 파일이 성공적으로 처리되었습니다:\n{os.path.basename(result_file)}"
            )
        except Exception as e:
            # 오류 메시지
            messagebox.showerror("오류", str(e))

def process(input_folder, output_folder):
    select_and_process_file(input_folder, output_folder)
