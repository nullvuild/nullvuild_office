import os
import shutil
import win32com.client

def process_office_file(file_path, app_name, output_folder, postfix):
    """Process Office files (Excel, Word, PowerPoint)"""
    try:
        f_name, f_ext = os.path.splitext(os.path.basename(file_path))
        temp_file = os.path.join(output_folder, f"{f_name}.temp")
        dcrypt_file = os.path.join(output_folder, f"{f_name}{postfix}{f_ext}")

        app = win32com.client.Dispatch(f'{app_name}.Application')
        app.Visible = False

        if app_name == 'Excel':
            workbook = app.Workbooks.Open(file_path)
            workbook.SaveCopyAs(temp_file)
            workbook.Close(False)
        elif app_name == 'Word':
            doc = app.Documents.Open(file_path)
            doc.SaveAs2(temp_file)
            doc.Close(False)
        elif app_name == 'PowerPoint':
            presentation = app.Presentations.Open(file_path)
            presentation.SaveCopyAs(temp_file)
            presentation.Close()

        app.Quit()

        if os.path.isfile(dcrypt_file):
            os.remove(dcrypt_file)
        os.rename(temp_file, dcrypt_file)

        return dcrypt_file
    except Exception as e:
        raise Exception(f"Error processing {app_name} file: {str(e)}")

def process_text_file(file_path, output_folder, postfix):
    """Process text files (TXT, CSV, LOG, etc.)"""
    try:
        f_name, f_ext = os.path.splitext(os.path.basename(file_path))
        dcrypt_file = os.path.join(output_folder, f"{f_name}{postfix}{f_ext}")
        shutil.copy2(file_path, dcrypt_file)
        return dcrypt_file
    except Exception as e:
        raise Exception(f"Error processing text file: {str(e)}")

def process(input_path, output_folder, postfix, is_file, is_folder):
    """Process either a single file or all files in a folder"""
    if is_file:
        file_ext = os.path.splitext(input_path)[1].lower()
        if file_ext in ['.xlsx', '.xls', '.xlsm']:
            process_office_file(input_path, 'Excel', output_folder, postfix)
        elif file_ext in ['.docx', '.doc']:
            process_office_file(input_path, 'Word', output_folder, postfix)
        elif file_ext in ['.pptx', '.ppt']:
            process_office_file(input_path, 'PowerPoint', output_folder, postfix)
        elif file_ext in ['.txt', '.csv', '.log', '.ini', '.md']:
            process_text_file(input_path, output_folder, postfix)
        else:
            process_text_file(input_path, output_folder, postfix)
    elif is_folder:
        for root, _, files in os.walk(input_path):
            for file in files:
                process(os.path.join(root, file), output_folder, postfix, True, False)
