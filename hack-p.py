import os
import datetime
from openpyxl import Workbook
import tkinter as tk
from tkinter import filedialog

def list_files_with_details(directory):
    file_types = ['.xlsx', '.docx', '.pdf', '.pptx', '.jpg', '.png', '.jpeg', '.mp3', '.mp4', '.txt']
    file_details = {}

    for root, dirs, files in os.walk(directory):
        for file in files:
            file_path = os.path.join(root, file)
            file_type = get_file_type(file_path, file_types)
            if file_type:
                file_size = os.path.getsize(file_path)
                created_date = get_created_date(file_path)
                if file_type not in file_details:
                    file_details[file_type] = []
                file_details[file_type].append([file, file_path, file_type, file_size, created_date])

    save_to_excel(file_details)

def get_file_type(file_path, file_types):
    for file_type in file_types:
        if file_path.lower().endswith(file_type):
            return file_type[1:].upper() + " File"
    return None

def get_created_date(file_path):
    created_timestamp = os.path.getctime(file_path)
    created_date = datetime.datetime.fromtimestamp(created_timestamp).strftime('%Y-%m-%d %H:%M:%S')
    return created_date

def save_to_excel(file_details):
    index = 1
    while True:
        filename = f"temp{index}.xlsx"
        if not os.path.exists(filename):
            wb = Workbook()
            for file_type, details in file_details.items():
                ws = wb.create_sheet(title=file_type)
                ws.append(["File Name", "File Path", "File Type", "File Size (bytes)", "Created Date"])
                for detail in details:
                    ws.append(detail)
            
            wb.remove(wb['Sheet'])
            wb.save(filename)
            print(f"File details sorted by type saved in '{filename}'")
            break
        else:
            index += 1

def get_drive():
    drive = filedialog.askdirectory()
    if drive:
        list_files_with_details(drive)

root = tk.Tk()
root.title("File Details Collector")

drive_button = tk.Button(root, text="Select Drive", command=get_drive, bg="white", fg="black")
drive_button.pack(padx=100, pady=70)

root.mainloop()
