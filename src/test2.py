import tkinter as tk
from tkinter import filedialog
import pandas as pd
import os
import glob
from datetime import datetime

def select_directory():
    directory = filedialog.askdirectory()
    directory_label.config(text=directory)
    return directory

def get_column_names():
    columns_input = entry_columns.get()  # Get the input string containing column names
    # Split the input string using a delimiter (assuming ',' as the delimiter)
    column_names = columns_input.split(',')
    
    # You may want to add additional validation or handling for empty strings
    return column_names

def concatenate_excel_files():
    print('*********************')
    directory = directory_label.cget("text")
    if not directory:
        result_label.config(text="디렉토리를 선택하세요.")
        return

    all_files = glob.glob(os.path.join(directory, "*.xlsx"))
    if not all_files:
        result_label.config(text="엑셀 파일이 없습니다.")
        return

    all_data = pd.concat([pd.read_excel(f) for f in all_files], ignore_index=True)
    
    selected_columns = get_column_names()
    data_extract = all_data[selected_columns]
    
    c_time = datetime.now()
    t_time = c_time.strftime("%H%M%S")
    t = str(c_time.year) + '.' + str(c_time.month) + '.' + str(c_time.day) + '.' + str(c_time.hour) + '.' + str(c_time.minute)    
    output_file = os.path.join("/output/", t + "concatenated.xlsx")
    data_extract.to_excel(output_file, index=False)
    result_label.config(text=f"합쳐진 파일 저장됨: {output_file}")
    print(output_file)
    print('저장 완료')

root = tk.Tk()
root.title("엑셀 파일 합치기")

select_button = tk.Button(root, text="디렉토리 선택", command=select_directory)
select_button.pack()

directory_label = tk.Label(root, text="", width=80)
directory_label.pack()

entry_columns_label = tk.Label(root, text="컬럼들(구분자로 구분):")
entry_columns_label.pack()
entry_columns = tk.Entry(root, width=40)
entry_columns.pack()

concat_button = tk.Button(root, text="엑셀 파일 합치기", command=concatenate_excel_files)
concat_button.pack()

result_label = tk.Label(root, text="")
result_label.pack()

root.mainloop()
