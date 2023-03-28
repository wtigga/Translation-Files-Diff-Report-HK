import openpyxl
import pandas as pd
from diff_match_patch import diff_match_patch
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import *
from tkinter import messagebox
import time
import webbrowser
import os

# Set Pandas display options
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
pd.set_option('display.expand_frame_repr', False)

### VARIABLES ###
# Source Files
# File 1
source_file_one_simplename = 'Default File1 name'
source_file_one = 'venv/source_files/real_Files/one.xlsx'
#source_file_one = 'venv/source_files/one.xlsx' #default

# File 2
source_file_two_simplename = 'Default File2 name'
source_file_two = 'venv/source_files/real_Files/two.xlsx'
#source_file_two = 'venv/source_files/two.xlsx' #default

# Output HTML report path
output_file = 'venv/report.html'

# Language list
language_codes = ['en', 'kr', 'cht', 'jp', 'th', 'vi', 'id', 'es', 'ru', 'pt', 'de', 'fr']

# Location of content in source files, case-sensitive
string_id_column = 'ID'
source_lang_column = 'CHS'
target_lang_column = 'ru'


### FUNCTIONALITY ###

# Load the source XLSX into memory
def open_excel_file(file_path):
    workbook = openpyxl.load_workbook(file_path)
    return workbook


def get_column_index(sheet, column_name):
    for cell in sheet[1]:  # Assuming the first row has headers
        if cell.value == column_name:
            return cell.column - 1
    return None


# Creating dataframe from the source file

def create_dataframe(workbook): #one useless argument for now
    all_data = []

    for sheet in workbook.worksheets:
        sheet_data = []

        string_id_index = get_column_index(sheet, string_id_column)
        source_lang_index = get_column_index(sheet, source_lang_column)
        target_lang_index = get_column_index(sheet, target_lang_column)

        if string_id_index is None or source_lang_index is None or target_lang_index is None:
            continue

        for row in sheet.iter_rows(min_row=2, values_only=True):  # Assuming the first row has headers
            sheet_data.append({
                'Sheet name': sheet.title,
                'ID': row[string_id_index],
                'Source': row[source_lang_index],
                'Target': row[target_lang_index]
            })

        all_data.extend(sheet_data)

    dataframe = pd.DataFrame(all_data)
    return dataframe

# Merging two dataframes from two files that has to be compared

def merging_df(dataframe1, dataframe2):
    result = pd.merge(dataframe1, dataframe2, on=['Sheet name', 'ID'], how='outer')
    result = result[['Sheet name', 'ID', 'Source1', 'Source2', 'Target1', 'Target2']]
    return result

# Filter rows to remove those where there's nothing changed
def filter_dataframe(df):
    return df.loc[
        ~(
            ((df['Source1'] == df['Source2']) | (pd.isna(df['Source1']) & pd.isna(df['Source2']))) &
            ((df['Target1'] == df['Target2']) | (pd.isna(df['Target1']) & pd.isna(df['Target2'])))
        )
    ]

# Adding columns with difference in source and target
def add_diff_columns(df):
    df = df.copy()

    dmp = diff_match_patch()

    def compute_diff(text1, text2):
        if pd.isna(text1) or pd.isna(text2):
            return None
        diffs = dmp.diff_main(text1, text2)
        dmp.diff_cleanupSemantic(diffs)
        html_diff = dmp.diff_prettyHtml(diffs)
        return html_diff

    df['Source_diff'] = df.apply(lambda row: compute_diff(row['Source1'], row['Source2']), axis=1)
    df['Target_diff'] = df.apply(lambda row: compute_diff(row['Target1'], row['Target2']), axis=1)
    return df

#Saving the file to HTML report
def save_df_to_html(df, file_name):
    html = df.to_html(escape=False)  # Set escape=False to render HTML content in the DataFrame
    with open(file_name, "w", encoding="utf-8") as f:
        f.write(html)



# Processing files
def process_files(source1, source2):
    workbook1 = open_excel_file(source1.get())
    df1 = create_dataframe(workbook1)
    df1 = df1.rename(columns={'Source': 'Source1', 'Target': 'Target1'})

    workbook2 = open_excel_file(source2.get())
    df2 = create_dataframe(workbook2)
    df2 = df2.rename(columns={'Source': 'Source2', 'Target': 'Target2'})

    merged_df = merging_df(df1, df2)
    filtered_df = filter_dataframe(merged_df)
    filtered_df_with_diff = add_diff_columns(filtered_df)

    return filtered_df_with_diff

### GUI ###
def browse_file_one():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    source_file_one.set(file_path)

def browse_file_two():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    source_file_two.set(file_path)

def save_file():
    file_path = filedialog.asksaveasfilename(filetypes=[("HTML files", "*.html")])
    output_file.set(file_path)

def execute_program():
    global source_file_one_name
    global source_file_two_name
    source_file_one_name = 'filename1'
    source_file_one_name = os.path.splitext(os.path.basename(source_file_one.get()))[0]
    source_file_two_name = 'filename2'
    source_file_two_name = os.path.splitext(os.path.basename(source_file_two.get()))[0]
#    progress_bar["maximum"] = 4
#    progress_bar["value"] = 0
    root.update_idletasks()

    # Replace this section with the actual steps of your process
    for i in range(1, 5):
#        progress_bar["value"] = i
        root.update_idletasks()
        time.sleep(1)  # Replace this with your processing logic

    # Add your code to process the files here
    filtered_df_with_diff = process_files(source_file_one, source_file_two)
    filtered_df_with_diff = filtered_df_with_diff.rename(columns={'Source1': ('Source 1: ' +str(source_file_one_name)),'Source2': ('Source 2: ' +str(source_file_two_name))})
    filtered_df_with_diff = filtered_df_with_diff.rename(columns={'Target1': ('Target 1: ' +str(source_file_one_name)),'Target2': ('Target 2: ' +str(source_file_two_name))})
    save_df_to_html(filtered_df_with_diff, output_file.get())
    messagebox.showinfo("Process complete", "Report has been generated. You can find it at: " + str(output_file.get()))
    pass

def exit_program():
    root.destroy()

root = Tk()
root.geometry("460x320")

# Set the window title
root.title("HK Diff Checker, v2023-03-28")

source_file_one = StringVar()
source_file_two = StringVar()
output_file = StringVar()

target_lang_code = tk.StringVar()
target_lang_combobox = ttk.Combobox(root, textvariable=target_lang_column, values=language_codes, width=6)
target_lang_combobox.current(language_codes.index('ru'))
target_lang_combobox.grid(row=5, column=0, sticky='w', padx=10, pady=10)

browse_one_button = Button(root, text="Browse file #1", command=browse_file_one)
browse_one_button.grid(row=0, column=0, sticky='w', padx=10, pady=10)

file_one_entry = Entry(root, textvariable=source_file_one, width=50)
file_one_entry.grid(row=0, column=1, sticky='w', padx=10, pady=10)

browse_two_button = Button(root, text="Browse file #2", command=browse_file_two)
browse_two_button.grid(row=1, column=0, sticky='w', padx=10, pady=10)

file_two_entry = Entry(root, textvariable=source_file_two, width=50)
file_two_entry.grid(row=1, column=1, sticky='w', padx=10, pady=10)

save_button = Button(root, text="Save report to...", command=save_file)
save_button.grid(row=2, column=0, sticky='w', padx=10, pady=10)

save_entry = Entry(root, textvariable=output_file, width=50)
save_entry.grid(row=2, column=1, sticky='w', padx=10, pady=10)

process_button = Button(root, text="CHECK DIFF", command=execute_program)
process_button.grid(row=3, column=0, sticky='w', padx=10, pady=10)

exit_button = Button(root, text="Exit", command=exit_program)
exit_button.grid(row=10, column=0, sticky='w', padx=50, pady=10)

#progress_bar = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
#progress_bar.grid(column=0, row=5, columnspan=3, sticky='w')


# Text in the bottom
def open_url(url):
    webbrowser.open(url)
about_label = tk.Label(root, text="github.com/wtigga  ||  Vladimir Zhdanov, 2023-03-28", fg="blue", cursor="hand2")
about_text = tk.Label(root, text="Diff for translation sources.")
about_text.grid(row=6, column=1, sticky='e', padx=0, pady=0)
about_label.bind("<Button-1>", lambda event: open_url("https://github.com/wtigga/Translation-Files-Diff-Report-HK"))
about_label.grid(row=7, column=1, sticky='e', padx=10, pady=0)


### EXECUTING ###

root.mainloop()


'''While the logic and architecture are products of the author's thinking capabilities,
lots of functions in the code were written with the help of OpenAi's ChatGPT 3.5 and ChatGPT 4.'''