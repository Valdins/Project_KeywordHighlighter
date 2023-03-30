import tkinter as tk
from tkinter import *
from tkinter import filedialog, messagebox, ttk, font

import pandas as pd

# initalise the tkinter GUI
root = tk.Tk()
root.title('Entity extractor')

# End Dataframe
result_file = pd.DataFrame({'Title':[],
                            'Abstract':[],
                            'DOI':[],
                            'DatabaseName':[],
                           'Location':[]})

root.geometry("900x500") # set the root dimensions
root.pack_propagate(False) # tells the root to not let the widgets inside it determine its size.
root.resizable(0, 0) # makes the root window fixed in size.

# Variables
global selected
selected = False

# Frame for Text Box
frame1 = tk.LabelFrame(root, text="Excel Data")
frame1.place(height=400, width=900)

# Frame for open file dialog
file_frame = tk.LabelFrame(root, text="Controls")
file_frame.place(height=100, width=400, rely=0.8, relx=0.3)

# Excel file DataFrame
main_df = ''

# Number of rows
row_number = -1

# Number of rows
total_num_rows = 0

# Number of rows in result file
row_result = 0

# Buttons
button1 = tk.Button(file_frame, text="Browse A File", command=lambda: File_dialog())
button1.place(rely=0.35, relx=0.02)
button1.config(height = 2, width = 10)

button2 = tk.Button(file_frame, text="Next Row", command=lambda: next_row())
button2.place(rely=0.15, relx=0.28)
button2.config(height = 2, width = 10)

button3 = tk.Button(file_frame, text="Copy", command=lambda: copy_text())
button3.place(rely=0.15, relx=0.53)
button3.config(height = 2, width = 10)

button4 = tk.Button(file_frame, text="Export", command=lambda: export_file())
button4.place(rely=0.35, relx=0.78)
button4.config(height = 2, width = 10)

# The info text
label_file = ttk.Label(root, text="No File Selected")
label_file.place(rely=0.8, relx=0.02)

label2 = ttk.Label(root, text="")
label2.place(rely=0.8, relx=0.75)

label3 = ttk.Label(root, text="")
label3.place(rely=0.85, relx=0.75)

label4 = ttk.Label(root, text="")
label4.place(rely=0.95, relx=0.75)

# Scrollbar for text
text_scrolly = Scrollbar(frame1, orient="vertical")
text_scrollx = Scrollbar(frame1, orient="horizontal")
text_scrolly.pack(side="right", fill="y")
text_scrollx.pack(side="bottom", fill="x")

# Text Box
my_text = Text(frame1, width = 97, height=25, font=('Helvetica', 14), selectbackground="yellow",selectforeground="black",undo=True, yscrollcommand=text_scrolly.set, xscrollcommand=text_scrollx.set)
my_text.pack()

# Configure Scrollbar
text_scrolly.config(command=my_text.yview)
text_scrollx.config(command=my_text.xview)

def File_dialog():
    label_file["text"] = "Loading file..... Please Wait"
    """This Function will open the file explorer and assign the chosen file path to label_file"""
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select A File",
                                          filetype=(("xlsx files", "*.xlsx"),("All Files", "*.*")))
    label_file["text"] = filename
    
    #Load file
    Load_excel_data()
    next_row()
    return None

def convert_position():
    value = my_text.index("sel.first").split(".")
    if int(value[0]) > 1:
        line_number = int(value[0])
        char_before = 0
        position = "1.0"
        for num in range(line_number-1):
            num += 2
            new_num = str(num)+".0"
            if len(my_text.get(position,new_num)) > 1:
                char_before += len(my_text.get(position,new_num))
            position = new_num
        result = my_text.index("sel.first").split(".")
        return char_before + int(result[1])
    else:
        result = my_text.index("sel.first").split(".")
        return result[1]

def Load_excel_data():
    """If the file selected is valid this will load the file into the Treeview"""
    file_path = label_file["text"]
    try:
        excel_filename = r"{}".format(file_path)
        if excel_filename[-4:] == ".csv":
            df = pd.read_csv(excel_filename)
        else:
            df = pd.read_excel(excel_filename)

    except ValueError:
        tk.messagebox.showerror("Information", "The file you have chosen is invalid")
        return None
    except FileNotFoundError:
        tk.messagebox.showerror("Information", f"No such file as {file_path}")
        return None
    global main_df
    main_df = df
    
    global total_num_rows
    total_num_rows = len(df.index)
    label2["text"] = "Rows: " + str(total_num_rows)
    
    return None
        
def highlight():
    start = 1.0
    pos = my_text.search("atabase", start, stopindex=END)
    while pos:
        length = len("atabase")
        row, col = pos.split('.')
        end = int(col) + length
        end = row + '.' + str(end)
        my_text.tag_add('highlight', pos, end)
        start = end
        pos = my_text.search("atabase", start, stopindex=END)
    my_text.tag_config('highlight', foreground="red")
    
def next_row():
    global row_number
    global total_num_rows
    
    row_number = row_number + 1
    
    if main_df.empty == True:
        tk.messagebox.showerror("Information", "No Rows available!")
        return None
    
    clear_data()

    df_rows = main_df.loc[row_number,'Abstract'] # turns the dataframe into a list of lists
    
    # line break after dots
    df_rows = df_rows.replace('. ', '.\n\n')
    my_text.insert(END, df_rows)
    highlight()
    
    label3["text"] = "Row: " + str(row_number) + ", out of " + str(total_num_rows)

def clear_data():
    my_text.delete("1.0",END)
    return None
    
def copy_text():
    global selected
    global row_number
    global main_df
    global result_file
    global row_result
    
    if my_text.selection_get():
        selected = my_text.selection_get()
        
    start_position = convert_position()
    end_position = int(start_position) + len(selected)
    start_position =int(start_position) + 1
    loc = str(start_position) + " : " + str(end_position)
    
    #Save data to new DataFrame
    df_temp = {'Title': main_df.loc[row_number,'Title'],
            'Abstract': main_df.loc[row_number,'Abstract'],
            'DOI': main_df.loc[row_number,'DOI'],
            'DatabaseName': selected,
            'Location': loc }
    
    result_file = result_file.append(df_temp, ignore_index=True)
    row_result = len(result_file.index)
    label4["text"] = "Rows in Result_file: " + str(row_result)

def export_file():
    global result_file
    result_file.to_excel("Export_directory.xlsx", sheet_name='Sheet1', index=False)
    label_file["text"] = "File exported to ....Databse_Export.xlsx"
    
root.mainloop()