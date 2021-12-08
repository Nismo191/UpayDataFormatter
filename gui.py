import tkinter as tk
import tkinter.filedialog as tkFileDialog

import main
import os

output_folder = ""

def load_file(ent):
    file_path = tkFileDialog.askopenfilename()
    ent.insert(0, file_path)

def process_data():
    data_url = ent_dateFile.get()
    teplate_url = ent_templateFile.get()
    try:
        main.main(data_url, teplate_url)
        lbl_status.config(text="Data processeed successfully - File saved in Ouput Folder")
    except Exception as e:
        print(e)
        lbl_status.config(text=e)

def check_for_folder(folder_name):
    if not os.path.exists(folder_name):
        os.makedirs(folder_name)

def open_folder(folder_name):
    os.startfile(folder_name)


if __name__ == '__main__':
    output_folder = check_for_folder("Output")
    check_for_folder("Input")

    window = tk.Tk()
    window.title("Upay Data Formatter")
    window.geometry("800x600")
    window.rowconfigure(0, minsize=100, weight=1)
    window.columnconfigure(0, minsize=100, weight=1)

    frame = tk.Frame(window)

    lbl_dateFile = tk.Label(master=frame, text="Date File")
    lbl_templateFile = tk.Label(master=frame, text="Template File")

    ent_dateFile = tk.Entry(master=frame, width=100)
    ent_templateFile = tk.Entry(master=frame, width=100)

    btn_dateFile = tk.Button(master=frame, text="Browse", command=lambda:load_file(ent_dateFile))
    btn_templateFile = tk.Button(master=frame, text="Browse", command=lambda:load_file(ent_templateFile))

    btn_process = tk.Button(master=frame, text="Process", command=process_data)
    btn_output = tk.Button(master=frame, text="Open Ouput Folder", command=lambda:open_folder("Output"))

    lbl_status = tk.Label(master=frame, text="Status")

    frame.grid(row=0, column=0, sticky="nsew")

    lbl_dateFile.grid(row=0, column=0, sticky="w", padx=5, pady=5)
    lbl_templateFile.grid(row=2, column=0, sticky="w", padx=5, pady=5)
    ent_dateFile.grid(row=0, column=1, sticky="e", padx=5, pady=5)
    ent_templateFile.grid(row=2, column=1, sticky="e", padx=5, pady=5)
    btn_dateFile.grid(row=0, column=2, sticky="e", padx=5, pady=5)
    btn_templateFile.grid(row=2, column=2, sticky="e", padx=5, pady=5)

    btn_process.grid(row=4, column=1, padx=5, pady=5)
    btn_output.grid(row=5, column=1, padx=5, pady=5)
    lbl_status.grid(row=6, column=1, padx=5, pady=5)



    window.mainloop()