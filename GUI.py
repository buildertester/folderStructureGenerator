import os
from tkinter import *
from tkinter import filedialog
from folder_gen import folder_gen
import tkinter.font as font
from help import help
from xlrd import open_workbook

main_win = Tk()
main_win.title("FolderGenerator")
main_win.resizable(width=True, height=False)
main_win.geometry("600x350")

wb_label = StringVar()
path_label = StringVar()

global wb_gui
wb_gui = ""
global path_gui
path_gui = ""
options = ["Select Sheet"]
selected = StringVar()
selected.set(options[0])

def open_wb():
    main_win.wb = filedialog.askopenfilename(title="Select Mapping Document",filetypes=(("XLSX files", "*.xlsx *.xlsm"), ("all files", "*.*")))
    wb_label.set(main_win.wb)
    global wb_gui
    wb_gui = os.path.basename(main_win.wb)
    global wbloc_gui
    wbloc_gui = os.path.dirname(main_win.wb)

    selected.set(options[0])
    mappingDoc = open_workbook(main_win.wb, on_demand=True)
    options[1::] = mappingDoc.sheet_names()
    dropdown = OptionMenu(main_win, selected, *options)
    dropdown.place(x=100, y=120, width=200, height=25)
    return

def select_path():
    main_win.path = filedialog.askdirectory(title="Select Path")
    path_label.set(main_win.path)
    global path_gui
    path_gui = os.path.abspath(main_win.path)
    return

def run_script():
    root_gui = entry_1.get()
    sheet_gui = selected.get()
    if (not (wb_gui and wb_gui.strip()) or
            (sheet_gui == options[0]) or not
            (path_gui and path_gui.strip()) or not
            (root_gui and root_gui.strip())):
        warning_win = Tk()
        warning_win.title("Warning")
        warning_win.resizable(width=True, height=False)
        warning_win.geometry("250x120")
        warning_win.attributes('-topmost', True)
        label_warning = Label(warning_win, text="Inputs incomplete, try again")
        label_warning.place(x=50, y=30, height=25)
        def close_warning():
            warning_win.destroy()
        button_warning = Button(warning_win, text="OK", command=close_warning)
        button_warning.place(x=75, y=70, width=100, height=25)
        return
    else:
        folder_gen(wb_gui,wbloc_gui,sheet_gui,path_gui,root_gui)
        return

def open_help():
    help_win = Tk()
    help_win.title("Help")
    help_win.resizable(width=False, height=True)
    help_win.geometry("800x550")
    label_help = Label(help_win, text=help.__doc__, anchor='nw', justify=LEFT, wraplength=780)
    label_help.place(x=0, y=5, width=780, height=600)
    return

# Canvas objects
button_1 = Button(text="Select Mapping Document",command=open_wb)
button_1.place(x = 100, y = 60, width=200, height=25)

label_1 = Label(main_win,textvariable=wb_label)
label_1.place(x = 100, y = 90, height=25)

dropdown = OptionMenu(main_win,selected,*options)
dropdown.place(x = 100, y = 120, width=200, height=25)

button_2 = Button(text="Select Path",command=select_path)
button_2.place(x = 100, y = 150, width=200, height=25)

label_2 = Label(main_win,textvariable=path_label)
label_2.place(x = 100, y = 180, height=25)

entry_1 = Entry(main_win)
entry_1.insert(0,"InsertRootName")
entry_1.place(x = 100, y = 210, width=200, height=25)

bold_font = font.Font(weight="bold")
button_3 = Button(text="Generate",command=run_script, font = bold_font)
button_3.place(x = 100, y = 270, width=200, height=40)

button_4 = Button(text="Help",command=open_help)
button_4.place(x = 0, y = 0, width=75, height=25)

main_win.mainloop()