import os
from xlrd import open_workbook
import tkinter as tk
from tkinter import ttk

def relocate(level):
    path = os.getcwd()
    new_path = os.path.normpath(os.path.join(path, *([".."] * level)))
    os.chdir(new_path)

def folder_gen(wb,wbloc,sheet,path,root):

    global exit_script
    exit_script = 0

    # load workbook spreadsheet
    os.chdir(wbloc)
    open_wb = open_workbook(wb, on_demand=True)
    sh = open_wb.sheet_by_name(sheet)

    os.chdir(path)

    if not (os.path.exists(root)):
        os.mkdir(root)

        # Status window
        status_win = tk.Tk()
        status_win.title("Status")
        status_win.resizable(width=False, height=False)
        status_win.geometry("300x120")
        progress_bar = ttk.Progressbar(status_win, length=200, mode='determinate')
        progress_bar.place(x=50, y=40)
        label = tk.Label(status_win, text="", anchor='w')
        label.place(x=50, y=15, width=150, height=25)

        def status_close():
            label['text'] = ""
            status_win.update()
            status_win.destroy()

        button = tk.Button(status_win, text="Close Window", command=status_close)
        button.place(x=75, y=80, width=150, height=25)
        status_win.update()
    else:
        # status_win.destroy()
        warning_win = tk.Tk()
        warning_win.title("Warning")
        warning_win.resizable(width=True, height=False)
        warning_win.geometry("300x120")
        warning_win.attributes('-topmost', True)
        label_warning = tk.Label(warning_win, text="Root already exists")
        label_warning.place(x=100, y=30, height=25)

        def abort():
            warning_win.destroy()
            global exit_script
            exit_script = 1

        button_abort = tk.Button(warning_win, text="Abort", command=abort)
        button_abort.place(x=100, y=70, width=100, height=25)
        return

    if exit_script == 1:
        pass

    os.chdir(root)

    old_rowx = 0
    rowx = 1
    colx = 1
    col_lim = False
    prev_colx = 1
    prefix = [-1]*(sh.ncols)
    prog_steps = sh.nrows/10
    prog_steps_rd = round(prog_steps)
    label['text'] = "Processing..."
    # loop through hierarchy on sheet
    while rowx <= sh.nrows:

        # progress bar
        if (rowx % prog_steps_rd == 0) or (rowx == sh.nrows):
            if rowx != old_rowx:
                progress_bar['value'] += (prog_steps_rd/sh.nrows)*100
                status_win.update()

        old_rowx = rowx

        # Build folder structure for activated folders/ sub-folders
        if sh.cell_value(rowx,0) == 1:
            # Avoid indexing out of range at column limit
            if col_lim == True:
                if rowx < sh.nrows-1:
                    rowx += 1
                else:
                    break
                old_colx = colx
                colx = 1
                while colx < sh.ncols:
                    if sh.cell_type(rowx, colx) == 0:  # empty cell
                        if colx < sh.ncols-1:
                            colx += 1
                    else:
                        relocate(old_colx-colx+1)
                        break
                col_lim = False
            # empty cell: navigate to folder on next row
            elif sh.cell_type(rowx, colx) == 0:
                if rowx < sh.nrows-1:
                    rowx += 1
                else:
                    break
                old_colx = colx
                colx = 1
                while colx < sh.ncols:
                    if sh.cell_type(rowx, colx) == 0:  # empty cell
                        if colx < sh.ncols-1:
                            colx += 1
                    else:
                        relocate(old_colx-colx)
                        break
            # value in cell: generate folder
            else:
                # read cell name
                name = str(sh.cell(rowx,colx).value)
                # generate folder prefix number
                if colx <= prev_colx:
                    prefix[colx] += 1
                else:
                    prefix[colx] = 0
                prev_colx = colx
                folder_name = str('0')+str(prefix[colx])+'_'+name
                if not (os.path.exists(folder_name)):
                    os.mkdir(folder_name)
                else:
                    # print("duplicate folder names\ncheck mapping document then try again")
                    break
                # enter directory
                os.chdir(folder_name)
                if colx < sh.ncols-1:
                    colx += 1
                else:
                    col_lim = True

        else:
            while sh.cell_type(rowx+1, colx) == 0:
                rowx += 1

    label['text'] = "Folder generation complete"
    progress_bar['value'] = 100
    status_win.update()
    os.chdir(path)
    return