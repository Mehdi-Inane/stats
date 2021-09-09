import openpyxl
from openpyxl.utils import get_column_letter
from pathlib import Path
import re
from tkinter import *
from tkinter import filedialog
from tkinter.filedialog import askdirectory
from sys import exit
from tkinter import messagebox
from tkinter import ttk
import os
from gui import *
from syntax import *
from file_creator import *



filelist = []
fichier = openpyxl.Workbook()
sheet = fichier.active
window = Tk()

#Configuration de la GUI
window.title('File Explorer')
window.geometry("500x350")
window.config(background = "#E4E4E0")
text = Text(window, height = 10,bd = 8,width = 50)
text.grid(column=0,row=0,sticky="nsew")
text.insert("1.0","Fichiers choisis : \n")
task = StringVar()
combo = ttk.Combobox(window,width = 5,textvariable = task)
combo['values']=["Hoppy","VOL"]
combo.current(1)
combo.place(x = 150,y=200)
button_frame = Frame(window)
button_frame.grid(column = 1,row = 0)
button_explore = Button(button_frame,
						text = "Explorer",
						command = lambda :browseFiles(filelist,text))

dest_dir = askdirectory(initialdir = "/",title = "Selectionnez l'emplacement de destination")
fdest_dir = os.path.join(dest_dir,"created_files")
try:
    os.makedirs(fdest_dir, exist_ok = True)
except OSError as error:
    print("Directory '%s' can not be created")
button_execute = Button(button_frame,text = "Execution",command = lambda:exec_dump(filelist,sheet,fdest_dir,task.get(),fichier))
button_exit = Button(button_frame,
					text = "Quitter",
					command = exit)
button_clear = Button(button_frame,text = "Tout effacer",command = lambda:clear_list(filelist))
button_frame.columnconfigure(0, weight=1)
button_frame.rowconfigure(1, weight=1)
button_frame.rowconfigure(2, weight=1)
button_frame.rowconfigure(3, weight=1)
button_explore.grid(column = 2, row = 0)
button_execute.grid(column = 2, row = 1)
button_exit.grid(column = 2,row = 2)
button_clear.grid(column = 2,row = 3)
ttk.Label(window,text="Sélectionnez la tâche:",anchor="w").place(x=0,y=200)
window.mainloop()