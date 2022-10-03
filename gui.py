import os
import threading
from tkinter.constants import RIGHT
from tkinter.constants import LEFT
from typing import List
from main import main as convert_to_pdf

import tkinter as tk
import tkinter.ttk as ttk
import tkinter.filedialog as filedialog

common_args = {"padx": 8, "pady": 8}


class Ui(tk.Frame):
    master: tk.Tk

    listbox: tk.Listbox
    open_files_btn: tk.Button
    dest_folder_entry: tk.Entry
    open_folder_btn: tk.Button
    progress_bar: ttk.Progressbar

    selected_files: List[str]
    dest_folder: str = os.getcwd()

    def __init__(self, master=None):
        super().__init__(master)
        self.selected_files = []

        self.master.title("Any to PDF")
        self.master.resizable(False, False)
        self.master = master
        self.pack()
        self.create_widgets()

    def create_widgets(self):
        self.listbox = tk.Listbox(self, height=10, width=100)
        self.listbox.grid(column=0, row=0, columnspan=3, **common_args)

        self.dest_folder_entry = tk.Entry(self, width=100)
        self.dest_folder_entry.insert(0, self.dest_folder)
        self.dest_folder_entry.grid(column=0, row=1, columnspan=3, sticky="E", **common_args)

        self.open_files_btn = tk.Button(self, width=20)
        self.open_files_btn["text"] = "Open Files"
        self.open_files_btn["command"] = self.open_file
        self.open_files_btn.grid(column=0, row=2, columnspan=2, **common_args)

        self.open_folder_btn = tk.Button(self, width=20)
        self.open_folder_btn["text"] = "Open Destination Folder"
        self.open_folder_btn["command"] = self.open_folder
        self.open_folder_btn.grid(column=1, row=2, columnspan=2, **common_args)

        self.progress_bar = ttk.Progressbar(self, length=400)
        self.progress_bar["value"] = 0
        self.progress_bar.grid(column=0, row=3, columnspan=2, sticky="W", **common_args)

        convert = tk.Button(self, text="Convert", command=self.convert, width=20)
        convert.grid(column=2, row=3, sticky="E", **common_args)

        frame = tk.Frame(self)
        frame.grid(column=0, row=4, columnspan=3)

        quit = tk.Button(frame, text="Quit", command=self.master.destroy, width=10)
        quit.pack(side=RIGHT, **common_args)

        about = tk.Button(frame, text="About", command=lambda: Dialog(tk.Tk()), width=10)
        about.pack(side=LEFT, **common_args)

        tk.Label(frame, text="Developed by Suraj Singh").pack(side=LEFT ,**common_args)

        # tk.Label(self, text="Developed by Suraj Singh").grid(column=0, row=5, columnspan=3)


    def open_file(self):
        self.selected_files = list(filedialog.askopenfilenames())
        self.listbox.delete(0, self.listbox.size() - 1)
        self.listbox.insert(0, *self.selected_files)

    def open_folder(self):
        self.dest_folder = filedialog.askdirectory()
        self.dest_folder_entry.delete(0, len(self.dest_folder_entry.get()) - 1)
        self.dest_folder_entry.insert(0, self.dest_folder)

    def convert(self):
        if len(self.selected_files) == 0:
            return

        def _convert():
            for i, file in enumerate(self.selected_files, start=1):
                print(f"[*] Converting {file}")
                convert_to_pdf(file, self.dest_folder)
                self.progress_bar["value"] = i / len(self.selected_files) * 100

        threading.Thread(target=_convert).start()


class Dialog(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.selected_files = []

        self.master.title("About")
        self.master.geometry('480x200')
        self.master.resizable(False, False)
        tk.Label(self.master, text="This is a tool that converts txt, docx, pptx and xlsx files to pdf files with ease.").pack()
        self.master = master
        self.pack()

root = tk.Tk()
app = Ui(master=root)
app.mainloop()
