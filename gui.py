import os
import threading
from tkinter.constants import RIGHT
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
        self.master = master
        self.pack()
        self.create_widgets()

    def create_widgets(self):
        self.listbox = tk.Listbox(self, height=16, width=80)
        self.listbox.grid(column=0, row=0, **common_args)

        self.open_files_btn = tk.Button(self)
        self.open_files_btn["text"] = "Open files"
        self.open_files_btn["command"] = self.open_file
        self.open_files_btn.grid(column=1, row=0, sticky="SE", **common_args)

        self.dest_folder_entry = tk.Entry(self, width=80)
        self.dest_folder_entry.insert(0, self.dest_folder)
        self.dest_folder_entry.grid(column=0, row=1, sticky="E", **common_args)

        self.open_folder_btn = tk.Button(self)
        self.open_folder_btn["text"] = "Open Destination Folder"
        self.open_folder_btn["command"] = self.open_folder
        self.open_folder_btn.grid(column=1, row=1, **common_args)

        self.progress_bar = ttk.Progressbar(self, length=400)
        self.progress_bar["value"] = 0
        self.progress_bar.grid(column=0, row=2, sticky="W", **common_args)

        convert = tk.Button(self, text="Convert", command=self.convert)
        convert.grid(column=1, row=2, sticky="E", **common_args)

        frame = tk.Frame(self)
        frame.grid(column=0, row=3, sticky="E", columnspan=2)

        quit = tk.Button(frame, text="Quit", command=self.master.destroy)
        quit.pack(side=RIGHT, **common_args)

        about = tk.Button(frame, text="About", command=lambda: Dialog(tk.Tk()))
        about.pack(side=RIGHT, **common_args)

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
        self.master = master
        self.pack()

root = tk.Tk()
app = Ui(master=root)
app.mainloop()
