import tkinter as tk
import _tkinter
from tkinter.filedialog import askopenfilename, asksaveasfilename

import time

import _thread

import os

from vidutinis_akcizas import run_calculation, run_calculation_from_ui

LARGE_FONT = ('Verdana', 12)

class VidutinisAkcizas(tk.Tk):

    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)
        container = tk.Frame(self)

        container.pack(side='top', fill='both', expand = True)

        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        self.frames = {}

        frame = StartPage(container, self)

        self.frames[StartPage] = frame

        frame.grid(row=0, column=0, sticky='nsew')

        self.show_frame(StartPage)

    def show_frame(self, cont):
        frame = self.frames[cont]
        frame.tkraise()


class StartPage(tk.Frame):

    def __init__(self, parent, controller):

        self.file_text = tk.StringVar()

        tk.Frame.__init__(self, parent)
        label = tk.Label(self, text="Akcizo skaičiavimas", font=LARGE_FONT)
        label.pack(pady=10, padx=10)

        text = tk.Label(self, textvariable=self.file_text)
        text.pack(pady=5, padx=5)

        browse_btn = tk.Button(self, text="Pasirinkti",
                          command=self.browse_func)
        browse_btn.pack(pady=5, padx=5)

        self.save_btn = tk.Button(self, text="Išsaugoti",
                             command=self.save_func)



    def browse_func(self):
        filename = askopenfilename()
        self.save_file_path = filename
        if filename is not None and filename is not "":
            self.save_btn.pack(pady=5, padx=5)
            self.file_text.set(filename)

    def save_func(self):
        path = asksaveasfilename()
        path = os.path.splitext(path)[0]
        self.file_text.set("Skaičiuojama...")
        time.sleep(1)
        if path != None and path != "":
            try:
                _thread.start_new_thread(run_calculation_from_ui,
                                  (self.save_file_path,
                                        path, self.file_text, ))
                # run_calculation(self.save_file_path, path)
                # tk._exit(0)
            except Exception as e:
                self.file_text.set(e)

app = VidutinisAkcizas()
app.mainloop()