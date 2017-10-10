import tkinter as tk
from tkinter.filedialog import askopenfilename, asksaveasfilename
from vidutinis_akcizas import run_calculation


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
        tk.Frame.__init__(self, parent)
        label = tk.Label(self, text="Akcizo skaičiavimas", font=LARGE_FONT)
        label.pack(pady=10, padx=10)

        browse_btn = tk.Button(self, text="Pasirinkti",
                          command=self.browse_func)
        browse_btn.pack()



    def browse_func(self):
        filename = askopenfilename()
        self.save_file_path = filename

        save_btn = tk.Button(self, text="Išsaugoti",
                             command=self.save_func)
        save_btn.pack()

    def save_func(self):
        path = asksaveasfilename()
        run_calculation(self.save_file_path, path)
        tk._exit(0)

app = VidutinisAkcizas()
app.mainloop()