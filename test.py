import tkinter as tk
import tkinter.filedialog as fd
import os
import sys
import PyPDF2


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        btn_file = tk.Button(self, text="Выбрать файл",
                             command=self.choose_file)
        btn_dir = tk.Button(self, text="Выбрать папку",
                             command=self.choose_directory)
        btn_file.pack(padx=60, pady=10)
        btn_dir.pack(padx=60, pady=10)
        btn_result =tk.Button(self, text ="Расчет", command=self.result1)
        btn_result.pack(padx=60, pady=10)
        btn_exit =tk.Button(self, text ="Выход", command=self.onExit)
        btn_exit.pack(padx=60, pady=10)

    def choose_file(self):
        filetypes = (("Pdf файл", "*.pdf"),
                     ("Изображение", "*.jpg *.gif *.png"),
                     ("Любой", "*"))
        filename = fd.askopenfilename(title="Открыть файл", initialdir="/",
                                      filetypes=filetypes)
        if filename:
            print(filename)
            pl = open(filename, 'rb')
            plread = PyPDF2.PdfFileReader(pl)
            print(plread.numPages)



    def choose_directory(self):
        directory = fd.askdirectory(title="Открыть папку", initialdir="/")
        if directory:
            print(directory)
            
    
    def onExit(self):
        global app
        app.quit()
    
    def result1(self):
        pl = open('c:/1.pdf', 'rb')
        plread = PyPDF2.PdfFileReader(pl)
        print(plread.numPages)
    

if __name__ == "__main__":
    app = App()
    app.mainloop()
