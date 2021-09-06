import tkinter as tk
import tkinter.filedialog as fd
import os
import sys
from pyPdf import PdfFileReader

PDF_EXTENSION = '.pdf'

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
        filetypes = (("Текстовый файл", "*.txt"),
                     ("Изображение", "*.jpg *.gif *.png"),
                     ("Любой", "*"))
        filename = fd.askopenfilename(title="Открыть файл", initialdir="/",
                                      filetypes=filetypes)
        if filename:
            print(filename)

    def choose_directory(self):
        directory = fd.askdirectory(title="Открыть папку", initialdir="/")
        if directory:
            print(directory)
    
    def onExit(self):
        global app
        app.quit()
    
    def result1(self):
         total_pages_count = 0
        for root, dir, files in  os.walk(directory):
            for file_name in files:
                if file_name[-len(PDF_EXTENSION):] == PDF_EXTENSION:
                    file_path = os.path.join(root, file_name)
                    file_pages_count = pages_count(file_path)
                    print file_path, file_pages_count
                total_pages_count += file_pages_count
        print 'total:', total_pages_count
 
    def pages_count(path):
        return PdfFileReader(file(path, "rb")).getNumPages()
        a=PdfFileReader(file(path, "rb", FloatingPointError)
           

if __name__ == "__main__":
    app = App()
    app.mainloop()
