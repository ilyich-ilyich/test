import pandas as pd
import tkinter as tk
import tkinter.filedialog as fd
import os
import sys
#import PyPDF2
#import fitz
#import pytesseract
#from PIL import Image

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        btn_file = tk.Button(self, text="Выбрать файл",
                             command=self.choose_file)
        
        btn_file.pack(padx=60, pady=10)
        btn_exit =tk.Button(self, text ="Выход", command=self.onExit)
        btn_exit.pack(padx=60, pady=10)
   
    def choose_file(self):
        filetypes = (("файл Exel", "*.xlsx"),
                     ("Pdf файл", "*.pdf"),
                     ("Любой", "*"))
        filename = fd.askopenfilename(title="Открыть файл", initialdir="/",filetypes=filetypes)
        if filename:
            print(filename)
            #text = " test"
            #print(text)
            #with open(f'{file_name}.txt', 'w') as text_file:
            #    text_file.write(text)    
            cols = [ 8]
            top_players = pd.read_excel(filename, header = 2, dtype={ 'Cсылка на акт':str}, usecols=cols)
            #top_players.head()
            print(top_players)
            #with open(f'{filename}.txt', 'w') as text_file:
             #   text_file.write(str(top_players)) 

    def onExit(self):
        global app
        app.quit()

if __name__ == "__main__":
    app = App()
    app.mainloop()
