import tkinter as tk
import tkinter.filedialog as fd
import os
import sys
#import PyPDF2
import fitz
import pytesseract
from PIL import Image


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        btn_file = tk.Button(self, text="Выбрать файл",
                             command=self.choose_file)
        btn_file.pack(padx=60, pady=10)
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
            # pl = open(filename, 'rb')
            #plread = PyPDF2.PdfFileReader(pl)
            #print(plread.numPages)
            pdf_file = fitz.open(filename)
            #location = input("Enter the location to save: ")
            location = 'c:\img' 
            #поиск количества страниц в pdf
            number_of_pages = len(pdf_file)
            print('Количество страниц=',number_of_pages)
            #итерация по каждой странице в pdf
            for current_page_index in range(number_of_pages):
            # итерация по каждому изображению на каждой странице PDF
                for img_index,img in enumerate(pdf_file.getPageImageList(current_page_index)):
                    xref = img[0]
                    image = fitz.Pixmap(pdf_file, xref)
                    # если это чёрно-белое или цветное изображение
                    img_filename = "{}/image{}-{}.png".format(location,current_page_index, img_index)   
                    if image.n < 5:
                             
                        image.writePNG(img_filename)
                        #если это CMYK: конвертируем в RGB
                    else:                
                        new_image = fitz.Pixmap(fitz.csRGB, image)
                        new_image.writePNG(img_filename)
                    img = Image.open(img_filename)
                    file_name = img.filename
                    file_name = file_name.split(".")[0]
                    text = pytesseract.image_to_string(img, lang='rus').strip()
                    print(text)
                    with open(f'{file_name}.txt', 'w') as text_file:
                        text_file.write(text)    



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
