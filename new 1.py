from tkinter import *  
from tkinter.ttk import Combobox  
  
  
window = Tk()  
window.title("Добро пожаловать в приложение PythonRu")  
window.geometry('400x250')  
combo = Combobox(window)  
combo['values'] = (1, 2, 3, 4, 5, "Текст")  
combo.current(1)  # установите вариант по умолчанию  
combo.grid(column=0, row=0)  
window.mainloop()