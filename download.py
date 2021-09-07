import requests #импортируем модуль
f=open(r'D:\file_bdseo.zip',"wb") #открываем файл для записи, в режиме wb
ufr = requests.get("http://site.ru/file.zip") #делаем запрос
f.write(ufr.content) #записываем содержимое в файл; как видите - content запроса
f.close()
1