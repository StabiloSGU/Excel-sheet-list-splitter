import pandas as pd
import os

def divider():
    print()
    print("----------------------------------")
    print()

def filelist():
    i = 1
    for filename in os.listdir():
        if filename.endswith(".xlsx"):
            files.append(filename)
            print(str(i)+") "+str(filename))
            i += 1
def fieldlist(df):
    i = 1
    for key in df.keys():
        fields.append(key)
        print(str(i)+") "+str(key))
        i += 1

        
def uniques(df, field):
    allRows = [row for row in df[fields[field-1]]]
    for line in allRows:
        if line not in rows:
            rows.append(line)

def splitBySheets():
    writer = pd.ExcelWriter("Разбивка по листам "+str(files[file-1]))
    for sheetname in rows:
        frame = datafile.query(str(fields[field-1])+'== "'+str(sheetname)+'"')
        sht = str(sheetname).replace("/"," - ")
        frame.to_excel(writer, sht[:31], index=False)
    
    writer.save()

def splitByFiles():
    for sheetname in rows:
        writer = pd.ExcelWriter("Разбивка по файлам "+str(sheetname)+" "+str(files[file-1]))
        frame = datafile.query(str(fields[field-1])+'== "'+str(sheetname)+'"')
        sht = str(sheetname).replace("/"," - ")
        frame.to_excel(writer, sht[:31], index=False)
    
    writer.save()
    
            
def split(mode):
    if mode == 0:
        splitBySheets()
    if mode == 1:
        splitByFiles()
    
cont = 1
while cont == 1:
    #Вывести список файлов, выбор файла
    files = []
    filelist()
    file = int(input("Выберите файл: "))
    datafile = pd.read_excel(files[file-1])
    divider()

    #выбор поля
    print("Поля таблицы:")
    fields = []
    fieldlist(datafile)
    field = int(input("Выберите поле: "))
    divider()

    #Выбор разделения на файлы\страницы
    splitmode = int(input("Разбить по листам/файлам (0 - листы, 1 - файлы): "))
    divider()

    #Логика программы
        #1) Получить список уникальный значений по которым сортировать выбранное поле
    rows = []
    uniques(datafile, field)
    #rows.sort()

        #2) Запуск разделителя

    split(splitmode)
    print("Готово.")
    print()
    
    #запрос на продолжение работы
    cont = int(input("Разделить еще файл? (0 - нет, 1 - да): "))


