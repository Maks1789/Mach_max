import os, openpyxl

import openpyxl
from openpyxl import load_workbook

location_file = (os.listdir(os.getcwd()))
all_find_file_text = []
all_find_file_exel = []


# находим все файли ексель, сохраняя их в - all_find_file_exel, и текстовые файлы в = all_find_file_text
def File_filter(location_file):
    try:
        for file in location_file:
            if file.endswith(".txt"):
                all_find_file_text.append(str(file))
            elif file.endswith(".xlsx"):
                all_find_file_exel.append(str(file))

    except Exception:
        print("Нет таких файлов")


File_filter(location_file)

print(f" Файлы для проверки совпадений {all_find_file_exel}")



# загружает файлы ексель и открывает рабочую страницк
def spisok(files):
    a = openpyxl.load_workbook(f"{files}")
    a = a.active.values
    return a




# добавляем в список данные из генератора, но только открытого файла
def save_value(data):
    line = []
    for col in data:
        for row in col:
            line.append(row)
    return line




# Сравнение двух файлов ексель
def mach_point(data_1, data_2):
    data_mach = []
    for i in data_1:
        for t in data_2:
            if i == t:
                data_mach.append(i)
    return data_mach



dat_1 = save_value(spisok(all_find_file_exel[0]))
dat_2 = save_value(spisok(all_find_file_exel[1]))
print(f"В указаных файлах есть такие совпадения {mach_point(dat_1, dat_2)}")


input("Press ENTER to exit")