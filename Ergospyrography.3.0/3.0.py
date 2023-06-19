import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import os
import pandas as pd
import numpy as np
import openpyxl
import os
from openpyxl.styles import Alignment
import time

file_name = ""

def start_program():
    global file_name
    # Загрузка данных из первого листа файла Excel
    df_first_sheet = pd.read_excel(file_name, sheet_name=0)
    # Получение значения из ячейки в шестой строке и втором столбце
    value = df_first_sheet.iloc[4, 1]
    # Преобразование значения в строку и разделение на две части
    value1, value2 = map(int, str(value).split()[0].split('-'))
    # Создание диапазона значений
    my_range = range(value1, value2 + 1)
    max_value = max(my_range)

    # Загрузка данных из первого листа файла Excel
    df_first_sheet = pd.read_excel(file_name, sheet_name=0)
    # Получение значения из ячейки в седьмой строке и втором столбце
    value1 = df_first_sheet.iloc[5, 1]
    # Преобразование значения в строку и разделение на две части
    value_1, value_2 = map(int, str(value1).split()[0].split('-'))
    # Создание диапазона значений
    my_range1 = range(value_1, value_2 + 1)
    max_value = max(my_range1)
    # Загрузка данных из файла Excel
    xlsx = pd.read_excel(file_name, sheet_name=None, skiprows=1)
    # Получение списка имен листов в файле
    sheet_names = list(xlsx.keys())
    # Выбор второго листа
    second_sheet_name = sheet_names[1]
    # Загрузка данных из второго листа
    df = xlsx[second_sheet_name]

    # Выбор столбца 'Скорость'
    speed_column = df['Скорость']
    # Поиск первого числового значения в столбце 'Скорость'
    first_numeric_value = \
    speed_column[speed_column.apply(lambda x: isinstance(x, (int, float)) and not np.isnan(x))].iloc[0]
    # Преобразование столбца 'Время' в формат datetime
    df['Время'] = pd.to_datetime(df['Время'].astype(str))

    # Выбор столбца 'Время'
    time_column = df['Время']
    # Поиск значения времени, соответствующего первому числовому значению в столбце 'Скорость'
    time_value = time_column.loc[speed_column == first_numeric_value].iloc[0]
    # Получение индекса строки с найденным значением времени
    time_index = time_column[time_column == time_value].index[0]
    # Вычисление времени, которое на 30 секунд раньше найденного значения времени
    time_30_seconds_ago = time_value - pd.Timedelta(seconds=30)
    # Поиск индекса строки, значение времени в которой ближе всего к вычисленному времени
    closest_time_index = (time_column - time_30_seconds_ago).abs().idxmin()
    # Выбор значений в найденной строке, начиная с третьего столбца
    result = df.iloc[closest_time_index][2:]

    # Вычисление количества строк для расчета
    start_index = min(time_index, closest_time_index)
    end_index = max(time_index, closest_time_index)
    row_count = end_index - start_index + 1

    # Вычисление средних значений для выбранных строк
    mean_values = df.iloc[start_index:end_index + 1, 2:].mean()

    # Выбор столбца 'ЧСС'
    hr_column = df['ЧСС ']
    # Фильтрация значений в столбце 'ЧСС' по диапазону значений
    filtered_values = hr_column[hr_column.isin(my_range)]
    # Поиск максимального числового значения в отфильтрованных значениях столбца 'ЧСС'
    max_hr_value = filtered_values.max()

    # Фильтрация значений в столбце 'ЧСС' по диапазону значений
    filtered_values1 = hr_column[hr_column.isin(my_range1)]
    # Поиск максимального числового значения в отфильтрованных значениях столбца 'ЧСС'
    max_hr_value1 = filtered_values1.max()

    ###############################
    # Поиск значения времени, соответствующего первому числовому значению в столбце 'ЧСС'
    time_value = time_column.loc[hr_column == max_hr_value].iloc[0]
    # Получение индекса строки с найденным значением времени
    time_index = time_column[time_column == time_value].index[0]
    # Вычисление времени, которое на 30 секунд раньше найденного значения времени
    time_30_seconds_ago = time_value - pd.Timedelta(seconds=30)
    # Поиск индекса строки, значение времени в которой ближе всего к вычисленному времени
    closest_time_index = (time_column - time_30_seconds_ago).abs().idxmin()
    # Выбор значений в найденной строке, начиная с третьего столбца
    result = df.iloc[closest_time_index][2:]

    # Вычисление количества строк для расчета
    start_index = min(time_index, closest_time_index)
    end_index = max(time_index, closest_time_index)
    row_count = end_index - start_index + 1

    # Вычисление средних значений для выбранных строк
    mean_values1 = df.iloc[start_index:end_index + 1, 2:].mean()

    # Поиск значения времени, соответствующего первому числовому значению в столбце 'ЧСС'
    time_value1 = time_column.loc[hr_column == max_hr_value1].iloc[0]
    # Получение индекса строки с найденным значением времени
    time_index = time_column[time_column == time_value1].index[0]
    # Вычисление времени, которое на 30 секунд раньше найденного значения времени
    time_30_seconds_ago = time_value1 - pd.Timedelta(seconds=30)
    # Поиск индекса строки, значение времени в которой ближе всего к вычисленному времени
    closest_time_index = (time_column - time_30_seconds_ago).abs().idxmin()
    # Выбор значений в найденной строке, начиная с третьего столбца
    result = df.iloc[closest_time_index][2:]

    ############################
    # Вычисление количества строк для расчета
    start_index = min(time_index, closest_time_index)
    end_index = max(time_index, closest_time_index)
    row_count = end_index - start_index + 1

    # Вычисление средних значений для выбранных строк
    mean_values2 = df.iloc[start_index:end_index + 1, 2:].mean()

    #########################3
    # Поиск последнего числового значения в столбце 'ЧСС'
    last_hr_value = hr_column[hr_column.apply(lambda x: isinstance(x, (int, float)) and not np.isnan(x))].iloc[-1]
    # Вывод результата на экран

    # Поиск значения времени, соответствующего последнему числовому значению в столбце 'ЧСС'
    time_value = time_column.loc[hr_column == last_hr_value].iloc[0]
    # Получение индекса строки с найденным значением времени
    time_index = time_column[time_column == time_value].index[0]
    # Вычисление времени, которое на 30 секунд раньше найденного значения времени
    time_30_seconds_ago = time_value - pd.Timedelta(seconds=30)
    # Поиск индекса строки, значение времени в которой ближе всего к вычисленному времени
    closest_time_index = (time_column - time_30_seconds_ago).abs().idxmin()
    # Выбор значений в найденной строке, начиная с третьего столбца
    result = df.iloc[closest_time_index][2:]

    # Вычисление количества строк для расчета
    start_index = min(time_index, closest_time_index)
    end_index = max(time_index, closest_time_index)
    row_count = end_index - start_index + 1
    print(f"Количество строк для расчета: {row_count}")

    # Вычисление средних значений для выбранных строк
    mean_values3 = df.iloc[start_index:end_index + 1, 2:].mean()
    # Вывод средних значений на экран с названиями показателей

    #####################

    # Создание списка названий колонок
    columns = ['',
               'ЧСС ',
               'Скорость',
               'Объем вдоха (л)',
               'Объем выдоха (л)',
               'частота дыхания',
               'FiO2 вдыхаемый кислород',
               'FetO2 выдыхаемый кислород',
               'FiCO2',
               'FetCO2  % утилизации О2',
               'Vвдоха без МП (л)',
               'Vвыдоха без МП (л)',
               'VO2 (л)',
               'VCO2 (л)',
               'VO2 (мл/кг/мин.)',
               'VCO2 (мл/кг/мин.)',
               'VO2 (л/мин.)',
               'VCO2 (л/мин.)',
               'ВЭКВ O2',
               'ВЭКВ CO2',
               'Дыхательный коэффициент',
               'Метаболизм по O2 (ккал/сут.)',
               'Метаболизм по CO2 (ккал/сут.)',
               'ДМП',
               'Минутн.вентиляция.(л)']

    # Создание датафрейма с указанными колонками
    df = pd.DataFrame(columns=columns)

    # Добавление строк в датафрейм
    df.loc[0] = ['ДО ТЕСТА'] + [None] * (len(columns) - 1)
    df.loc[1] = ['ПАНО'] + [None] * (len(columns) - 1)
    df.loc[2] = ['МПК'] + [None] * (len(columns) - 1)
    df.loc[3] = ['ПЛАТО'] + [None] * (len(columns) - 1)

    ######################
    # Удаление пробелов с начала и конца названий столбцов в датафрейме
    df.columns = df.columns.str.strip()

    # Удаление пробелов с начала и конца названий столбцов в mean_values
    mean_values.index = mean_values.index.str.strip()
    mean_values1.index = mean_values.index.str.strip()
    mean_values2.index = mean_values.index.str.strip()
    mean_values3.index = mean_values.index.str.strip()
    #################
    # Добавление средних значений в датафрейм
    df.iloc[0, 1:] = mean_values.values
    df.iloc[1, 1:] = mean_values1.values
    df.iloc[2, 1:] = mean_values2.values
    df.iloc[3, 1:] = mean_values3.values
    #####################

    end_time = time.time()

    # Показываем сообщение о завершении работы программы
    messagebox.showinfo("Работа завершена", "Программа выполнена успешно!")

    # Запрос имени нового файла у пользователя
    filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])

    if filename:


        # Получение пути к папке сохранения
        folder_path = os.path.dirname(filename)

        # Создание полного пути к новому файлу
        new_file_path = os.path.join(folder_path, os.path.basename(filename))


        # Сохранение датафрейма в файл Excel
        df.to_excel(new_file_path, index=False)

        # Показываем сообщение о завершении работы программы
        messagebox.showinfo("Файл сохранен", f"Файл сохранен по следующему пути: {filename} ")


root = tk.Tk()
root.title("Ergospyrography")
root.geometry("500x300")
root.configure(bg="#E8EAF6")

def choose_file():
    global file_name
    file_path = filedialog.askopenfilename()
    file_name = file_path
    file_path_entry.delete(0, tk.END)
    file_path_entry.insert(0, file_path)

choose_file_button = tk.Button(root, text="Выбрать файл", command=choose_file, bg="#7986CB", fg="white", bd=0, padx=10, pady=5)
choose_file_button.pack(pady=10)

file_path_entry = tk.Entry(root, width=50, bd=1, relief="solid")
file_path_entry.pack()

start_button = tk.Button(root, text="Старт", command=start_program, bg="#4CAF50", fg="white", bd=0, padx=10, pady=5)
start_button.pack(pady=10)

root.mainloop()
