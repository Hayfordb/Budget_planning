import tkinter as tk
from tkinter import filedialog
import pandas as pd

def transform_data(input_file_path, output_file_path):
    # Чтение данных
    data = pd.read_excel(input_file_path)
    
    # Список для хранения преобразованных данных
    transformed_data_list = []
    
    # Итерация по всем столбцам с датами, начиная со второго
    for date_column in data.columns[1:]:
        # Создание временного DataFrame для текущей даты
        temp_df = data[['Статья', date_column]].copy()
        temp_df.rename(columns={date_column: 'План'}, inplace=True)
        
        # Преобразование названия столбца даты к формату "МММ.ГГ" и добавление как столбец "Дата"
        date = pd.to_datetime(date_column).strftime('%b.%y').upper()
        temp_df['Дата'] = date
        
        # Добавление преобразованного DataFrame в список
        transformed_data_list.append(temp_df)
    
    # Объединение всех временных DataFrame в один
    transformed_data = pd.concat(transformed_data_list, ignore_index=True)
    
    # Переупорядочивание столбцов
    transformed_data = transformed_data[['Дата', 'Статья', 'План']]
    
    # Сохранение преобразованных данных
    transformed_data.to_excel(output_file_path, index=False)

def select_files_and_transform():
    # Создание окна выбора файла
    root = tk.Tk()
    root.withdraw()  # Не показываем полное окно Tkinter

    # Запрос у пользователя файла для чтения
    input_file_path = filedialog.askopenfilename(title="Выберите исходный файл Excel для преобразования")
    if not input_file_path:
        return
    
    # Запрос у пользователя пути для сохранения результата
    output_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], title="Сохранить файл как")
    if not output_file_path:
        return
    
    # Преобразование данных
    transform_data(input_file_path, output_file_path)
    print(f"Файл успешно сохранен: {output_file_path}")

# Запуск функции с GUI
select_files_and_transform()
