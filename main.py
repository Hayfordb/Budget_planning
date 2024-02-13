import os
import tkinter as tk
from tkinter import filedialog, scrolledtext
import openpyxl
from openpyxl.utils.exceptions import InvalidFileException

def check_excel_parameters(file):
    """
    Проверяет, содержит ли лист "Параметры" в файле Excel непустые значения в диапазоне C7:C10.
    Возвращает True, если условия выполнены, иначе False.
    """
    try:
        workbook = openpyxl.load_workbook(file, data_only=True)
        if "Параметры" in workbook.sheetnames:
            sheet = workbook["Параметры"]
            for row in sheet['C7:C10']:
                for cell in row:
                    if cell.value is None or cell.value == "":
                        return False
            return True
        else:
            return False
    except InvalidFileException:
        return False  # В случае, если файл не может быть открыт как Excel файл.

def append_text_with_color(text_widget, text, color):
    """
    Добавляет текст в виджет Text с заданным цветом.
    """
    text_widget.tag_config(color, foreground=color)
    text_widget.insert(tk.END, text, color)

def get_excel_files(directory):
    """
    Возвращает список путей к Excel файлам в указанной директории.
    """
    excel_files = [os.path.join(directory, f) for f in os.listdir(directory) if f.endswith('.xlsx')]
    return excel_files

def browse_directory():
    """
    Позволяет пользователю выбрать директорию и проверяет Excel файлы перед отображением.
    """
    directory = filedialog.askdirectory()
    if directory:  # Если пользователь выбрал директорию
        excel_files = get_excel_files(directory)
        text_area.delete('1.0', tk.END)  # Очистить текстовое поле
        if excel_files:
            for file in excel_files:
                if check_excel_parameters(file):
                    append_text_with_color(text_area, file + '\n', "black")
                else:
                    append_text_with_color(text_area, file + '\n', "red")
        else:
            append_text_with_color(text_area, "Excel файлы не найдены в указанной папке.\n", "black")

# Создание главного окна
root = tk.Tk()
root.title("Поиск Excel файлов")
root.geometry("600x400")

# Создание виджетов
browse_button = tk.Button(root, text="Выбрать папку", command=browse_directory)
text_area = scrolledtext.ScrolledText(root, wrap=tk.WORD)
browse_button.pack(pady=10)
text_area.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

# Запуск главного цикла обработки событий
root.mainloop()
