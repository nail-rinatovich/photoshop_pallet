#swatch generate  -i rgb.csv -o mypallete.aco
import tkinter as tk
from tkinter import ttk
from tkinter import colorchooser
import openpyxl
import pandas as pd

def find_color():
    color_name = combo.get()
    color_code = find_color_code(color_name)
    
    if color_code is not None:
        saved_colors.append((color_name, color_code))
        update_saved_colors()
        combo.set('')  # Сбросить выбор цвета после добавления в список сохраненных

def find_color_code(color_name):
    workbook = openpyxl.load_workbook('C:/Users/admin/Desktop/Палитра кодов.xlsx')
    worksheet = workbook.active
    
    for row in worksheet.iter_rows(values_only=True):
        if row[0] == color_name:
            return row[1:4]  # Возвращаем код цвета

    return None  # Если цвет не найден

def update_saved_colors():
    colors_listbox.delete(0, tk.END)
    for color in saved_colors:
        colors_listbox.insert(tk.END, f'{color[0]}: {color[1]}')

def save_palette():
    if len(saved_colors) == 0:
        return
    
    # Создаем DataFrame для палитры
    df = pd.DataFrame(saved_colors, columns=['Название цвета', 'Код цвета'])
    
    # Выберите формат палитры: .act или .aco
    file_type = file_type_var.get()
    if file_type == 'ACT':
        output_file = 'C:/Users/admin/Desktop/saved_palette.act'
        df.to_csv(output_file, sep='\t', index=False, header=False)
    elif file_type == 'ACO':
        output_file = 'C:/Users/admin/Desktop/saved_palette.aco'
        palette = []
        for _, row in df.iterrows():
            color_name = str(row['Название цвета'])
            color_code = str(row['Код цвета'])
            palette.append(f'{color_name}\t{color_code}')
        with open(output_file, 'w') as f:
            f.write('\n'.join(palette))
    else:
        return
    
    print(f'Палитра успешно сохранена в файл: {output_file}')

# Инициализация окна
root = tk.Tk()
root.title('Поиск и сохранение цветов')
root.geometry('400x400')

# Создание и расположение элементов
frame = tk.Frame(root)
frame.pack(pady=20)

label = tk.Label(frame, text='Выберите цвет:')
label.grid(row=0, column=0, padx=10)

combo = ttk.Entry(frame, width=20)
combo.grid(row=0, column=1, padx=10)

find_button = tk.Button(frame, text='Найти', command=find_color)
find_button.grid(row=0, column=2, padx=10)

colors_listbox = tk.Listbox(root, width=40)
colors_listbox.pack(pady=10)

save_frame = tk.Frame(root)
save_frame.pack()

file_type_var = tk.StringVar(value='ACO')  # Значение по умолчанию для выбора типа файла

act_radio = tk.Radiobutton(save_frame, text='.ACT', variable=file_type_var, value='ACT')
act_radio.grid(row=0, column=0, padx=5)

aco_radio = tk.Radiobutton(save_frame, text='.ACO', variable=file_type_var, value='ACO')
aco_radio.grid(row=0, column=1, padx=5)

save_button = tk.Button(root, height=25, text='Сохранить палитру', command=save_palette)
save_button.pack(pady=10)

# Список сохраненных цветов
saved_colors = []

root.mainloop()

