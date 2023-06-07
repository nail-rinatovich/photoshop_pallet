#swatch generate  -i rgb.csv -o mypallete.aco
import tkinter as tk
from tkinter import ttk
from tkinter import colorchooser
import openpyxl
import pandas as pd
import os
from tkinter import filedialog
    
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
            new_format = []
            for _ in list(row[1:4]):
                new_format.append(round(_))
                print(new_format)
            return rgb_to_hex(tuple(new_format))  # Возвращаем код цвета

    return None  # Если цвет не найден

def update_saved_colors():
    colors_listbox.delete(0, tk.END)
    for color in saved_colors:
        colors_listbox.insert(tk.END, fr'{color[0]}: {color[1]}')
def rgb_to_hex(rgb):
    #return '%02x%02x%02x' % rgb
    hex_value = "{:02X}{:02X}{:02X}".format(rgb[0], rgb[1], rgb[2])
    return hex_value
def save_palette():
    if len(saved_colors) == 0:
        return
    
    # Создаем DataFrame для палитры
    df = pd.DataFrame(saved_colors, columns=['Название цвета', 'Код цвета'])
    
    output_file = 'C:/Users/admin/Desktop/saved_palette.csv'
    palette = []
    
    palette.append('name,space_id,color')
    for _, row in df.iterrows():
        color_name = row['Название цвета']
        color_code = '#' + row['Код цвета']
        new_string = color_name + ', 0, ' + color_code
        palette.append(new_string)
        print(palette)
    with open(output_file, 'w') as f:
        f.write('\n'.join(palette))
    file_path = filedialog.asksaveasfilename(defaultextension='.ACO',
        initialfile='mypalette.ACO',
        filetypes=(('ACO files', '*.ACO'), ('All files', '*.*')))
    
    print(f'Палитра успешно сохранена в файл: {output_file}')
    os.system(f"swatch generate  -i {output_file} -o {file_path}")
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



save_button = tk.Button(root, height=25, text='Сохранить палитру', command=save_palette)
save_button.pack(pady=10)



# Список сохраненных цветов
saved_colors = []

root.mainloop()

