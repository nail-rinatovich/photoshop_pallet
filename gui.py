#swatch generate  -i rgb.csv -o mypallete.aco
import tkinter as tk
import customtkinter
from tkinter import ttk
from tkinter import colorchooser
import openpyxl
import pandas as pd
import os
from tkinter import filedialog
from tkinter.colorchooser import askcolor



# Список сохраненных цветов
saved_colors = []
def change_color():
    colors = askcolor(title="Tkinter Color Chooser")
    
    
    saved_colors.append((len(set(saved_colors)) + 1, colors[1]))
    
    update_saved_colors()
def find_color():
    color_name = combo.get()
    color_code = find_color_code(color_name)
    
    if color_code is not None:
        
        saved_colors.append((len(set(saved_colors)) + 1, color_code))
        update_saved_colors()
        

def find_color_code(color_name):
    workbook = openpyxl.load_workbook('C:/Users/admin/Desktop/Палитра кодов.xlsx')
    worksheet = workbook.active
    
    for row in worksheet.iter_rows(values_only=True):
        if row[0] == color_name or f'{row[0]}\n' == color_name:
                   
            print("saved_colors: ", saved_colors)
            print('row[0]: ', row[0])
            new_format = []
            for _ in list(row[1:4]):
                new_format.append(round(_))
            return rgb_to_hex(tuple(new_format))  # Возвращаем код цвета

    return None  # Если цвет не найден

def update_saved_colors():
    saved_colors_new = list(set(saved_colors))
    
    temp = [] 
    for x in saved_colors_new:
        if x[1] not in temp:
            print("x: ", x[1])
            temp.append(x)
    saved_colors_new = temp
    print(saved_colors_new)
    saved_colors_new.sort()
    colors_listbox.delete(0, tk.END)

    for color in saved_colors_new:
        colors_listbox.insert(tk.END, f'{color[0]}: {color[1]}')
        
        print("saved_colors_new: ",saved_colors_new)
        
        
    
def rgb_to_hex(rgb):
    
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
        new_string = str(color_name) + ', 0, ' + color_code
        palette.append(new_string)
        print(palette)
    with open(output_file, 'w') as f:
        f.write('\n'.join(palette))
    file_path = filedialog.asksaveasfilename(defaultextension='.ACO',
        initialfile='mypalette.ACO',
        filetypes=(('ACO files', '*.ACO'), ('All files', '*.*')))
    
    print(f'Палитра успешно сохранена в файл: {output_file}')
    os.system(f"swatch generate  -i {output_file} -o {file_path}")
def delete_selected_item():
    selected_index = colors_listbox.curselection()
    if selected_index:
        print(saved_colors)
        colors_listbox.delete(selected_index)
        print((selected_index))
        print((selected_index[0]))
        saved_colors.remove(saved_colors[selected_index[0]])
        print(saved_colors)
       
# Инициализация окна



app = customtkinter.CTk()
app.geometry("600x600")
app.title("Поиск и сохранение цветов")




customtkinter.set_appearance_mode("dark")
# Создание и расположение элементов


label = customtkinter.CTkLabel(app, text='Выберите цвет:')
label.pack(pady=20, padx=10)

combo = customtkinter.CTkEntry(app, width=150)
combo.pack(pady=20, padx=10)

find_button = customtkinter.CTkButton(app, text='Найти в таблице', command=find_color)
find_button.pack(pady=20, padx=10)

delete_button = customtkinter.CTkButton(app, text="Удалить", command=delete_selected_item)
delete_button.pack(pady=20, padx=10)

select_button = customtkinter.CTkButton(app, text='Выбрать цвет вручную', command=change_color)
select_button.pack(pady=20, padx=10)

colors_listbox = tk.Listbox(app, width=40, font=20)
colors_listbox.pack(padx=10)

save_frame = tk.Frame(app)
save_frame.pack()



save_button = customtkinter.CTkButton(app, text='Сохранить палитру', command=save_palette)
save_button.pack(pady=10)





app.configure(bg='blue')
app.mainloop()

