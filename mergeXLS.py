import os
from tkinter import Tk, filedialog, Button, Label, Text, Scrollbar
import pandas as pd


# Функция для объединения файлов Excel
def merge_excel_files():
    # Выбор нескольких файлов
    file_paths = filedialog.askopenfilenames(
        title="Выберите файлы Excel",
        filetypes=(("Excel files", "*.xlsx"), ("All Files", "*.*"))
    )

    if not file_paths:
        return

    try:
        df_list = []

        # Проход по каждому файлу
        for path in file_paths:
            df = pd.read_excel(path)

            # Поиск строки с "Артикул"
            start_row_idx = None
            for idx, row in df.iterrows():
                if any(cell == 'Артикул' for cell in row):
                    start_row_idx = idx
                    break

            if start_row_idx is None:
                raise Exception(f"В файле '{path}' не найдено ключевое слово 'Артикул'.")

            # Зацепились за следующую строку после "Артикул" как шапку таблицы
            headers = df.iloc[start_row_idx].values.tolist()

            # Проверка наличия нужных колонок
            required_columns = ["Наименование материала", "Количество в заказе"]
            missing_cols = set(required_columns) - set(headers)
            if missing_cols:
                raise Exception(f"В файле '{path}' не хватает колонок: {missing_cols}")

            # Начинаем брать данные с следующей строки после шапки
            useful_data = df.iloc[start_row_idx + 1:]

            # Переименовываем колонки согласно заголовкам
            useful_data.columns = headers

            # Оставляем только нужные колонки
            useful_data = useful_data[required_columns]

            # Исключаем строки с пустыми значениями
            useful_data.dropna(how='any', subset=required_columns, inplace=True)

            df_list.append(useful_data)

        # Объединение всех собранных таблиц
        merged_df = pd.concat(df_list)

        # Суммирование по материалам
        grouped_df = merged_df.groupby('Наименование материала')['Количество в заказе'].sum().reset_index()

        # Путь для сохранения итогового файла
        output_path = filedialog.asksaveasfilename(
            defaultextension='.xlsx',
            title="Сохранить объединённый файл"
        )

        if output_path:
            grouped_df.to_excel(output_path, index=False)
            result_text.delete('1.0', 'end')
            result_text.insert('end', f'Файл успешно сохранён в {output_path}')
        else:
            raise Exception('Отмена сохранения файла.')

    except FileNotFoundError:
        result_text.delete('1.0', 'end')
        result_text.insert('end', 'Ошибка: Один или несколько файлов не найдены.')
    except KeyError:
        result_text.delete('1.0', 'end')
        result_text.insert('end', 'Ошибка: Неверные названия колонок в файлах.')
    except Exception as e:
        result_text.delete('1.0', 'end')
        result_text.insert('end', f'Ошибка: {e}')


# Создание окна приложения
root = Tk()
root.title("Объединение Excel-файлов")

# Кнопка выбора файлов
select_button = Button(root, text="Выбрать файлы", command=merge_excel_files)
select_button.pack(pady=10)

# Поле вывода результата
result_label = Label(root, text="Результат:")
result_label.pack()

scrollbar = Scrollbar(root)
scrollbar.pack(side='right', fill='y')

result_text = Text(root, yscrollcommand=scrollbar.set, height=10, width=80)
result_text.pack(fill='both', expand=True)
scrollbar.config(command=result_text.yview)

# Запуск главного цикла
root.mainloop()