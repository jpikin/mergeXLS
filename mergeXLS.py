import os
from tkinter import Tk, filedialog, Button, Label, Text, Scrollbar
import pandas as pd
from openpyxl.workbook import Workbook


# Функция для объединения файлов Excel
def merge_excel_files():
    order = ""

    # Выбор нескольких файлов
    file_paths = filedialog.askopenfilenames(
        title="Выберите файлы Excel",
        filetypes=(("Excel files", "*.xls*"), ("All Files", "*.*"))
    )

    if not file_paths:
        return

    try:
        df_list = []

        # Проход по каждому файлу
        for path in file_paths:
            df = pd.read_excel(path)

            # Поиск строки с "Заказ"
            for idx, row in df.iterrows():
                values = row.to_list()
                if any(cell == 'Заказ' for cell in row):
                    pos = values.index('Заказ')
                    if len(values) > pos + 1:
                        order = values[pos + 1]
                    else:
                        order = None
                    break

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
            required_columns = ["Наименование материала", "Количество в заказе", "Ед. изм."]
            missing_cols = set(required_columns) - set(headers)
            if missing_cols:
                raise Exception(f"В файле '{path}' не хватает колонок: {missing_cols}")

            # Начинаем брать данные со следующей строки после шапки
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
        grouped_df = merged_df.groupby(['Наименование материала', 'Ед. изм.'])[
            'Количество в заказе'].sum().reset_index()

        # Путь для сохранения итогового файла
        output_path = os.path.join(os.getcwd(), 'Заказ_'+order+'.xlsx')

        if output_path:
            # Создаем новый dataframe для записи заказа
            info_df = pd.DataFrame({'Заказ': [''], order: ['']}, index=[0])

            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                # Экспортим информацию о заказе
                info_df.to_excel(writer, sheet_name='Итоги', index=False, startrow=0)

                # Переносимся ниже и выводим итоговые материалы
                grouped_df.to_excel(writer, sheet_name='Итоги', index=False, startrow=len(info_df)+2)

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


root = Tk()
root.title("Объединение Excel-файлов")

select_button = Button(root, text="Выбрать файлы", command=merge_excel_files)
select_button.pack(pady=10)

result_label = Label(root, text="Результат:")
result_label.pack()

scrollbar = Scrollbar(root)
scrollbar.pack(side='right', fill='y')

result_text = Text(root, yscrollcommand=scrollbar.set, height=10, width=80)
result_text.pack(fill='both', expand=True)
scrollbar.config(command=result_text.yview)

root.mainloop()