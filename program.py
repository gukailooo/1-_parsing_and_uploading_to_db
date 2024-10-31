import pandas as pd
import os
import re
import csv
import pandas as pd
from sqlalchemy import create_engine

class DataParser:
    def __init__(self, categories: dict):
        self.categories = categories

    def parsing_xls(self, file_path: str, info_start: int, info_end: int, final_file_path: str):
        document_number = re.search(r'счет (\d+)', os.path.basename(file_path))
        document_number = document_number.group(1) if document_number else None

        df = pd.read_excel(file_path)
        columns = ['Наименование объекта нефинансового актива', 
                   'Номер (код) объекта учета (инвентарный или иной)', 
                   'Единица измерения', 
                   'цена (оценочная стоимость), руб',
                   'количество',
                   'сумма, руб.',
                   'номер (код) счета',
                   'Примечание']
        
        number = [3, 8, 11, 14, 17, 22, 33, 47]
        objects = []

        for string in range(info_start - 2, info_end - 1):
            total_names = []
            for column in number:
                value = df.iloc[string, column]
                total_names.append(str(value).strip())
            objects.append(total_names)

        data = pd.DataFrame(objects, columns=columns)
        data['Номер документа'] = document_number
        
        pattern = r',?\s*(\d{4})\b'
        data['Дата'] = data['Примечание'].apply(
            lambda x: int(re.search(pattern, str(x)).group(1)) if re.search(pattern, str(x)) else None
        )
        data['Дата'] = data['Дата'].apply(
            lambda x: x if x is None or (1900 <= x <= 2100) else None
        )

        data['Примечание'] = data['Примечание'].apply(lambda x: re.sub(pattern, '', str(x)))
        data['Дата'] = data['Дата'].astype('Int64')

        if not os.path.exists(final_file_path):
            with open(final_file_path, 'w', newline='') as f:
                csv_string = data.to_csv(sep='|', index=False, header=True)
                f.write(csv_string.replace('"', ''))
        else:
            with open(final_file_path, 'a', newline='') as f:
                csv_string = data.to_csv(sep='|', index=False, header=False)
                f.write(csv_string.replace('"', ''))

    def clean_and_save_csv(self, input_file_path: str):
        df = pd.read_csv(input_file_path, encoding='cp1251', sep='|')
        df['Cтатус'] = None
        df['Закреплен'] = None
        df['Категория'] = None
        df.to_csv(input_file_path, index=False, encoding='cp1251', sep='|')

        with open(input_file_path, 'r', encoding='cp1251') as infile:
            reader = csv.reader(infile)
            cleaned_data = []
            for row in reader:
                cleaned_row = [cell.replace('"', '') for cell in row]
                cleaned_data.append(cleaned_row)

        header = cleaned_data[0]
        cleaned_data[1:] = sorted(cleaned_data[1:], key=lambda x: x[0].lower())

        with open(input_file_path, 'w', newline='', encoding='cp1251') as outfile:
            writer = csv.writer(outfile)
            writer.writerows(cleaned_data)

    def assign_unique_numbers(self, input_csv, output_csv):
        df = pd.read_csv(input_csv, encoding='cp1251', sep='|')
        df.insert(0, 'Уникальный номер', '')

        category_counters = {}
        unique_numbers_and_categories = []
        unique_number = 1

        # получение уникального номера
        def get_unique_number(category):
            nonlocal unique_number
            if category not in category_counters:
                category_counters[category] = f'{unique_number:04}'
                unique_number += 1
            return category_counters[category]

        # присвоение уникального номера объектам
        for index, row in df.iterrows():
            object_name = row['Наименование объекта нефинансового актива'].lower()
            
            found_category = None
            for category, items in self.categories.items():
                for item in items:
                    if item.lower() in object_name:
                        found_category = category
                        break
                if found_category:
                    break
            
            if found_category:
                base_unique_number = get_unique_number(found_category)
                
                # подсчет количества объектов в данной категории
                count_in_category = sum(1 for entry in unique_numbers_and_categories if entry['Категория'] == found_category)
                
                # уникальный номер с порядковым номером
                unique_number_value = f'{base_unique_number}-{count_in_category + 1}'
                
                # уникальный номер в столбец 'Уникальный номер'
                df.at[index, 'Уникальный номер'] = unique_number_value
                            
                # добавление категории в столбец 'Категория'
                df.at[index, 'Категория'] = found_category
                
                # уникальный номер и категория в список
                unique_numbers_and_categories.append({'Уникальный номер': unique_number_value, 'Категория': found_category})

        df.to_csv(output_csv, index=False, encoding='cp1251', sep='|')

    def save_result(self, df: pd.DataFrame, output_file_path: str, input_file_path: str):
        df.to_csv(output_file_path, index=False, encoding='utf-8', sep='|')
        if os.path.exists(input_file_path):
            os.remove(input_file_path)

    def save_to_database(self, df: pd.DataFrame, db_path: str, table_name: str):
        # Создание подключения к базе данных SQLite
        engine = create_engine(f'sqlite:///{db_path}')
        
        # Преобразование и очистка данных перед сохранением
        df['Дата'] = pd.to_numeric(df['Дата'], errors='coerce', downcast='integer')
        df['Дата'] = df['Дата'].fillna('').astype(str)
        df['Дата'] = df['Дата'].str.replace('.0', '', regex=False)

        df['Cтатус'] = df['Cтатус'].astype(str)
        df['Закреплен'] = df['Закреплен'].astype(str)

        df['Номер документа'] = pd.to_numeric(df['Номер документа'], errors='coerce', downcast='integer')
        df['Номер документа'] = df['Номер документа'].fillna(0).astype(int)

        df['количество'] = pd.to_numeric(df['количество'], errors='coerce', downcast='integer')
        df['количество'] = df['количество'].fillna(0).astype(int)

        # Сохранение DataFrame в таблицу SQLite
        df.to_sql(table_name, con=engine, if_exists='replace', index=False)
        print(f"Данные успешно сохранены в таблицу '{table_name}' базы данных '{db_path}'.")

# Пример использования
if __name__ == "__main__":
 
    categories = {'Мебель' : ['Стул', 'Кресло', 'Шкаф', 'Ведро', 'Банкетка', 'Вешалка', 'Стойка', 'Стол', 'Тумба', 'Ф/рамка'], 
              'Накопитель': ['Flash Накопитель', 'USB-Flash', 'Flash память', 'Накопитель', 'Накопител', 
                             'Носители информации', 'Флеш Диск', 'Флеш накопитель', 'Флеш-накопитель',
                             'Модуль памяти', 'Память оперативная', 'USB', 'Внешний жесткий диск', 'Диск'],
              'АП' : ['АП', 'ЗАП'],
              'Аудио и видео оборудование' : ['Аудиосистема', 'Акустическая система', 'Ресивер', 'Усилитель',
                                              'Система видео-отображения', 'Документ-камера'],
              'Бытовая техника': ['Телевизор', 'Пылесос', 'Холодильник', 'Кондиционер', 'Сплит-система'],
              'Документы' : ['Отчет о НИР'],
              'Проекционное оборудование' : ['Комплект проекционного оборудования', 'Проектор', 
                                             'Экран', 'Мобильный презентационный комплект', 'Видеопроектор', 'Презентационный манипулятор'],
              'Компьютерная техника' : ['Компьютер', 'Моноблок', 'Монитор', 'Привод DVD', 'Принтер', 'Плоттер',
                                        'Сканер', 'Видеокарта', 'МФУ', 'Многофункциональное устройство',
                                        'ПЭВМ', 'ПВЭМ', 'Персональная электронная вычислительная машина',
                                        'Ноутбук', 'Комплект ноутбука', 'Цефей-01', 'Устройство лазерного типа'],
              'Офисные принадлежности' : ['Бейдж', 'Бумага цветная', 'Папка адресная А4', 
                                          'Средства скрепления и склеивания', 'Степлер', 'Точилка', 'Ножницы',
                                          'Вертикальный накопитель', 'Брошюровочная машина', 'Информационные карманы'],
              'Офисное оборудование' : ['Бумагоуничтожающая машина', 'Бумагоуничтожительная', 'Уничтожитель бумаги', 
                                        'Уничтожитель документов', 'Бумагорезательная машина', 'Уничтожители документов',
                                        'Уничтожитель'],
              'Рабочее место' : ['АРМ', 'Автоматизированное рабочее место', 'Рабочее место преподавателя'],
              'Расходные материалы' : ['Блок фотобарабана', 'Структурированная кабельная система', 
                                       'Картридж', 'Комплект картриджей', 'Тонер-картридж'],
              'Специальное оборудование' : ['Защищенное рабочее место', 'Облучатель-рециркулятор', 'Терминальная станция',
                                            'Технологическая станция', 'Изделие ЩИТ-ИВС', 'Маскиратор',
                                            'Обеспечивающее оборудование', 'Сервер', 'Переключатель'],
              'Электроника' : ['ИБП', 'Аккумуляторная батарея', 'Блок питания', 'Источник бесперебойного питания',
                               'Конвертер', 'Медиаконвертер', 'Фотоаппарат', 'Коммутатор'],
              'Прочее' : ['Швабра', 'Жалюзи', 'Перчатки медицинские латексные', 'Пленка для ламинатора', 'Доска', 
                          'Лампа', 'Ламповый блок', 'Набор для досок', 'Огнетушитель', 'Крепление потолочное',
                          'Разветвитель', 'Сетевой фильтр', 'Интерактивная система', 'Стационарная полка'],
                          
              }

    parser = DataParser(categories)

    # 02
    parser.parsing_xls("data/Двойченков счет 02 на 09.09.2024.xls", 42, 45, 'parsed_data/total_db.csv')
    # 02,6
    parser.parsing_xls('data/Двойченков счет 02,6 на 09.09.2024.xls', 42, 60, 'parsed_data/total_db.csv')
    parser.parsing_xls('data/Двойченков счет 02,6 на 09.09.2024.xls', 74, 74, 'parsed_data/total_db.csv')
    # 21
    parser.parsing_xls('data/Двойченков счет 21 на 09.09.2024.xls', 42, 75, 'parsed_data/total_db.csv')
    parser.parsing_xls('data/Двойченков счет 21 на 09.09.2024.xls', 89, 122, 'parsed_data/total_db.csv')
    parser.parsing_xls('data/Двойченков счет 21 на 09.09.2024.xls', 136, 153, 'parsed_data/total_db.csv')
    parser.parsing_xls('data/Двойченков счет 21 на 09.09.2024.xls', 167, 167, 'parsed_data/total_db.csv')
    #101
    parser.parsing_xls('data/Двойченков счет 101 на 09.09.2024.xls', 42, 63, 'parsed_data/total_db.csv')
    parser.parsing_xls('data/Двойченков счет 101 на 09.09.2024.xls', 77, 90, 'parsed_data/total_db.csv')
    parser.parsing_xls('data/Двойченков счет 101 на 09.09.2024.xls', 104, 116, 'parsed_data/total_db.csv')
    parser.parsing_xls('data/Двойченков счет 101 на 09.09.2024.xls', 130, 138, 'parsed_data/total_db.csv')
    parser.parsing_xls('data/Двойченков счет 101 на 09.09.2024.xls', 152, 154, 'parsed_data/total_db.csv')
    parser.parsing_xls('data/Двойченков счет 101 на 09.09.2024.xls', 168, 177, 'parsed_data/total_db.csv')
    parser.parsing_xls('data/Двойченков счет 101 на 09.09.2024.xls', 191, 214, 'parsed_data/total_db.csv')
    parser.parsing_xls('data/Двойченков счет 101 на 09.09.2024.xls', 228, 243, 'parsed_data/total_db.csv')
    parser.parsing_xls('data/Двойченков счет 101 на 09.09.2024.xls', 257, 273, 'parsed_data/total_db.csv')
    parser.parsing_xls('data/Двойченков счет 101 на 09.09.2024.xls', 287, 297, 'parsed_data/total_db.csv')
    parser.parsing_xls('data/Двойченков счет 101 на 09.09.2024.xls', 311, 323, 'parsed_data/total_db.csv')
    parser.parsing_xls('data/Двойченков счет 101 на 09.09.2024.xls', 337, 351, 'parsed_data/total_db.csv')
    parser.parsing_xls('data/Двойченков счет 101 на 09.09.2024.xls', 365, 371, 'parsed_data/total_db.csv')
    parser.parsing_xls('data/Двойченков счет 101 на 09.09.2024.xls', 385, 393, 'parsed_data/total_db.csv')
    parser.parsing_xls('data/Двойченков счет 101 на 09.09.2024.xls', 407, 422, 'parsed_data/total_db.csv')
    parser.parsing_xls('data/Двойченков счет 101 на 09.09.2024.xls', 436, 455, 'parsed_data/total_db.csv')
    parser.parsing_xls('data/Двойченков счет 101 на 09.09.2024.xls', 469, 484, 'parsed_data/total_db.csv')
    parser.parsing_xls('data/Двойченков счет 101 на 09.09.2024.xls', 498, 503, 'parsed_data/total_db.csv')
    parser.parsing_xls('data/Двойченков счет 101 на 09.09.2024.xls', 517, 521, 'parsed_data/total_db.csv')
    parser.parsing_xls('data/Двойченков счет 101 на 09.09.2024.xls', 535, 545, 'parsed_data/total_db.csv')
    parser.parsing_xls('data/Двойченков счет 101 на 09.09.2024.xls', 559, 579, 'parsed_data/total_db.csv')
    parser.parsing_xls('data/Двойченков счет 101 на 09.09.2024.xls', 593, 616, 'parsed_data/total_db.csv')
    parser.parsing_xls('data/Двойченков счет 101 на 09.09.2024.xls', 630, 651, 'parsed_data/total_db.csv')
    parser.parsing_xls('data/Двойченков счет 101 на 09.09.2024.xls', 665, 665, 'parsed_data/total_db.csv')
    #105
    parser.parsing_xls('data/Двойченков счет 105 на 09.09.2024.xls', 42, 69, 'parsed_data/total_db.csv')
    parser.parsing_xls('data/Двойченков счет 105 на 09.09.2024.xls', 83, 102, 'parsed_data/total_db.csv')
    parser.parsing_xls('data/Двойченков счет 105 на 09.09.2024.xls', 116, 138, 'parsed_data/total_db.csv')
    parser.parsing_xls('data/Двойченков счет 105 на 09.09.2024.xls', 152, 152, 'parsed_data/total_db.csv')
    #111
    parser.parsing_xls('data/Двойченков счет 111 на 09.09.2024.xls', 42, 42, 'parsed_data/total_db.csv')

    parser.clean_and_save_csv('parsed_data/total_db.csv')

    # df = pd.read_csv('parsed_data/total_db.csv', encoding='cp1251', sep='|')

        # Присвоение уникальных номеров
    parser.assign_unique_numbers('parsed_data/total_db.csv', 'result_data.csv')

    # Сохранение результата в базу данных
    df = pd.read_csv('result_data.csv', encoding='cp1251', sep='|')
    parser.save_to_database(df, 'equipments.db', 'table_name')

