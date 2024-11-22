import sqlite3
import os
import qrcode
import tkinter as tk
from tkinter import Toplevel, messagebox, StringVar, OptionMenu, PhotoImage
from tkinter import Toplevel, StringVar, OptionMenu, Button, Text
import sqlite3
import csv
from docx import Document
from datetime import datetime

class Equipment:
    def __init__(self):
        self.number_categories = {
            'Мебель': ['0001'], 
            'Накопитель': ['0002'],
            'АП': ['0007'],
            'Аудио и видео оборудование': ['0006'],
            'Бытовая техника': ['0013'],
            'Документы': ['0015'],
            'Проекционное оборудование': ['0012'],
            'Компьютерная техника': ['0008'],
            'Офисные принадлежности': ['0009'],
            'Офисное оборудование': ['0011'],
            'Рабочее место': ['0004'],
            'Расходные материалы': ['0010'],
            'Специальное оборудование': ['0014'],
            'Электроника': ['0005'],
            'Прочее': ['0003'],
        }

        self.employees = {
            '': [''],
            'SAN': ['SAN'],
            'KIA': ['KIA'], 
            'SAI': ['SAI'], 
            'ZDA': ['ZDA'], 
            'INA': ['INA'], 
            'BDI': ['BDI'], 
            'GVV': ['GVV'], 
            'KAV': ['KAV'],
            'KSV': ['KSV'],
            'VMV': ['VMV'], 
            'DDV': ['DDV'], 
            'IMA': ['IMA'], 
            'MVA': ['MVA'], 
            'KEV': ['KEV'], 
            'KDM': ['KDM'], 
            'MLS': ['MLS'], 
            'AAV': ['AAV']
        }

        self.room_number = {
            '': [''],
            'r406a': ['r406a'], 
            'r406b': ['r406b'], 
            'r408': ['r408'], 
            'r410': ['r410'], 
            'r412': ['r412'], 
            'r416': ['r416'], 
            'r418': ['r418'], 
            'r420': ['r420'], 
            'r420+1': ['r420+1'], 
            'r422': ['r422'], 
            'r424': ['r424'], 
            'r6UNC': ['r6UNC'], 
            'r11L': ['r11L'], 
            'r11P': ['r11P'], 
            'r15': ['r15'], 
            'rExtLab': ['rExtLab']
        }

        self.status = {
            '': ['']
        }
            
    def generate_unique_number(self, category):
        conn = sqlite3.connect('equipmts.db')
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT "Уникальный номер" FROM inventory WHERE "Категория" = ?
        ''', (category,))
        
        existing_numbers = cursor.fetchall()
        existing_numbers = {num[0] for num in existing_numbers}

        number = 1
        while True:
            unique_number = f"{category}-{number}"
            if unique_number not in existing_numbers:
                cursor.execute('''
                    SELECT COUNT(*) FROM inventory WHERE "Уникальный номер" = ?
                ''', (unique_number,))
                count = cursor.fetchone()[0]
                if count == 0:
                    break
            number += 1

        conn.close()
        return unique_number

    def add_item_to_inventory(self, item_data):
        conn = sqlite3.connect('equipmts.db')
        cursor = conn.cursor()

        qr_code_path = self.create_qr_code(item_data[0])

        cursor.execute('''
            INSERT INTO inventory (
                "Уникальный номер",
                "Наименование объекта нефинансового актива",
                "Номер (код) объекта учета (инвентарный или иной)",
                "Единица измерения",
                "цена (оценочная стоимость), руб",
                "количество",
                "сумма, руб.",
                "номер (код) счета",
                "Примечание",
                "Номер документа",
                "Дата",
                "Cтатус",
                "Закреплен",
                "Номер кабинета",
                "Категория",
                "qr_code"
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (*item_data, qr_code_path))

        conn.commit()
        conn.close()

    def create_qr_code(self, data):
        output_dir = 'qr_codes'
        os.makedirs(output_dir, exist_ok=True)
        qr_code_filename = os.path.join(output_dir, f'qr_code_{data}.png')
        
        if not os.path.exists(qr_code_filename):
            qr = qrcode.QRCode(version=1, box_size=10, border=5)
            qr.add_data(data)
            qr.make(fit=True)
            img = qr.make_image(fill='black', back_color='white')
            img.save(qr_code_filename)
        
        return qr_code_filename

    def get_row_by_qrcode(self, qrcode_number):
        """Retrieve a row from the database using the unique QR code number."""
        conn = sqlite3.connect('equipmts.db')
        cursor = conn.cursor()
        
        print(f"Ищем запись с QR-кодом: {qrcode_number}")
        cursor.execute("SELECT * FROM inventory WHERE [Уникальный номер] = ?", (qrcode_number,))
        row = cursor.fetchone()
        conn.close()
        
        return row 
    
    def get_item_by_accounting_code(self, accounting_code):
        """Retrieve an item from the database using the accounting code."""
        conn = sqlite3.connect('equipmts.db')
        cursor = conn.cursor()
        
        cursor.execute("SELECT * FROM inventory WHERE [Номер (код) объекта учета (инвентарный или иной)] = ?", (accounting_code,))
        row = cursor.fetchone()
        conn.close()
        
        return row
    
    def get_items_by_employee(self, employee_name):
        conn = sqlite3.connect('equipmts.db')
        cursor = conn.cursor()

        cursor.execute("SELECT * FROM inventory WHERE [Закреплен] = ?", (employee_name,))
        rows = cursor.fetchall()
        # Вывод результатов
        print("Уникальный номер | Наименование объекта нефинансового актива | Номер (код) объекта учета | Единица измерения | Цена (оценочная стоимость), руб | Количество | Сумма, руб. | Номер (код) счета | Примечание | Номер документа | Дата | Статус | Закреплен | Номер кабинета | Категория")
        print("-" * 150)
        if rows:
            print(f'\nПо запросу на сотрудника {employee_name} закреплена следующая техника:\n')
            for row in rows:
                print(" | ".join(map(str, row)))
        else:
            print(f"Нет техники, принадлежащей сотруднику {employee_name}.")

        # Закрытие соединения
        conn.close()

        return rows

    def get_items_by_room(self, room_number):
        conn = sqlite3.connect('equipmts.db')
        cursor = conn.cursor()

        cursor.execute("SELECT * FROM inventory WHERE [Номер кабинета] = ?", (room_number,))
        rows = cursor.fetchall()
        # Вывод результатов
        print("Уникальный номер | Наименование объекта нефинансового актива | Номер (код) объекта учета | Единица измерения | Цена (оценочная стоимость), руб | Количество | Сумма, руб. | Номер (код) счета | Примечание | Номер документа | Дата | Статус | Закреплен | Номер кабинета | Категория")
        print("-" * 150)
        if rows:
            print(f'\nПо запросу в кабинете {room_number} закреплена следующая техника:\n')
            for row in rows:
                print(" | ".join(map(str, row)))
        else:
            print(f"Нет техники, закрепленной за кабинетом {room_number}.")

        # Закрытие соединения
        conn.close()

        return rows

class EquipmentApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Сканер QRcode")
        self.equipment = Equipment()

        # Создаем фрейм для центрирования элементов
        self.main_frame = tk.Frame(root)
        self.main_frame.grid(row=0, column=0, sticky='nsew')  # Занимает всю ширину и высоту

        # Настройка растяжения строк и столбцов
        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.columnconfigure(1, weight=1)

        self.qr_input_label = tk.Label(self.main_frame, text="Введите уникальный номер QR-кода:")
        self.qr_input_label.grid(row=1, column=0, padx=5, pady=5, sticky='e')

        self.qr_input = tk.Entry(self.main_frame)
        self.qr_input.grid(row=1, column=1, padx=5, pady=5, sticky='w')
        
        self.qr_input.bind("<Return>", self.on_enter)

        self.result_frame = tk.Frame(self.main_frame)
        self.result_frame.grid(row=2, column=0, columnspan=2, pady=10)

        self.query_button = tk.Button(self.main_frame, text="Задать запрос по QR-коду", command=self.get_data_by_qr)
        self.query_button.grid(row=3, column=0, columnspan=2, pady=10)

        self.add_button = tk.Button(self.main_frame, text="Добавить новые данные", command=self.open_add_data_window)
        self.add_button.grid(row=4, column=0, columnspan=2, pady=10)

        self.search_employee_button = tk.Button(self.main_frame, text="Поиск по сотруднику", command=self.open_employee_search)
        self.search_employee_button.grid(row=5, column=0, columnspan=2, pady=10)

        self.search_room_button = tk.Button(self.main_frame, text="Поиск по номеру кабинета", command=self.open_room_search)
        self.search_room_button.grid(row=6, column=0, columnspan=2, pady=10)


    def open_employee_search(self):
        employee_keys = list(self.equipment.employees.keys())  # Convert to list
        self.open_search_window("Выберите сотрудника", employee_keys, self.search_by_employee)

    def open_room_search(self):
        room_keys = list(self.equipment.room_number.keys())  # Convert to list
        self.open_search_window("Выберите номер кабинета", room_keys, self.search_by_room)

    def open_search_window(self, title, options, search_function):
        search_window = Toplevel(self.root)
        search_window.title(title)

        selected_var = StringVar(search_window)
        selected_var.set(options[0])

        option_menu = OptionMenu(search_window, selected_var, *options)
        option_menu.pack(padx=10, pady=10)

        search_button = tk.Button(search_window, text="Поиск", command=lambda: search_function(selected_var.get(), search_window))
        search_button.pack(pady=10)

    def search_by_employee(self, employee, window):
        results = self.equipment.get_items_by_employee(employee)
        self.display_search_results_two(results)
        window.destroy()

    def search_by_room(self, room_number, window):
        results = self.equipment.get_items_by_room(room_number)
        self.display_search_results_two(results)
        window.destroy()


    def get_data_by_qr(self):
        """Retrieve data from the database using the QR code number."""
        qrcode_number = self.qr_input.get().strip()
        if not qrcode_number:
            messagebox.showwarning("Предупреждение", "Пожалуйста, введите QR-код.")
            return

        result = self.equipment.get_row_by_qrcode(qrcode_number)

        for widget in self.result_frame.winfo_children():
            widget.destroy()

        self.qr_input.delete(0, tk.END)

        if result:
            try:
                unique_number, name, accounting_code, unit, price, quantity, total_sum, account_code, note, document_number, date, status, fixed, cabinet_number, category, qr_code = result
            except ValueError as e:
                messagebox.showerror("Ошибка", f"Ошибка при извлечении данных: {e}")
                return

            data = [
                ("Уникальный номер", unique_number),
                ("Наименование", name),
                ("Код объекта", accounting_code),
                ("Единица измерения", unit),
                ("Цена", f"{price} руб."),
                ("Количество", quantity),
                ("Сумма", f"{total_sum} руб."),
                ("Код счета", account_code),
                ("Примечание", note),
                ("Номер документа", document_number),
                ("Дата", date),
                ("Статус", status),
                ("Закреплен", fixed),
                ("Номер кабинета", cabinet_number),
                ("Категория", category),
                ("QR код", qr_code)
            ]
            for i, (label_text, value_text) in enumerate(data):
                label = tk.Label(self.result_frame, text=label_text)
                label.grid(row=i, column=0, padx=5, pady=2, sticky='e')
                value_label = tk.Label(self.result_frame, text=value_text)
                value_label.grid(row=i, column=1, padx=5, pady=2, sticky='w')

            # Добавляем кнопку для открытия QR-кода
            qr_code_button = tk.Button(self.result_frame, text="Посмотреть QR-код", command=lambda: self.show_qr_code(qr_code))
            qr_code_button.grid(row=len(data), column=0, columnspan=2, pady=10)

            # Добавляем кнопку для очистки результатов
            clear_button = tk.Button(self.result_frame, text="Очистить", command=self.clear_results)
            clear_button.grid(row=len(data) + 1, column=0, columnspan=2, pady=10)

        else:
            no_record_label = tk.Label(self.result_frame, text="Запись не найдена.")
            no_record_label.grid(row=0, column=0, columnspan=2, pady=10)

    def clear_results(self):
        """Очистить результаты запроса."""
        for widget in self.result_frame.winfo_children():
            widget.destroy()
        self.qr_input.delete(0, tk.END)

    def display_search_results_two(self, results):
        result_window = Toplevel(self.root)
        result_window.title("Результаты поиска")

        text_area = Text(result_window, wrap='word')
        text_area.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)  # Убедитесь, что текстовое поле заполняет окно

        # Форматирование результатов
        if results:
            output = "Уникальный номер | Наименование объекта нефинансового актива | Номер (код) объекта учета | Единица измерения | Цена (оценочная стоимость), руб | Количество | Сумма, руб. | Номер (код) счета | Примечание | Номер документа | Дата | Статус | Закреплен | Номер кабинета | Категория\n"
            output += "-" * 150 + "\n"
            for row in results:
                output += " | ".join(map(str, row)) + "\n"
            text_area.insert(tk.END, output)
        else:
            text_area.insert(tk.END, "Нет техники, соответствующей запросу.")

        # Кнопки для экспорта
        csv_button = Button(result_window, text="Сохранить в CSV", command=lambda: self.save_to_csv(results))
        csv_button.pack(pady=5)

        doc_button = Button(result_window, text="Сохранить в DOC", command=lambda: self.save_to_doc(results))
        doc_button.pack(pady=5)

    def save_to_csv(self, results):
        current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        filename = f'CSV_{current_time}.csv'

        with open(filename, mode='w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerow(["Уникальный номер", "Наименование объекта нефинансового актива", "Номер (код) объекта учета", "Единица измерения", "Цена (оценочная стоимость), руб", "Количество", "Сумма, руб.", "Номер (код) счета", "Примечание", "Номер документа", "Дата", "Статус", "Закреплен", "Номер кабинета", "Категория"])
            writer.writerows(results)
            
        # Уведомление об успешной записи в приложении
        root = tk.Tk()
        root.withdraw()  # Скрываем главное окно
        messagebox.showinfo("Успех", f"Данные успешно сохранены в {filename}")
        root.destroy()  # Закрываем окно после показа сообщения

    def save_to_doc(self, results):
        doc = Document()
        table = doc.add_table(rows=1, cols=len(results[0]))

        # Заполняем заголовки
        hdr_cells = table.rows[0].cells
        headers = ["Уникальный номер", "Наименование объекта нефинансового актива", "Номер (код) объекта учета", 
                "Единица измерения", "Цена (оценочная стоимость), руб", "Количество", "Сумма, руб.", 
                "Номер (код) счета", "Примечание", "Номер документа", "Дата", "Статус", "Закреплен", 
                "Номер кабинета", "Категория"]
        
        for i, header in enumerate(headers):
            hdr_cells[i].text = header

        # Заполняем данные
        for row in results:
            row_cells = table.add_row().cells  # Добавляем новую строку
            for i, value in enumerate(row):
                if i < len(row_cells):  # Проверяем, что индекс не выходит за пределы
                    row_cells[i].text = str(value)

        current_datetime = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        filename = f"DOC_{current_datetime}.docx"
        doc.save(filename)  # Сохраняем документ

        # Уведомление об успешной записи в приложении
        root = tk.Tk()
        root.withdraw()  # Скрываем главное окно
        messagebox.showinfo("Успех", f"Документ успешно сохранен: {filename}")
        root.destroy()  # Закрываем окно после показа сообщения

    def clear_results(self):
        """Очистить результаты запроса."""
        for widget in self.result_frame.winfo_children():
            widget.destroy()
        self.qr_input.delete(0, tk.END)
    

    def show_qr_code(self, qr_code_path):
        """Open a new window to display the QR code image."""
        qr_window = Toplevel(self.root)
        qr_window.title("QR Код")

        # Загружаем изображение QR-кода
        qr_image = PhotoImage(file=qr_code_path)

        img_label = tk.Label(qr_window, image=qr_image)
        img_label.image = qr_image  # Сохраняем ссылку на изображение
        img_label.pack(padx=10, pady=10)

        close_button = tk.Button(qr_window, text="Закрыть", command=qr_window.destroy)
        close_button.pack(pady=10)

    def open_add_data_window(self):
        add_window = Toplevel(self.root)
        add_window.title("Добавить новые данные")

        tk.Label(add_window, text="Категория:").grid(row=0, column=0, padx=5, pady=5)
        
        category_var = StringVar(add_window)
        category_var.set(list(self.equipment.number_categories.keys())[0])  # Устанавливаем значение по умолчанию

        category_menu = OptionMenu(add_window, category_var, *self.equipment.number_categories.keys())
        category_menu.grid(row=0, column=1, padx=5, pady=5)

        tk.Label(add_window, text="*Наименование объекта:").grid(row=1, column=0, padx=5, pady=5)
        name_entry = tk.Entry(add_window)
        name_entry.grid(row=1, column=1, padx=5, pady=5)

        tk.Label(add_window, text="*Номер (код) объекта учета:").grid(row=2, column=0, padx=5, pady=5)
        accounting_code_entry = tk.Entry(add_window)
        accounting_code_entry.grid(row=2, column=1, padx=5, pady=5)

        tk.Label(add_window, text="Единица измерения:").grid(row=3, column=0, padx=5, pady=5)
        unit_entry = tk.Entry(add_window)
        unit_entry.grid(row=3, column=1, padx=5, pady=5)

        tk.Label(add_window, text="Цена (руб):").grid(row=4, column=0, padx=5, pady=5)
        price_entry = tk.Entry(add_window)
        price_entry.grid(row=4, column=1, padx=5, pady=5)

        tk.Label(add_window, text="Количество:").grid(row=5, column=0, padx=5, pady=5)
        quantity_entry = tk.Entry(add_window)
        quantity_entry.grid(row=5, column=1, padx=5, pady=5)

        tk.Label(add_window, text="Номер (код) счета:").grid(row=6, column=0, padx=5, pady=5)
        account_code_entry = tk.Entry(add_window)
        account_code_entry.grid(row=6, column=1, padx=5, pady=5)

        tk.Label(add_window, text="Примечание:").grid(row=7, column=0, padx=5, pady=5)
        note_entry = tk.Entry(add_window)
        note_entry.grid(row=7, column=1, padx=5, pady=5)

        tk.Label(add_window, text="Номер документа:").grid(row=8, column=0, padx=5, pady=5)
        document_number_entry = tk.Entry(add_window)
        document_number_entry.grid(row=8, column=1, padx=5, pady=5)

        tk.Label(add_window, text="Дата:").grid(row=9, column=0, padx=5, pady=5)
        date_entry = tk.Entry(add_window)
        date_entry.grid(row=9, column=1, padx=5, pady=5)

        tk.Label(add_window, text="Статус:").grid(row=10, column=0, padx=5, pady=5)

        status_var = StringVar(add_window)
        status_var.set(list(self.equipment.status.keys())[0])

        status_menu = OptionMenu(add_window, status_var, *self.equipment.status.keys())
        status_menu.grid(row=10, column=1, padx=5, pady=5)

        tk.Label(add_window, text="Закреплен (выберите сотрудника):").grid(row=11, column=0, padx=5, pady=5)
        
        employees_var = StringVar(add_window)
        employees_var.set(list(self.equipment.employees.keys())[0]) 

        employees_menu = OptionMenu(add_window, employees_var, *self.equipment.employees.keys())
        employees_menu.grid(row=11, column=1, padx=5, pady=5)

        tk.Label(add_window, text="Номер кабинета:").grid(row=12, column=0, padx=5, pady=5)

        room_var = StringVar(add_window)
        room_var.set(list(self.equipment.room_number.keys())[0]) 

        room_menu = OptionMenu(add_window, room_var, *self.equipment.room_number.keys())
        room_menu.grid(row=12, column=1, padx=5, pady=5)

        save_button = tk.Button(add_window, text="Сохранить", command=lambda: self.save_item(
            category_var.get(),
            name_entry.get(),
            accounting_code_entry.get(),
            unit_entry.get(),
            price_entry.get(), 
            quantity_entry.get(),
            account_code_entry.get(),
            note_entry.get(),
            document_number_entry.get(),
            date_entry.get(),
            status_var.get(),
            employees_var.get(),
            room_var.get(),
            add_window
        ))
        save_button.grid(row=13, column=0, columnspan=2, pady=10)

    def save_item(self, category_name, name, accounting_code, unit, price, quantity, account_code, note, document_number, date, status, employee, cabinet_number, add_window):

        # Проверка на наименование объекта поля
        if not all([name, accounting_code]):
            messagebox.showerror("Ошибка", "Пожалуйста, заполните обязательные поля (*)!")
            return

        # Генерация уникального номера
        category_number = self.equipment.number_categories[category_name][0]
        unique_number = self.equipment.generate_unique_number(category_number)

        # Обработка необязательных полей
        if price == '':
            price = 0.0  # Устанавливаем цену по умолчанию
        else:
            price = float(price)  # Преобразуем в число

        if quantity == '':
            quantity = 0  # Устанавливаем количество по умолчанию
        else:
            quantity = int(quantity)  # Преобразуем в целое число

        # Проверка на существование accounting_code
        existing_item = self.equipment.get_item_by_accounting_code(accounting_code)

            
        if existing_item:
            # Получаем уникальный номер существующего объекта
            existing_unique_number = existing_item[0]  # Предполагаем, что уникальный номер - это первый элемент кортежа
            messagebox.showwarning("Ошибка", f"Объект с номером {accounting_code} уже существует с уникальным номером {existing_unique_number}.")
            self.show_item_details(existing_item)
            return
        
        # Подготовка данных для добавления
        item_data = (
            unique_number,
            name,
            accounting_code,
            unit,
            price,
            quantity,
            float(price) * int(quantity),
            account_code,
            note,
            document_number,
            date,
            status,
            employee, 
            cabinet_number,
            category_name
        )

        # Добавление предмета в инвентарь
        self.equipment.add_item_to_inventory(item_data)

        # Закрытие окна после успешного добавления
        add_window.destroy()
        messagebox.showinfo("Успех", "Данные успешно добавлены!")

    def show_item_details(self, item):
        # Предполагаем, что item - это кортеж, содержащий данные о существующем объекте
        if item:
            details = (
                f"Детали объекта:\n"
                f"Имя: {item[1]}\n"  # Наименование объекта
                f"Код учета: {item[2]}\n"  # Номер (код) объекта учета
                f"Уникальный номер: {item[0]}\n"  # Уникальный номер
                f"Единица измерения: {item[3]}\n"  # Единица измерения
                f"Цена: {item[4]} руб.\n"  # Цена
                f"Количество: {item[5]}\n"  # Количество
                f"Сумма: {item[6]} руб.\n"  # Сумма
                f"Код счета: {item[7]}\n"  # Код счета
                f"Примечание: {item[8]}\n"  # Примечание
                f"Номер документа: {item[9]}\n"  # Номер документа
                f"Дата: {item[10]}\n"  # Дата
                f"Статус: {item[11]}\n"  # Статус
                f"Закреплен: {item[12]}\n"  # Закреплен
                f"Номер кабинета: {item[13]}\n"  # Номер кабинета
                f"Категория: {item[14]}\n"  # Категория
            )
            messagebox.showinfo("Детали объекта", details)
        else:
            messagebox.showwarning("Предупреждение", "Объект не найден.")

    def on_enter(self, event):
        """Обработчик события нажатия клавиши Enter."""
        self.clear_results()  # Очищаем предыдущие результаты
        self.get_data_by_qr()  # Выполняем запрос по QR-коду

    def clear_results(self):
        """Очистить результаты запроса."""
        for widget in self.result_frame.winfo_children():
            widget.destroy()
        self.result_frame.config(height=5, width=5) 

if __name__ == "__main__":
    root = tk.Tk()
    app = EquipmentApp(root)
    root.mainloop()
