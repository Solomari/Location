import os

import folium
from PyQt5.QtCore import QUrl
from PyQt5.QtWidgets import QMainWindow, QVBoxLayout, QWidget, QLineEdit, QPushButton, QLabel, QTableWidget, \
    QApplication, QTableWidgetItem, QPlainTextEdit, QHeaderView
from PyQt5.QtWebEngineWidgets import QWebEngineView
import pandas as pd


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Связка дислокации с заказом")
        self.setGeometry(20, 20, 1000, 1000)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        self.layout = QVBoxLayout()

        # Поле для ввода номера вагона
        self.input_field_wagon = QLineEdit()
        self.layout.addWidget(QLabel("Введите номер вагона:"))
        self.layout.addWidget(self.input_field_wagon)

        # Кнопка для выполнения поиска по номеру вагона
        self.search_button_wagon = QPushButton("Поиск по номеру вагона")
        self.search_button_wagon.clicked.connect(self.search_data_wagon)
        self.layout.addWidget(self.search_button_wagon)

        # Поле для ввода номера заказа
        self.input_field_order = QLineEdit()
        self.layout.addWidget(QLabel("Введите номер заказа:"))
        self.layout.addWidget(self.input_field_order)

        # Кнопка для выполнения поиска по номеру заказа
        self.search_button_order = QPushButton("Поиск по номеру заказа")
        self.search_button_order.clicked.connect(self.search_data_order)
        self.layout.addWidget(self.search_button_order)

        # текстовое поле для отображения номеров заказов
        self.order_numbers_text = QPlainTextEdit()
        self.layout.addWidget(QLabel("Информация о заказ(ах):"))
        self.layout.addWidget(self.order_numbers_text)

        # Виджет для отображения карты
        self.map_view = QWebEngineView()
        self.layout.addWidget(self.map_view)

        self.central_widget.setLayout(self.layout)

        # Считываем и объединяем данные из файлов Excel

        self.sap_df = pd.read_excel("file_input/SAP.xlsx")
        self.disl_df = pd.read_excel("file_input/disl.xlsx")

        self.sap_df['№ ТС'] = self.sap_df['№ ТС'].astype(str)
        self.disl_df['N вагона'] = self.disl_df['N вагона'].astype(str)

        # Объединение таблиц по общему столбцу '№ ТС' и 'N вагона'
        self.merged_df = pd.merge(self.sap_df, self.disl_df, left_on='№ ТС', right_on='N вагона')

        selected_columns = [' ЗК №','Сокр. наим. ОКРО грузоотпр.', 'Дата погр.', ' Заказчик', ' Грузополучатель',
                            'Наим. ст. назн.', 'Дата дисл.', 'Время дисл.', 'Наим. ст. дисл.', 'Посл. опер.',
                            'Дата дост.', 'N вагона']

        self.merged_df = self.merged_df[selected_columns]

        # Создаем экземпляр карты и отображаем
        self.create_and_show_map()


        # Таблица для вывода результатов
        self.table = QTableWidget()
        self.table.setColumnCount(9)
        self.table.setHorizontalHeaderLabels(
            ["Сокр. наим. ОКРО грузоотпр.", "Дата погр.", "N вагона", "Наим. ст. назн.", "Дата дисл.", "Время дисл.",
             "Наим. ст. дисл.", "Посл. опер.", "Дата дост."])

        # Устанавливаем флаг ResizeToContents для горизонтального заголовка
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.layout.addWidget(self.table)

    def create_and_show_map(self):
        # Создаем карту с начальными координатами
        m = folium.Map(location=[58.49617513783592, 49.75513364768645], zoom_start=10)

        # Добавляем маркер
        folium.Marker(location=[58.49617513783592, 49.75513364768645], popup="Segezha Group").add_to(m)

        # Получаем текущую директорию
        current_dir = os.path.dirname(__file__)

        # Создаем абсолютный путь к файлу map.html
        html_file_path = os.path.join(current_dir, "map.html")

        # Сохраняем карту в HTML-файл и отображаем его в виджете QWebEngineView
        m.save(html_file_path)
        self.map_view.setUrl(QUrl.fromLocalFile(html_file_path))

    def search_data_wagon(self):
        # Получаем введенное пользователем значение
        search_value = str(self.input_field_wagon.text())

        # Выполняем поиск
        result_df = self.merged_df[self.merged_df["N вагона"].astype(str) == search_value]

        # Очищаем текстовое поле с номерами заказов
        self.order_numbers_text.clear()

        # Получаем список уникальных номеров заказов для найденных вагонов
        order_numbers = result_df[" ЗК №"].unique()

        # Собираем информацию о заказах, заказчиках и грузополучателях
        order_info = ""
        for order_number in order_numbers:
            if pd.notna(order_number):  # Проверяем, что значение не является NaN
                try:
                    order_number_int = int(order_number)  # Пробуем преобразовать строку в целое число
                    order_data = result_df[
                        result_df[" ЗК №"] == order_number_int]  # Преобразовываем номер заказа в целое число
                    if not order_data.empty:  # Проверяем, есть ли данные по данному заказу
                        customer = order_data[" Заказчик"].iloc[0]  # Извлекаем заказчика
                        consignee = order_data[" Грузополучатель"].iloc[0]  # Извлекаем грузополучателя
                        last_disl_date = order_data["Дата дисл."].iloc[-1]  # Получаем последнюю дату дислокации
                        last_disl_station = order_data["Наим. ст. дисл."].iloc[
                            -1]  # Получаем последнюю станцию дислокации
                        last_operation = order_data["Посл. опер."].iloc[-1]  # Получаем последнюю операцию
                        delivery_date = order_data["Дата дост."].iloc[-1]  # Получаем дату доставки

                        # Добавляем информацию о последних данных в текстовое поле
                        order_info += f"Номер заказа: {order_number_int}, Заказчик: {customer}, " \
                                      f"Грузополучатель: {consignee}, \n" \
                                      f"Последняя дата дислокации: {last_disl_date}, \n" \
                                      f"Последняя станция дислокации: {last_disl_station}, \n" \
                                      f"Последняя операция: {last_operation}, \n" \
                                      f"Дата доставки: {delivery_date}\n\n"
                except ValueError:
                    print(f"Некорректное значение номера заказа: {order_number}")

        # Выводим информацию о заказах, заказчиках, грузополучателях и последних данных в текстовое поле
        self.order_numbers_text.setPlainText(order_info)

        # Удаляем строки, в которых отсутствуют значения в указанных столбцах
        result_df = result_df.dropna(
            subset=["Сокр. наим. ОКРО грузоотпр.", "Дата погр.", "N вагона", "Наим. ст. назн.", "Дата дисл.",
                    "Время дисл.", "Наим. ст. дисл.", "Посл. опер.", "Дата дост."])

        # Очищаем таблицу перед выводом новых данных
        self.table.setRowCount(0)

        print("Result DataFrame:")
        print(result_df)  # Выводим данные для отладки

        # Выбираем только определенные столбцы для отображения в таблице
        selected_columns = ["Сокр. наим. ОКРО грузоотпр.", "Дата погр.", "N вагона", "Наим. ст. назн.", "Дата дисл.",
                            "Время дисл.", "Наим. ст. дисл.", "Посл. опер.", "Дата дост."]
        result_df = result_df[selected_columns]

        # Преобразуем DataFrame в список списков
        data = result_df.values.tolist()

        # Устанавливаем количество строк в таблице равным количеству строк в данных
        self.table.setRowCount(len(data))

        # Заполняем таблицу данными из списка списков
        for row_index, row_data in enumerate(data):
            for col_index, value in enumerate(row_data):
                self.table.setItem(row_index, col_index, QTableWidgetItem(str(value)))

        # Автоматически изменяем размеры строк и столбцов
        self.table.resizeRowsToContents()
        self.table.resizeColumnsToContents()

        # Ваш код для отображения карты по номеру вагона здесь
        if not result_df.empty:
            last_station = result_df["Наим. ст. дисл."].iloc[-1]
            self.show_map_for_station(last_station)

    def show_map_for_station(self, station_name):
        # Считываем данные о координатах из файла
        locations_df = pd.read_excel("file_input/location.xlsx")
        # Фильтруем данные для получения координат выбранной станции
        station_data = locations_df[locations_df["Наим. ст. дисл."] == station_name]
        if not station_data.empty:
            # Получаем координаты выбранной станции
            latitude = station_data["Широта"].iloc[0]
            longitude = station_data["Долгота"].iloc[0]

            # Создаем карту с начальными координатами
            m = folium.Map(location=[latitude, longitude], zoom_start=10)

            # Добавляем маркер
            folium.Marker(location=[latitude, longitude], popup=station_name).add_to(m)

            # Получаем текущую директорию
            current_dir = os.path.dirname(__file__)

            # Создаем абсолютный путь к файлу map.html
            html_file_path = os.path.join(current_dir, "map.html")

            # Сохраняем карту в HTML-файл и отображаем его в виджете QWebEngineView
            m.save(html_file_path)
            self.map_view.setUrl(QUrl.fromLocalFile(html_file_path))
            print("Путь к файлу map.html:", html_file_path)
        else:
            print("Данные о координатах для выбранной станции не найдены.")

    def search_data_order(self):
        # Получаем введенное пользователем значение номера заказа
        search_value = self.input_field_order.text()

        try:
            # Преобразуем введенное значение в целое число
            search_value = int(search_value)

            # Находим все вагоны, соответствующие введенному номеру заказа
            wagons_for_order = self.merged_df[self.merged_df[' ЗК №'] == search_value]['N вагона'].tolist()

            # Создаем пустой список для хранения уникальных данных о заказах и вагонах
            unique_order_and_wagon_info = set()

            # Добавляем информацию о заказах и вагонах
            for wagon_number in wagons_for_order:
                unique_order_and_wagon_info.add((search_value, wagon_number))

            # Собираем информацию о заказах, заказчиках и грузополучателях для текстового поля
            order_info = ""
            for order_number, wagon_number in unique_order_and_wagon_info:
                if pd.notna(order_number):  # Проверяем, что значение не является NaN
                    try:
                        order_number_int = int(order_number)  # Пробуем преобразовать строку в целое число
                        order_data = self.merged_df[
                            self.merged_df[" ЗК №"] == order_number_int]  # Получаем данные по заказу
                        if not order_data.empty:  # Проверяем, есть ли данные по данному заказу
                            customer = order_data[" Заказчик"].iloc[0]  # Извлекаем заказчика
                            consignee = order_data[" Грузополучатель"].iloc[0]  # Извлекаем грузополучателя
                            last_disl_date = order_data["Дата дисл."].iloc[-1]  # Получаем последнюю дату дислокации
                            last_disl_station = order_data["Наим. ст. дисл."].iloc[
                                -1]  # Получаем последнюю станцию дислокации
                            last_operation = order_data["Посл. опер."].iloc[-1]  # Получаем последнюю операцию
                            delivery_date = order_data["Дата дост."].iloc[-1]  # Получаем дату доставки

                            # Добавляем информацию о последних данных в текстовое поле
                            order_info += f"Номер заказа: {order_number_int}, Вагон: {wagon_number}, Заказчик: {customer}, " \
                                          f"Грузополучатель: {consignee}, \n" \
                                          f"Последняя дата дислокации: {last_disl_date}, \n" \
                                          f"Последняя станция дислокации: {last_disl_station}, \n" \
                                          f"Последняя операция: {last_operation}, \n" \
                                          f"Дата доставки: {delivery_date}\n\n"

                    except ValueError:
                        print(f"Некорректное значение номера заказа: {order_number}")

            # Выводим информацию о заказах, заказчиках, грузополучателях и последних данных в текстовое поле
            self.order_numbers_text.setPlainText(order_info)

            # Теперь мы можем добавить логику для вывода в таблицу
            # Создаем DataFrame из данных по заказу
            result_df = pd.DataFrame(order_data, columns=self.merged_df.columns)

            # Удаляем дубликаты, если они есть
            result_df.drop_duplicates(inplace=True)

            # Очищаем таблицу перед выводом новых данных
            self.table.setRowCount(0)

            # Выводим данные найденных записей в таблицу
            selected_columns = ["Сокр. наим. ОКРО грузоотпр.", "Дата погр.", "N вагона",
                                "Наим. ст. назн.", "Дата дисл.", "Время дисл.", "Наим. ст. дисл.",
                                "Посл. опер.", "Дата дост."]
            result_df = result_df[selected_columns]

            # Преобразуем DataFrame в список списков
            data = result_df.values.tolist()

            # Устанавливаем количество строк в таблице равным количеству строк в данных
            self.table.setRowCount(len(data))

            # Заполняем таблицу данными из списка списков
            for row_index, row_data in enumerate(data):
                for col_index, value in enumerate(row_data):
                    self.table.setItem(row_index, col_index, QTableWidgetItem(str(value)))

            # Автоматически изменяем размеры строк и столбцов
            self.table.resizeRowsToContents()
            self.table.resizeColumnsToContents()

            # Ваш код для отображения карты по номеру вагона здесь
            if not result_df.empty:
                last_station = result_df["Наим. ст. дисл."].iloc[-1]
                self.show_map_for_station(last_station)

        except ValueError:
            # Если введенное значение не является числом, выводим сообщение об ошибке
            print("Номер заказа должен быть числом")


app = QApplication([])
main_window = MainWindow()
main_window.show()
app.exec_()
