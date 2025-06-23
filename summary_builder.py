import sys
import re
import os
from collections import defaultdict
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from PyQt5.QtWidgets import (QApplication, QWidget, QPushButton, QFileDialog,
                             QVBoxLayout, QMessageBox, QHBoxLayout)


class SummaryBuilder(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()
        self.comp_order = {'УК': 1, 'ОПК': 2, 'ПК': 3}
        self.summary_data = []
        self.all_tasks = []

    def init_ui(self):
        self.setWindowTitle("Построение сводного файла компетенций")
        self.setGeometry(300, 300, 500, 200)

        layout = QVBoxLayout()

        btn_layout = QHBoxLayout()

        self.select_btn = QPushButton("Выбрать папку с компетенциями", self)
        self.select_btn.clicked.connect(self.select_directory)

        self.build_btn = QPushButton("Построить сводный файл", self)
        self.build_btn.clicked.connect(self.build_summary)
        self.build_btn.setEnabled(False)

        btn_layout.addWidget(self.select_btn)
        btn_layout.addWidget(self.build_btn)

        layout.addLayout(btn_layout)
        self.setLayout(layout)

    def select_directory(self):
        dir_path = QFileDialog.getExistingDirectory(self, "Выберите папку с файлами компетенций")
        if dir_path:
            self.process_directory(dir_path)
            self.build_btn.setEnabled(True)
            QMessageBox.information(self, "Успех",
                                    f"Обработано {len(self.summary_data)} дисциплин!\n"
                                    f"Собрано данных о {len(self.all_tasks)} заданиях!\n"
                                    "Теперь можно построить сводный файл.")

    def process_directory(self, dir_path):
        """Обрабатывает все файлы DOCX в указанной директории"""
        self.summary_data = []
        self.all_tasks = []

        for filename in os.listdir(dir_path):
            if filename.endswith('.docx'):
                file_path = os.path.join(dir_path, filename)
                self.process_competency_file(file_path)

    def process_competency_file(self, file_path):
        """Извлекает данные о компетенции из файла"""
        doc = Document(file_path)
        if len(doc.tables) < 2:
            return

        # Извлекаем код компетенции из названия файла
        comp_code = os.path.basename(file_path).split('_')[0]

        # Первая таблица содержит основную информацию
        first_table = doc.tables[0]

        # Обрабатываем первую таблицу
        for row_idx, row in enumerate(first_table.rows):
            if row_idx == 0:  # Пропускаем заголовок
                continue

            cells = row.cells
            if len(cells) < 6:
                continue

            # Извлекаем данные из ячеек
            discipline = cells[3].text.strip()
            semester = cells[4].text.strip()
            tasks = cells[5].text.strip()

            # Сохраняем данные
            self.summary_data.append({
                'comp_code': comp_code,
                'discipline': discipline,
                'semester': semester,
                'tasks': tasks,
                'file_path': file_path
            })

        # Вторая таблица содержит детали заданий
        second_table = doc.tables[1]

        # Собираем данные из второй таблицы
        for row_idx, row in enumerate(second_table.rows):
            if row_idx == 0:  # Пропускаем заголовок
                continue

            cells = row.cells
            if len(cells) < 6:
                continue

            # Проверяем, содержит ли первая ячейка номер задания
            if re.match(r'^\d+\.', cells[0].text.strip()):
                # Сохраняем данные задания с привязкой к файлу
                task_data = {
                    'file_path': file_path,
                    'cells': [cell.text.strip() for cell in cells]
                }
                self.all_tasks.append(task_data)

    def build_summary(self):
        """Создает сводный файл на основе собранных данных"""
        if not self.summary_data or not self.all_tasks:
            QMessageBox.warning(self, "Ошибка", "Нет данных для построения сводного файла")
            return

        # Сортируем данные по типу компетенции, номеру и семестру
        sorted_data = sorted(
            self.summary_data,
            key=lambda x: (
                self.get_comp_order(x['comp_code']),
                self.parse_semester(x['semester']),
                x['discipline']
            )
        )

        # Создаем новый документ
        summary_doc = Document()

        # Добавляем шапку из шаблона
        self.add_template_header(summary_doc)

        # Добавляем название первой таблицы
        summary_doc.add_heading('Распределение тестовых заданий по компетенциям и дисциплинам', level=2)

        # Создаем таблицу
        table = summary_doc.add_table(rows=1, cols=6)
        table.style = 'Table Grid'
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        # Заголовки таблицы
        headers = table.rows[0].cells
        headers[0].text = "Код компетенции"
        headers[1].text = "Наименование компетенции"
        headers[2].text = "Наименование индикаторов"
        headers[3].text = "Наименование дисциплины/модуля/практики"
        headers[4].text = "Семестр"
        headers[5].text = "Номер задания"

        # Форматирование заголовков
        for cell in headers:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.bold = True
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Заполняем таблицу данными и создаем список дисциплин в порядке следования
        row_idx = 1
        comp_row_map = defaultdict(list)
        disciplines_order = []
        current_task_num = 1
        task_mapping = {}

        for comp_code in sorted(set(item['comp_code'] for item in sorted_data), key=self.get_comp_order):
            # Фильтруем дисциплины для текущей компетенции
            comp_disciplines = [item for item in sorted_data if item['comp_code'] == comp_code]

            # Сортируем дисциплины по семестру
            comp_disciplines_sorted = sorted(
                comp_disciplines,
                key=lambda x: self.parse_semester(x['semester'])
            )

            # Добавляем дисциплины в общий порядок
            for disc in comp_disciplines_sorted:
                disciplines_order.append(disc)

                # Рассчитываем количество заданий для дисциплины
                task_count = self.calculate_task_count(disc['tasks'])
                new_start = current_task_num
                new_end = current_task_num + task_count - 1
                current_task_num = new_end + 1

                # Сохраняем маппинг заданий
                task_mapping[disc['file_path']] = {
                    'discipline': disc['discipline'],
                    'start': new_start,
                    'end': new_end
                }

                # Добавляем строку в таблицу
                row_cells = table.add_row().cells
                row_cells[0].text = comp_code
                row_cells[1].text = ""
                row_cells[2].text = ""
                row_cells[3].text = disc['discipline']
                row_cells[4].text = disc['semester']
                row_cells[5].text = f"{new_start}-{new_end}"

                # Сохраняем индекс строки для объединения
                comp_row_map[comp_code].append(row_idx)
                row_idx += 1

        # Объединяем ячейки для компетенций
        for comp_code, row_indices in comp_row_map.items():
            if len(row_indices) > 1:
                start_row = min(row_indices)
                end_row = max(row_indices)

                # Объединяем ячейки в первых трех столбцах
                for col_idx in range(3):
                    self.merge_cells(table, start_row, end_row, col_idx)

                    # Центрируем текст в объединенных ячейках
                    cell = table.cell(start_row, col_idx)
                    for paragraph in cell.paragraphs:
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Добавляем вторую таблицу с распределением заданий по типам
        self.add_second_table(summary_doc, disciplines_order, task_mapping)

        # Сохраняем документ
        save_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить сводный файл", "", "Word Files (*.docx)")

        if save_path:
            summary_doc.save(save_path)
            QMessageBox.information(self, "Готово",
                                    f"Сводный файл успешно создан:\n{save_path}")

    def add_template_header(self, doc):
        """Добавляет шапку из шаблона"""
        doc.add_heading('Фонд оценочных средств', level=1)

        p = doc.add_paragraph()
        p.add_run('для оценки остаточных знаний обучающихся по направлению подготовки').bold = True

        doc.add_paragraph()
        p = doc.add_paragraph()
        p.add_run('Направление: ').bold = True
        p.add_run('')

        p = doc.add_paragraph()
        p.add_run('Профиль: ').bold = True
        p.add_run('')

        p = doc.add_paragraph()
        p.add_run('Год начала подготовки -- ').bold = True
        p.add_run('20__')

        doc.add_paragraph()

    def add_second_table(self, doc, disciplines_order, task_mapping):
        """Добавляет вторую таблицу с распределением заданий по типам"""
        doc.add_heading('Распределение заданий по типам и уровням сложности', level=2)
        doc.add_heading('Ключи к оцениванию', level=3)

        # Создаем таблицу
        table = doc.add_table(rows=1, cols=6)
        table.style = 'Table Grid'
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        # Заголовки таблицы
        headers = table.rows[0].cells
        headers[0].text = "№ задания"
        headers[1].text = "Верный ответ"
        headers[2].text = "Критерии"
        headers[3].text = "Тип задания"
        headers[4].text = "Уровень сложности"
        headers[5].text = "Время выполнения (мин.)"

        # Форматирование заголовков
        for cell in headers:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.bold = True
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Группируем задания по файлам
        tasks_by_file = defaultdict(list)
        for task in self.all_tasks:
            tasks_by_file[task['file_path']].append(task)

        # Добавляем задания в порядке дисциплин из первой таблицы
        for disc in disciplines_order:
            file_tasks = tasks_by_file.get(disc['file_path'], [])
            if not file_tasks:
                continue

            # Добавляем заголовок дисциплины
            row = table.add_row().cells
            row[0].merge(row[5])
            row[0].text = disc['discipline']
            for paragraph in row[0].paragraphs:
                for run in paragraph.runs:
                    run.bold = True
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

            # Определяем диапазон заданий для этой дисциплины
            mapping = task_mapping[disc['file_path']]
            start_num = mapping['start']

            # Добавляем задания с новыми номерами
            for idx, task in enumerate(file_tasks):
                row_cells = table.add_row().cells
                new_num = start_num + idx

                # Заменяем номер задания
                task_cells = task['cells'].copy()
                task_cells[0] = f"{new_num}."

                # Заполняем ячейки
                for i in range(min(6, len(task_cells))):
                    row_cells[i].text = task_cells[i]

    def merge_cells(self, table, start_row, end_row, col_idx):
        """Объединяет ячейки в таблице по вертикали"""
        cell_start = table.cell(start_row, col_idx)
        for row in range(start_row + 1, end_row + 1):
            cell_next = table.cell(row, col_idx)
            cell_start.merge(cell_next)

    def get_comp_order(self, code):
        """Определяет порядок сортировки для кода компетенции"""
        match = re.match(r'([А-Я]+)-(\d+)', code)
        if match:
            comp_type = match.group(1)
            comp_num = int(match.group(2))
            return (self.comp_order.get(comp_type, 99), comp_num)
        return (99, 99)  # Для некорректных кодов

    def parse_semester(self, semester_str):
        """Парсит семестр для корректной сортировки"""
        try:
            # Удаляем все нецифровые символы, кроме минуса
            clean_str = re.sub(r'[^\d-]', '', semester_str)

            # Обрабатываем диапазон семестров (например, "1-2")
            if '-' in clean_str:
                parts = clean_str.split('-')
                return int(parts[0])

            return int(clean_str) if clean_str else 0
        except ValueError:
            return 0

    def calculate_task_count(self, tasks_str):
        """Вычисляет количество заданий по строке"""
        # Удаляем все символы, кроме цифр, запятых и минусов
        clean_str = re.sub(r'[^\d,-]', '', tasks_str)

        # Обрабатываем диапазон (например, "1-16")
        if '-' in clean_str:
            parts = clean_str.split('-')
            if len(parts) == 2 and parts[0].isdigit() and parts[1].isdigit():
                return int(parts[1]) - int(parts[0]) + 1

        # Обрабатываем перечисление (например, "1,2,3")
        elif ',' in clean_str:
            return len(clean_str.split(','))

        # Одиночное число
        try:
            return int(clean_str)
        except ValueError:
            return 0  # Некорректный формат


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = SummaryBuilder()
    window.show()
    sys.exit(app.exec_())