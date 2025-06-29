import sys
import re
import os
import shutil
from copy import deepcopy
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QFileDialog, QVBoxLayout, QMessageBox

OUTPUT_DIR = "Разрезанные"


class CompetencyExtractor(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Разделение компетенций Word-файлов")
        self.setGeometry(300, 300, 400, 150)

        self.button = QPushButton("Выбрать файлы Word (.docx)", self)
        self.button.clicked.connect(self.process_files)

        layout = QVBoxLayout()
        layout.addWidget(self.button)
        self.setLayout(layout)

    def process_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Выберите Word-файлы", "", "Word Files (*.docx)")
        if not files:
            return

        if os.path.exists(OUTPUT_DIR):
            shutil.rmtree(OUTPUT_DIR)
        os.makedirs(OUTPUT_DIR)

        for file_path in files:
            try:
                self.process_file(file_path)
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Ошибка обработки файла {os.path.basename(file_path)}:\n{e}")

        QMessageBox.information(self, "Готово",
                                "Файлы успешно обработаны!\nСейчас выберите папку для сохранения результатов.")
        self.save_results()

    def save_results(self):
        dest_dir = QFileDialog.getExistingDirectory(self, "Выберите папку для сохранения результатов")
        if dest_dir:
            for filename in os.listdir(OUTPUT_DIR):
                shutil.copy(os.path.join(OUTPUT_DIR, filename), dest_dir)
            QMessageBox.information(self, "Сохранение завершено", f"Результаты сохранены в:\n{dest_dir}")
        else:
            QMessageBox.warning(self, "Отмена",
                                "Сохранение было отменено. Файлы остались в папке 'Разрезанные' рядом с программой.")

    def process_file(self, file_path):
        original_doc = Document(file_path)
        tables = original_doc.tables
        if len(tables) < 2:
            raise Exception("Недостаточно таблиц в документе")

        # Применяем границы ко всем таблицам в исходном документе
        for table in original_doc.tables:
            self.set_table_borders(table)

        first_table = tables[0]
        second_table = tables[1]

        # Читаем коды компетенций из первой таблицы, удаляем пробелы
        competencies = []
        for row in first_table.rows[1:]:
            code = row.cells[0].text.strip().replace(" ", "")
            competencies.append({'code': code, 'row': row})

        # Находим индекс начала раздела "Перечень заданий"
        per_list_start = None
        for i, para in enumerate(original_doc.paragraphs):
            if "Перечень заданий" in para.text:
                per_list_start = i
                break
        if per_list_start is None:
            raise Exception("Раздел 'Перечень заданий' не найден")

        # Собираем все элементы документа после заголовка "Перечень заданий"
        document_elements = []
        for element in original_doc.element.body[per_list_start + 1:]:
            document_elements.append(element)

        # Регулярка для поиска кода компетенции (пример: ОПК-1, ПК-2 и т.д.)
        competency_code_pattern = re.compile(r'^[A-ZА-Я]+-\d+', re.IGNORECASE)

        # Создаем словарь с номерами заданий и строками из второй таблицы
        task_rows = []
        for row in second_table.rows[1:]:
            task_num_text = row.cells[0].text.strip()
            try:
                task_num = int(task_num_text.replace(".", "").strip())
                task_rows.append((task_num, row))
            except ValueError:
                continue

        for comp in competencies:
            new_doc = Document()

            # 1. Добавляем первую таблицу с компетенцией
            new_table1 = new_doc.add_table(rows=1, cols=len(first_table.columns))
            self.copy_column_widths(first_table, new_table1)
            self.set_table_borders(new_table1)
            for i, cell in enumerate(first_table.rows[0].cells):
                new_table1.rows[0].cells[i].text = cell.text
            new_row = new_table1.add_row()
            for i, cell in enumerate(comp['row'].cells):
                new_row.cells[i].text = cell.text

            new_doc.add_paragraph("\n")

            # 2. Добавляем вторую таблицу с заданиями
            task_nums_for_comp = self.extract_task_nums_from_first_table_row(comp['row'])
            new_table2 = new_doc.add_table(rows=1, cols=len(second_table.columns))
            self.copy_column_widths(second_table, new_table2)
            self.set_table_borders(new_table2)
            for i, cell in enumerate(second_table.rows[0].cells):
                new_table2.rows[0].cells[i].text = cell.text
            for num, row in task_rows:
                if num in task_nums_for_comp:
                    new_row = new_table2.add_row()
                    for i, cell in enumerate(row.cells):
                        new_row.cells[i].text = cell.text

            new_doc.add_paragraph("\n")

            # 3. Добавляем заголовок "Перечень заданий" перед третьим блоком
            heading = new_doc.add_paragraph("Перечень заданий")
            heading.style = 'Heading 2'  # Можно использовать другой стиль
            new_doc.add_paragraph("\n")

            # 4. Копируем содержимое для текущей компетенции
            copying = False
            current_comp_elements = []

            for element in document_elements:
                if element.tag.endswith('p'):
                    para = self._element_to_paragraph(element)
                    para_text = para.text.strip().replace(" ", "")

                    if para_text == comp['code']:
                        copying = True
                        continue
                    elif competency_code_pattern.match(para_text) and copying:
                        break

                if copying:
                    current_comp_elements.append(element)

            # Добавляем все собранные элементы в новый документ
            for element in current_comp_elements:
                new_doc.element.body.append(deepcopy(element))

            # Применяем границы ко всем таблицам в новом документе
            for table in new_doc.tables:
                self.set_table_borders(table)

            filename = f"{comp['code']}_{os.path.basename(file_path)}"
            output_path = os.path.join(OUTPUT_DIR, filename)
            new_doc.save(output_path)

    def _element_to_paragraph(self, element):
        """Вспомогательная функция для преобразования элемента в параграф"""
        from docx.oxml import parse_xml
        from docx.text.paragraph import Paragraph
        return Paragraph(parse_xml(element.xml), None)

    def set_table_borders(self, table):
        """Устанавливает все границы таблицы (внешние и внутренние)"""
        tbl = table._tbl

        # Получаем свойства таблицы (tblPr)
        tblPr = tbl.tblPr
        if tblPr is None:
            tblPr = OxmlElement('w:tblPr')
            tbl.insert(0, tblPr)

        # Удаляем старые границы, если они есть
        existing_borders = tblPr.find(qn('w:tblBorders'))
        if existing_borders is not None:
            tblPr.remove(existing_borders)

        # Создаем новые границы
        tblBorders = OxmlElement('w:tblBorders')

        # Настраиваем все типы границ
        borders = [
            ('top', 'single', 4, '000000'),
            ('left', 'single', 4, '000000'),
            ('bottom', 'single', 4, '000000'),
            ('right', 'single', 4, '000000'),
            ('insideH', 'single', 4, '000000'),
            ('insideV', 'single', 4, '000000')
        ]

        # Добавляем все границы
        for border_type, border_style, border_size, border_color in borders:
            border = OxmlElement(f'w:{border_type}')
            border.set(qn('w:val'), border_style)
            border.set(qn('w:sz'), str(border_size))
            border.set(qn('w:space'), '0')
            border.set(qn('w:color'), border_color)
            tblBorders.append(border)

        # Добавляем границы к свойствам таблицы
        tblPr.append(tblBorders)

    def extract_task_nums_from_first_table_row(self, row):
        """Извлекает номера заданий из строки таблицы"""
        text = row.cells[-1].text.strip()
        match = re.match(r'(\d+)-(\d+)', text)
        if match:
            start, end = int(match.group(1)), int(match.group(2))
            return list(range(start, end + 1))
        else:
            nums = re.findall(r'\d+', text)
            return list(map(int, nums)) if nums else []

    def copy_column_widths(self, source_table, target_table):
        """Копирует ширину столбцов из исходной таблицы в целевую"""
        for i, source_col in enumerate(source_table.columns):
            if i < len(target_table.columns):
                target_col = target_table.columns[i]
                target_col.width = source_col.width

    def append_elements_with_images(self, source_elements, target_doc):
        """
        Копирует элементы (параграфы, таблицы и изображения) в целевой документ
        """
        for element in source_elements:
            target_doc.element.body.append(deepcopy(element))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = CompetencyExtractor()
    window.show()
    sys.exit(app.exec_())