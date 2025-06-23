import sys
import re
import os
import shutil
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

        QMessageBox.information(self, "Готово", "Файлы успешно обработаны!\nСейчас выберите папку для сохранения результатов.")
        self.save_results()

    def save_results(self):
        dest_dir = QFileDialog.getExistingDirectory(self, "Выберите папку для сохранения результатов")
        if dest_dir:
            for filename in os.listdir(OUTPUT_DIR):
                shutil.copy(os.path.join(OUTPUT_DIR, filename), dest_dir)
            QMessageBox.information(self, "Сохранение завершено", f"Результаты сохранены в:\n{dest_dir}")
        else:
            QMessageBox.warning(self, "Отмена", "Сохранение было отменено. Файлы остались в папке 'Разрезанные' рядом с программой.")

    def process_file(self, file_path):
        original_doc = Document(file_path)
        tables = original_doc.tables
        if len(tables) < 2:
            raise Exception("Недостаточно таблиц в документе")

        first_table = tables[0]
        second_table = tables[1]

        # Читаем коды компетенций из первой таблицы, удаляем пробелы
        competencies = []
        for row in first_table.rows[1:]:
            code = row.cells[0].text.strip().replace(" ", "")
            competencies.append({'code': code, 'row': row})

        paragraphs = original_doc.paragraphs

        # Находим индекс начала раздела "Перечень заданий"
        per_list_start = None
        for i, para in enumerate(paragraphs):
            if "Перечень заданий" in para.text:
                per_list_start = i
                break
        if per_list_start is None:
            raise Exception("Раздел 'Перечень заданий' не найден")

        per_list_paragraphs = paragraphs[per_list_start + 1:]

        # Регулярка для поиска кода компетенции (пример: ОПК-1, ПК-2 и т.д.)
        competency_code_pattern = re.compile(r'^[A-ZА-Я]+-\d+', re.IGNORECASE)

        # Создаем словарь с номерами заданий и строками из второй таблицы (для старого подхода)
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

            # Копируем первую таблицу (заголовок + строка с текущей компетенцией)
            new_table1 = new_doc.add_table(rows=1, cols=len(first_table.columns))
            self.copy_column_widths(first_table, new_table1)
            self.set_table_borders(new_table1)
            for i, cell in enumerate(first_table.rows[0].cells):
                new_table1.rows[0].cells[i].text = cell.text
            new_row = new_table1.add_row()
            for i, cell in enumerate(comp['row'].cells):
                new_row.cells[i].text = cell.text

            new_doc.add_paragraph("\n")

            # СТАРЫЙ ПОДХОД: Копируем вторую таблицу с заданиями по номерам, если они есть в первой таблице
            # Для этого извлечём номера заданий из первой таблицы, если есть
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

            # НОВЫЙ ПОДХОД: копируем блок текста заданий по названию кода компетенции
            copying = False
            for para in per_list_paragraphs:
                para_text = para.text.strip()
                clean_para_text = para_text.replace(" ", "")
                if clean_para_text == comp['code']:
                    copying = True
                    continue
                elif competency_code_pattern.match(clean_para_text) and copying:
                    break

                if copying:
                    new_doc.add_paragraph(para.text)

            filename = f"{comp['code']}_{os.path.basename(file_path)}"
            output_path = os.path.join(OUTPUT_DIR, filename)
            new_doc.save(output_path)

    def extract_task_nums_from_first_table_row(self, row):
        # Парсим диапазон заданий из последней ячейки, например "1-5"
        text = row.cells[-1].text.strip()
        match = re.match(r'(\d+)-(\d+)', text)
        if match:
            start, end = int(match.group(1)), int(match.group(2))
            return list(range(start, end + 1))
        else:
            # Если формат не подходит, можно попытаться вернуть отдельные номера, например, "1,3,5"
            nums = re.findall(r'\d+', text)
            return list(map(int, nums)) if nums else []

    def set_table_borders(self, table):
        tbl = table._tbl
        tblPr = tbl.tblPr
        borders = OxmlElement('w:tblBorders')

        for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
            border = OxmlElement(f'w:{border_name}')
            border.set(qn('w:val'), 'single')
            border.set(qn('w:sz'), '8')
            border.set(qn('w:space'), '0')
            border.set(qn('w:color'), '000000')
            borders.append(border)

        tblPr.append(borders)

    def copy_column_widths(self, source_table, target_table):
        for i, source_col in enumerate(source_table.columns):
            if i < len(target_table.columns):
                target_col = target_table.columns[i]
                target_col.width = source_col.width

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = CompetencyExtractor()
    window.show()
    sys.exit(app.exec_())
