from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment
from datetime import datetime
import locale
from parser_csv import parser  # Ваш парсер CSV

class Excel():
    def __init__(self, path, path_to_xlsx):
        self.csv = parser.parse(path)
        try:
            self.path_to_xlsx = path_to_xlsx
            self.wb = Workbook()
            self.wb = load_workbook(self.path_to_xlsx)
        except:
            self.wb = Workbook()
            self.Sheet1 = self.wb['Sheet']
            self.wb.remove(self.Sheet1)

    def create_list(self):
        locale.setlocale(category=locale.LC_ALL, locale="Russian")
        self.a = datetime.today()
        sheet = self.wb.create_sheet(title=self.a.strftime('%B')+'_предметы')

        # Устанавливаем ширину первой колонки и заголовок
        sheet.column_dimensions['A'].width = len(max(self.csv["Предметы"].keys()))
        sheet['A1'] = "Дисциплина"
        sheet['A1'].alignment = Alignment(horizontal='center')

        j = 2
        for i in list(self.csv["Предметы"].keys()):
            sheet[f'A{j}'] = i
            sheet[f'A{j}'].alignment = Alignment(horizontal='center', vertical='center')
            j += 1

        list_subj = list(self.csv["Предметы"].keys())
        list_quest = list(self.csv["Предметы"][list_subj[0]].keys())

        # Оставляем только те вопросы, где ВСЕ ответы — числа
        valid_questions = [
            q for q in list_quest if all(isinstance(ans, (int, float)) for subj in list_subj for ans in self.csv["Предметы"][subj][q])
        ]

        # Определяем последнюю колонку для записи среднего
        last_col = len(valid_questions) + 2  # +2, т.к. 1-я колонка уже занята дисциплинами

        # Заполняем заголовки вопросов и данных
        for col_idx, question in enumerate(valid_questions, start=2):  # Начинаем со 2-й колонки
            sheet.cell(row=1, column=col_idx, value=question).alignment = Alignment(horizontal='center')

        # Заполняем данные и считаем среднее
        for row_idx, subject in enumerate(list_subj, start=2):  # Начинаем со 2-й строки
            values = []
            for col_idx, question in enumerate(valid_questions, start=2):
                scores = self.csv["Предметы"][subject][question]

                avg_score = sum(scores) / len(scores) if scores else 0  # Среднее значение
                sheet.cell(row=row_idx, column=col_idx, value=avg_score).alignment = Alignment(horizontal='center')
                values.append(avg_score)

            # Записываем среднее значение по строке в последнюю колонку
            if values:
                sheet.cell(row=row_idx, column=last_col, value=sum(values) / len(values)).alignment = Alignment(horizontal='center')

        # Добавляем заголовок для колонки со средним
        sheet.cell(row=1, column=last_col, value="Итог").alignment = Alignment(horizontal='center')

        avg_row = len(list_subj) + 2


        # Среднее значение для столбца "Итог"
        total_values = [sheet.cell(row=row_idx, column=last_col).value for row_idx in range(2, len(list_subj) + 2)]
        total_avg = sum(total_values) / len(total_values) if total_values else 0
        sheet.cell(row=avg_row, column=last_col, value=total_avg).alignment = Alignment(horizontal='center')

        self.wb.save(self.path_to_xlsx)
    
    def create_teachers_list(self):
        locale.setlocale(category=locale.LC_ALL, locale="Russian")
        self.a = datetime.today()
        sheet = self.wb.create_sheet(title=self.a.strftime('%B')+"_преподаватели")

        # Устанавливаем ширину первой колонки и заголовок
        sheet.column_dimensions['A'].width = len(max(self.csv["Преподаватели"].keys()))
        sheet['A1'] = "Преподаватель"
        sheet['A1'].alignment = Alignment(horizontal='center')

        j = 2
        for i in list(self.csv["Преподаватели"].keys()):
            sheet[f'A{j}'] = i
            sheet[f'A{j}'].alignment = Alignment(horizontal='center', vertical='center')
            j += 1

        list_teachers = list(self.csv["Преподаватели"].keys())
        list_quest = list(self.csv["Преподаватели"][list_teachers[0]].keys())

        # Оставляем только те вопросы, где ВСЕ ответы — числа
        valid_questions = [
            q for q in list_quest if all(isinstance(ans, (int, float)) for teacher in list_teachers for ans in self.csv["Преподаватели"][teacher][q])
        ]

        # Определяем последнюю колонку для записи среднего
        last_col = len(valid_questions) + 2  # +2, т.к. 1-я колонка уже занята преподавателями

        # Заполняем заголовки вопросов и данных
        for col_idx, question in enumerate(valid_questions, start=2):  # Начинаем со 2-й колонки
            sheet.cell(row=1, column=col_idx, value=question).alignment = Alignment(horizontal='center')

        # Заполняем данные и считаем среднее
        for row_idx, teacher in enumerate(list_teachers, start=2):  # Начинаем со 2-й строки
            values = []
            for col_idx, question in enumerate(valid_questions, start=2):
                scores = self.csv["Преподаватели"][teacher][question]

                avg_score = sum(scores) / len(scores) if scores else 0  # Среднее значение
                sheet.cell(row=row_idx, column=col_idx, value=avg_score).alignment = Alignment(horizontal='center')
                values.append(avg_score)

            # Записываем среднее значение по строке в последнюю колонку
            if values:
                sheet.cell(row=row_idx, column=last_col, value=sum(values) / len(values)).alignment = Alignment(horizontal='center')

        # Добавляем заголовок для колонки со средним
        avg_row = len(list_teachers) + 2
        sheet.cell(row=1, column=last_col, value="Итог").alignment = Alignment(horizontal='center')
        total_values = [sheet.cell(row=row_idx, column=last_col).value for row_idx in range(2, len(list_teachers) + 2)]
        total_avg = sum(total_values) / len(total_values) if total_values else 0
        sheet.cell(row=avg_row, column=last_col, value=total_avg).alignment = Alignment(horizontal='center')
        self.wb.save(self.path_to_xlsx)



# nw = Excel("1.csv", "1.xlsx")
# nw.create_list()
# nw.create_teachers_list()

