import json
import csv


class parser:
    @staticmethod
    def parse(path):
        with open("settings.json", "r", encoding="utf-8") as settings_file:
            data1 = json.load(settings_file)

        result = {"Предметы": {}, "Преподаватели": {}}

        with open(path, "r", encoding="utf-8") as table:
            # Используем csv.reader для корректного парсинга
            reader = csv.reader(table)
            next(reader)  # Пропускаем заголовок

            for row in reader:
                answers = row[1:]  # Пропускаем первый столбец (дата)
                answer_index = 0

                for subject, teachers in data1["Subjects"].items():
                    if subject not in result["Предметы"]:
                        result["Предметы"][subject] = {}

                    for question in data1["Questions_for_subject"]:
                        if answer_index < len(answers):
                            # Убираем пробелы и кавычки
                            value = answers[answer_index].strip()
                            if value.isdigit():
                                value = int(value)

                            result["Предметы"][subject].setdefault(
                                question, []).append(value)
                            answer_index += 1
                    for teacher in teachers:
                        if answers[answer_index].strip() == 'да':
                            answer_index += 1

                            if teacher not in result["Преподаватели"]:
                                result["Преподаватели"][teacher] = {}
                            for question in data1["Questions_for_teachers"]:

                                if answer_index < len(answers):
                                    value = answers[answer_index].strip()

                                    if value.isdigit():
                                        value = int(value)

                                    result["Преподаватели"][teacher].setdefault(
                                        question, []).append(value)
                                    answer_index += 1
                        else:
                            answer_index += 1
                            for question in data1["Questions_for_teachers"]:
                                answer_index += 1

        return result
