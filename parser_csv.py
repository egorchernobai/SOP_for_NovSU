import json


class parser:
    def parse(path):
        with open("settings.json", "r", encoding="utf-8") as settings_file:
            data1 = json.load(settings_file)

        with open(path, "r", encoding="utf-8") as table:
            i = 0
            for line in table:
                if i == 0:
                    i+=1
                    continue

                answers = line[:-1].split("\",\"")[1:]
                result = {"Преподаватели":{}}
                answer_index = 0

                for subject, teachers in data1["Subjects"].items():
                    result[subject] = {"questions": {}}
                    
                    # Заполняем вопросы по предмету
                    for question in data1["Questions_for_subject"]:
                        if answer_index < len(answers):
                            try:
                                result[subject]["questions"][question].append(answers[answer_index])
                            except:
                                result[subject]["questions"][question] = []
                                result[subject]["questions"][question].append(answers[answer_index])
                            answer_index += 1
                    
                    # Заполняем вопросы по преподавателям
                    for teacher in teachers:
                        result["Преподаватели"][teacher] = {}
                        for question in data1["Questions_for_teachers"]:
                            if answer_index < len(answers):
                                try:
                                    result["Преподаватели"][teacher][question].append(answers[answer_index])
                                except:
                                    result["Преподаватели"][teacher][question] = []
                                    result["Преподаватели"][teacher][question].append(answers[answer_index])
                                answer_index += 1
        return result

