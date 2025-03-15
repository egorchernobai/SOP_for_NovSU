import create_gform
import json


with open("settings.json", "r", encoding="utf-8") as settings_file:
    data = json.load(settings_file)


create_gform.Gform.create_google_form(data["Url"], data["Name_form"], data["Desciption_form"], data["Subjects"], data["Questions_for_subject"], data["Questions_for_teachers"])