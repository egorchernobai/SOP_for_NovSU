import requests
import json

# URL развернутого Google Apps Script
SCRIPT_URL = "https://script.google.com/macros/s/AKfycbzTCiZu5cH2b10o1TBrURP7wXUgDZpTakWJwQdaRWCnys-e6kl7W0l6e-A-am2x2iP41Q/exec"

def create_google_form(form_title, subjects, questions1, questions2):
    payload = {
        "formTitle": form_title,
        "json1": subjects,
        "json2": questions1,
        "json3": questions2
    }
    headers = {"Content-Type": "application/json"}

    
    response = requests.post(SCRIPT_URL, json=payload, headers=headers, timeout=30)
        
    if response.status_code == 200:
        print("Google Form создана:", response.text)
    else:
        print("Ошибка сервера:", response.text)



# Запуск
create_google_form(
    "Тестовая форма",
    {
        "Математика": ["Учитель 1", "Учитель 2"],
        "Физика": ["Учитель 3", "Учитель 4"]
    },
    {
        "Полезность курса для вашей будущей карьеры:": ["0", "1", "2", "3", "4", "5", "Затрудняюсь ответить", 1],
        "Бесполезность курса для вашей будущей карьеры:": [0]
    },
    {
        "Полезность курса для вашей будущей карьеы:": ["0", "1", "2", "3", "4", "5", "Затрудняюсь ответить", 1],
        "Бесполезность курса для вашей будущей карьеры:": [1]
    }
)
