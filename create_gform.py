import requests
import json


class Gform():
    def create_google_form(url, form_title, description, subjects, questions1, questions2):
        payload = {
            "formTitle": form_title,
            "description": description,
            "json1": subjects,
            "json2": questions1,
            "json3": questions2
        }
        headers = {"Content-Type": "application/json"}

        response = requests.post(
            url, json=payload, headers=headers, timeout=180)

        if response.status_code == 200:
            return response.text
        else:
            return 0
