# SOP_for_NovSU
Implementation of the HSE "Студенческая оценка преподавания" System for NovSU

# Структура
- "1.csv" - Файл с ответами формы гугла
- "create_gform.py" - Создает google form
- "excel_create.py" - Получает словарь из "parser_csv.py" и создает/изменяет excel файл с оценками
- "google.js" - Создает google form(google script app)
- "main.py" - Главный файл работы программы
- "Mainwindow.ui" - Qt designer файл gui
- "parser_csv.py" - Создает python словарь из файла ответов на форму
- "settings.json" - Файл настроек программы
- "ui.py" - Файл GUI