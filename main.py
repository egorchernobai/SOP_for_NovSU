from PyQt6.QtWidgets import QApplication, QMainWindow, QMessageBox
from ui import Ui_MainWindow  # Ваш сгенерированный UI класс
import json
from PyQt6.QtCore import Qt
import create_gform
import webbrowser
from PyQt6.QtCore import QThread, pyqtSignal
import asyncio
from parser_csv import parser

class GFormWorker(QThread):
    finished = pyqtSignal(str)  # Сигнал с URL формы или сообщением об ошибке

    def run(self):
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        result = loop.run_until_complete(self.create_form())
        self.finished.emit(result)

    async def create_form(self):
        with open("settings.json", "r", encoding="utf-8") as settings_file:
            data = json.load(settings_file)

        return create_gform.Gform.create_google_form(
            data["Url"], data["Name_form"], data["Desciption_form"],
            data["Subjects"], data["Questions_for_subject"], data["Questions_for_teachers"]
        )

class MainWindow(QMainWindow):
    def __init__(self):
        with open("settings.json", "r", encoding="utf-8") as settings_file:
            data = json.load(settings_file)
        super().__init__()
        self.ui = Ui_MainWindow()
        self.data = data
        self.ui.setupUi(self)  # Передаём экземпляр QMainWindow
        self.ui.url_script.setText(data['Url'])
        self.ui.name_form.setText(data['Name_form'])
        self.ui.description_form.setText(data['Desciption_form'])
        if(data["Subjects"]):
            self.ui.subjects.setText("\"" + "\", \"".join(data['Subjects'].keys()) + "\"")
        
        for subject in data['Subjects'].keys():
            self.ui.subjects_combobox.addItem(subject)
        
        if(self.ui.subjects_combobox.currentText()):
            if(data['Subjects'][self.ui.subjects_combobox.currentText()]!= []):
                self.ui.teachers.setText("\"" + "\", \"".join(data['Subjects'][self.ui.subjects_combobox.currentText()]) + "\"")
        
        if(data["Questions_for_subject"]):
            self.ui.questions_for_subject.setText("\"" + "\", \"".join(data['Questions_for_subject'].keys()) + "\"")

        for question in data['Questions_for_subject'].keys():
            self.ui.questions_for_subject_combobox.addItem(question)

        if(self.ui.questions_for_subject_combobox.currentText()):
            if(len(self.data['Questions_for_subject'][self.ui.questions_for_subject_combobox.currentText()]) > 1):
                self.ui.variants_for_questions.setText("\"" + "\", \"".join(self.data['Questions_for_subject'][self.ui.questions_for_subject_combobox.currentText()][0:-1]) + "\"")
                if(self.data['Questions_for_subject'][self.ui.questions_for_subject_combobox.currentText()][-1]):
                    self.ui.isrequired_subject.setChecked(True)
                else:
                    self.ui.isrequired_subject.setChecked(False)
            else:
                if(self.data['Questions_for_subject'][self.ui.questions_for_subject_combobox.currentText()][-1]):
                    self.ui.isrequired_subject.setChecked(True)
                else:
                    self.ui.isrequired_subject.setChecked(False)
                self.ui.variants_for_questions.setText("")
        
        if(data["Questions_for_teachers"]):
            self.ui.questions_for_teachers.setText("\"" + "\", \"".join(data['Questions_for_teachers'].keys()) + "\"")

        for question in data['Questions_for_teachers'].keys():
            self.ui.questions_for_teachers_combobox.addItem(question)

        if(self.ui.questions_for_teachers_combobox.currentText()):
            self.ui.variants_for_teacher.setText("\"" + "\", \"".join(data['Questions_for_teachers'][self.ui.questions_for_teachers_combobox.currentText()][0:-1]) + "\"")
            if(data['Questions_for_teachers'][self.ui.questions_for_teachers_combobox.currentText()][-1]):
                self.ui.isrequired_subject_2.setChecked(True)

        self.ui.url_script.textChanged.connect(self.change_url)
        self.ui.name_form.textChanged.connect(self.change_name_form)
        self.ui.description_form.textChanged.connect(self.change_description_form)
        self.ui.subjects.textChanged.connect(self.change_subjects_form)
        self.ui.subjects_combobox.currentIndexChanged.connect(self.change_teachers)
        self.ui.teachers.textChanged.connect(self.change_fio_form)
        self.ui.questions_for_subject.textChanged.connect(self.change_quest_subj_form)
        self.ui.questions_for_subject_combobox.currentIndexChanged.connect(self.change_variants_for_questions)
        self.ui.isrequired_subject.stateChanged.connect(self.change_isrequired_subject)
        self.ui.variants_for_questions.textChanged.connect(self.change_otv_subj_form)

        self.ui.questions_for_teachers.textChanged.connect(self.change_quest_teacher_form)
        self.ui.questions_for_teachers_combobox.currentIndexChanged.connect(self.change_variants_for_teacher_questions)
        self.ui.isrequired_subject_2.stateChanged.connect(self.change_isrequired_subject_2)
        self.ui.variants_for_teacher.textChanged.connect(self.change_otv_teacher_form)

        self.ui.finish_button.clicked.connect(self.start_creating_gform)



    def change_url(self):
        self.data["Url"] = self.ui.url_script.text()
        with open("settings.json", "w", encoding="utf-8") as settings_file_write:
            json.dump(self.data, settings_file_write, ensure_ascii=False)
    
    def change_name_form(self):
        self.data["Name_form"] = self.ui.name_form.text()
        with open("settings.json", "w", encoding="utf-8") as settings_file_write:
            json.dump(self.data, settings_file_write, ensure_ascii=False)

    def change_description_form(self):
        self.data["Desciption_form"] = self.ui.description_form.text()
        with open("settings.json", "w", encoding="utf-8") as settings_file_write:
            json.dump(self.data, settings_file_write, ensure_ascii=False)

    def change_subjects_form(self):
        lines = self.ui.subjects.toPlainText()[1:-1].split("\", \"")
        subjects_dict = {}
        for line in lines:
            if line != "":
                subjects_dict[line] = []
        self.data["Subjects"] = subjects_dict
        with open("settings.json", "w", encoding="utf-8") as settings_file_write:
            json.dump(self.data, settings_file_write, ensure_ascii=False)

        self.ui.subjects_combobox.clear()
        for subject in self.data['Subjects'].keys():
            self.ui.subjects_combobox.addItem(subject)

    def change_teachers(self):
        if(self.ui.subjects_combobox.currentText()):
            if(self.data['Subjects'][self.ui.subjects_combobox.currentText()] != []):
                self.ui.teachers.setText("\"" + "\", \"".join(self.data['Subjects'][self.ui.subjects_combobox.currentText()]) + "\"")
            else:
                self.ui.teachers.setText("")

    def change_fio_form(self):
        if(self.ui.teachers.toPlainText()[1:-1]):
            self.data["Subjects"][self.ui.subjects_combobox.currentText()] = self.ui.teachers.toPlainText()[1:-1].split("\", \"")
            with open("settings.json", "w", encoding="utf-8") as settings_file_write:
                json.dump(self.data, settings_file_write, ensure_ascii=False)
        else:
            self.data["Subjects"][self.ui.subjects_combobox.currentText()] = []
            with open("settings.json", "w", encoding="utf-8") as settings_file_write:
                json.dump(self.data, settings_file_write, ensure_ascii=False)

    def change_quest_subj_form(self):
        lines = self.ui.questions_for_subject.toPlainText()[1:-1].split("\", \"")
        subjects_dict = {}
        for line in lines:
            if line:
                subjects_dict[line] = [0]
        self.data["Questions_for_subject"] = subjects_dict
        with open("settings.json", "w", encoding="utf-8") as settings_file_write:
            json.dump(self.data, settings_file_write, ensure_ascii=False)
        
        self.ui.questions_for_subject_combobox.clear()
        for subject in self.data['Questions_for_subject'].keys():
            self.ui.questions_for_subject_combobox.addItem(subject)

    def change_variants_for_questions(self):
        if(self.ui.questions_for_subject_combobox.currentText()):
            if(len(self.data['Questions_for_subject'][self.ui.questions_for_subject_combobox.currentText()]) > 1):
                self.ui.variants_for_questions.setText("\"" + "\", \"".join(self.data['Questions_for_subject'][self.ui.questions_for_subject_combobox.currentText()][0:-1]) + "\"")
                if(self.data['Questions_for_subject'][self.ui.questions_for_subject_combobox.currentText()][-1]):
                    self.ui.isrequired_subject.setChecked(True)
                else:
                    self.ui.isrequired_subject.setChecked(False)
            else:
                if(self.data['Questions_for_subject'][self.ui.questions_for_subject_combobox.currentText()][-1]):
                    self.ui.isrequired_subject.setChecked(True)
                else:
                    self.ui.isrequired_subject.setChecked(False)
                self.ui.variants_for_questions.setText("")

    def change_isrequired_subject(self, state):
        if(state == 2):
            self.data["Questions_for_subject"][self.ui.questions_for_subject_combobox.currentText()][-1] = 1
        else:
            self.data["Questions_for_subject"][self.ui.questions_for_subject_combobox.currentText()][-1] = 0
        with open("settings.json", "w", encoding="utf-8") as settings_file_write:
            json.dump(self.data, settings_file_write, ensure_ascii=False)

    def change_otv_subj_form(self):
        state = self.data["Questions_for_subject"][self.ui.questions_for_subject_combobox.currentText()][-1]
        if(self.ui.variants_for_questions.toPlainText()[1:-1]):
            self.data["Questions_for_subject"][self.ui.questions_for_subject_combobox.currentText()] = self.ui.variants_for_questions.toPlainText()[1:-1].split("\", \"")
            self.data["Questions_for_subject"][self.ui.questions_for_subject_combobox.currentText()].append(state)
            with open("settings.json", "w", encoding="utf-8") as settings_file_write:
                json.dump(self.data, settings_file_write, ensure_ascii=False)
        else:
            self.data["Questions_for_subject"][self.ui.questions_for_subject_combobox.currentText()] = []
            self.data["Questions_for_subject"][self.ui.questions_for_subject_combobox.currentText()].append(state)
            with open("settings.json", "w", encoding="utf-8") as settings_file_write:
                json.dump(self.data, settings_file_write, ensure_ascii=False)

    def change_quest_teacher_form(self):
        lines = self.ui.questions_for_teachers.toPlainText()[1:-1].split("\", \"")
        subjects_dict = {}
        for line in lines:
            if line:
                subjects_dict[line] = [0]
        self.data["Questions_for_teachers"] = subjects_dict
        with open("settings.json", "w", encoding="utf-8") as settings_file_write:
            json.dump(self.data, settings_file_write, ensure_ascii=False)
        
        self.ui.questions_for_teachers_combobox.clear()
        for subject in self.data['Questions_for_teachers'].keys():
            self.ui.questions_for_teachers_combobox.addItem(subject)

    def change_variants_for_teacher_questions(self):
        if(self.ui.questions_for_teachers_combobox.currentText()):
            if(len(self.data['Questions_for_teachers'][self.ui.questions_for_teachers_combobox.currentText()]) > 1):
                self.ui.variants_for_teacher.setText("\"" + "\", \"".join(self.data['Questions_for_teachers'][self.ui.questions_for_teachers_combobox.currentText()][0:-1]) + "\"")
                if(self.data['Questions_for_teachers'][self.ui.questions_for_teachers_combobox.currentText()][-1]):
                    self.ui.isrequired_subject_2.setChecked(True)
                else:
                    self.ui.isrequired_subject_2.setChecked(False)
            else:
                if(self.data['Questions_for_teachers'][self.ui.questions_for_teachers_combobox.currentText()][-1]):
                    self.ui.isrequired_subject_2.setChecked(True)
                else:
                    self.ui.isrequired_subject_2.setChecked(False)
                self.ui.variants_for_teacher.setText("")

    def change_isrequired_subject_2(self, state):
        if(state == 2):
            self.data["Questions_for_teachers"][self.ui.questions_for_teachers_combobox.currentText()][-1] = 1
        else:
            self.data["Questions_for_teachers"][self.ui.questions_for_teachers_combobox.currentText()][-1] = 0
        with open("settings.json", "w", encoding="utf-8") as settings_file_write:
            json.dump(self.data, settings_file_write, ensure_ascii=False)

    def change_otv_teacher_form(self):
        state = self.data["Questions_for_teachers"][self.ui.questions_for_teachers_combobox.currentText()][-1]
        if(self.ui.variants_for_teacher.toPlainText()[1:-1]):
            self.data["Questions_for_teachers"][self.ui.questions_for_teachers_combobox.currentText()] = self.ui.variants_for_teacher.toPlainText()[1:-1].split("\", \"")
            self.data["Questions_for_teachers"][self.ui.questions_for_teachers_combobox.currentText()].append(state)
            with open("settings.json", "w", encoding="utf-8") as settings_file_write:
                json.dump(self.data, settings_file_write, ensure_ascii=False)
        else:
            self.data["Questions_for_teachers"][self.ui.questions_for_teachers_combobox.currentText()] = []
            self.data["Questions_for_teachers"][self.ui.questions_for_teachers_combobox.currentText()].append(state)
            with open("settings.json", "w", encoding="utf-8") as settings_file_write:
                json.dump(self.data, settings_file_write, ensure_ascii=False)
        
    def start_creating_gform(self):
        """Запускает создание Google Forms в отдельном потоке."""
        self.ui.statusBar.showMessage("Создание формы...", 180000)

        self.worker = GFormWorker()
        self.worker.finished.connect(self.on_form_created)
        self.worker.start()

    def on_form_created(self, url_gform):
        """Вызывается после завершения работы потока GFormWorker."""
        if "http" in url_gform:
            self.ui.statusBar.showMessage("Форма создана", 5000)
            dlg = QMessageBox(self)
            dlg.setWindowTitle("Форма создана")
            dlg.setText("Открыть её?")
            dlg.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            dlg.setIcon(QMessageBox.Icon.Question)
            button = dlg.exec()

            if button == QMessageBox.StandardButton.Yes:
                webbrowser.open(url_gform)
        else:
            self.ui.statusBar.showMessage("Ошибка создания формы", 5000)
            dlg = QMessageBox(self)
            dlg.setWindowTitle("Ошибка")
            dlg.setText(url_gform)
            dlg.setStandardButtons(QMessageBox.StandardButton.Ok)
            dlg.exec()


def parsers(path):
    return parser.parse(path)

app = QApplication([])
window = MainWindow()
window.show()
app.exec()