import matplotlib.pyplot as plt
from collections import Counter
from docx import Document
from io import BytesIO

# Ваш массив с данными (например, ответы пользователей)
data = ['А', 'Б', 'В', 'А', 'А', 'Б', 'А', 'В', 'Г']

# Подсчитываем количество вхождений каждого значения
counter = Counter(data)

# Получаем значения и их процентное соотношение
total = len(data)
percentages = [(count / total) * 100 for count in counter.values()]

# Метки для категорий (ответов)
labels = counter.keys()

# Создаем круговую диаграмму
fig, ax = plt.subplots()
ax.pie(percentages, labels=labels, autopct='%1.1f%%', startangle=90)
ax.axis('equal')  # Равномерное распределение сегментов

# Сохраняем диаграмму в буфер
img_stream = BytesIO()
plt.savefig(img_stream, format='png')
img_stream.seek(0)

# Создание документа Word
doc = Document()
doc.add_heading('Процентное соотношение встречаемости ответов', 0)

# Добавляем изображение диаграммы в Word
doc.add_picture(img_stream)

# Сохраняем документ
doc.save('chart_in_word.docx')

# Закрываем поток изображения
img_stream.close()

print("Документ с диаграммой успешно создан!")
