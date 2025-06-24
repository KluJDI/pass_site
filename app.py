from flask import Flask, request, send_file, render_template
from docx import Document
import io
import re

app = Flask(__name__)

@app.route('/')
def form():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate():
    # Собираем все данные из формы
    fields = {key: value for key, value in request.form.items()}

    # Открываем шаблон документа
    doc = Document("pasport_mesta_massovogo_prebuvaniya_lyudei.docx")

    # Функция для замены подчеркиваний в тексте на значения
    def replace_placeholders(paragraphs):
        for para in paragraphs:
            for key, val in fields.items():
                # Заменяем длинные подчеркивания на введенные значения
                para.text = re.sub(r'_+', val, para.text, count=1)

    # Заменяем в абзацах
    replace_placeholders(doc.paragraphs)

    # Заменяем в таблицах
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_placeholders(cell.paragraphs)

    # Отправляем документ пользователю
    output = io.BytesIO()
    doc.save(output)
    output.seek(0)

    return send_file(output, download_name="Паспорт_безопасности.docx", as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
