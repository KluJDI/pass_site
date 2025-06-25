```python
from flask import Flask, request, send_file, render_template
from docx import Document
import io
import re
import os
import logging

app = Flask(__name__)

# Настройка логирования
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

@app.route('/')
def form():
    logger.info("Открыта страница формы")
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate():
    try:
        # Путь к шаблону
        template_path = os.path.join(os.path.dirname(__file__), "pasport_mesta_massovogo_prebuvaniya_lyudei.docx")
        logger.info(f"Проверка шаблона по пути: {template_path}")

        # Проверка существования шаблона
        if not os.path.exists(template_path):
            logger.error(f"Шаблон {template_path} не найден")
            return "Ошибка: Шаблон документа не найден", 500

        # Собираем данные из формы
        fields = {key: value for key, value in request.form.items() if not key.endswith('[]')}
        table_fields = {
            'objects_on_territory': {
                'name': request.form.getlist('object_on_territory_name[]'),
                'details': request.form.getlist('object_on_territory_details[]'),
                'location': request.form.getlist('object_on_territory_location[]'),
                'security': request.form.getlist('object_on_territory_security[]')
            },
            'objects_nearby': {
                'name': request.form.getlist('object_nearby_name[]'),
                'details': request.form.getlist('object_nearby_details[]'),
                'side': request.form.getlist('object_nearby_side[]'),
                'distance': request.form.getlist('object_nearby_distance[]')
            },
            'transport': {
                'type': request.form.getlist('transport_type[]'),
                'name': request.form.getlist('transport_name[]'),
                'distance': request.form.getlist('transport_distance[]')
            },
            'service_orgs': {
                'name': request.form.getlist('service_org_name[]'),
                'activity': request.form.getlist('service_org_activity[]'),
                'schedule': request.form.getlist('service_org_schedule[]')
            },
            'dangerous_sections': {
                'name': request.form.getlist('dangerous_section_name[]'),
                'workers': request.form.getlist('dangerous_section_workers[]'),
                'risk': request.form.getlist('dangerous_section_risk[]')
            },
            'consequences': {
                'name': request.form.getlist('threat_name[]'),
                'victims': request.form.getlist('threat_victims[]'),
                'scale': request.form.getlist('threat_scale[]')
            },
            'patrol_composition': {
                'type': request.form.getlist('patrol_type[]'),
                'units': request.form.getlist('patrol_units[]'),
                'people': request.form.getlist('patrol_people[]')
            },
            'critical_elements': {
                'name': request.form.getlist('critical_element_name[]'),
                'requirements': request.form.getlist('critical_element_requirements[]'),
                'physical_protection': request.form.getlist('critical_element_physical_protection[]'),
                'terrorism_prevention': request.form.getlist('critical_element_terrorism_prevention[]'),
                'sufficiency': request.form.getlist('critical_element_sufficiency[]'),
                'compensation': request.form.getlist('critical_element_compensation[]')
            }
        }

        logger.debug(f"Данные формы (поля): {fields}")
        logger.debug(f"Данные формы (таблицы): {table_fields}")

        # Проверка, что данные для таблиц не пустые
        for table_name, data in table_fields.items():
            if not data[list(data.keys())[0]]:
                logger.warning(f"Данные для таблицы {table_name} пустые, пропускаем")
                table_fields[table_name] = {key: [''] for key in data.keys()}  # Заполняем пустыми значениями

        # Загружаем шаблон
        doc = Document(template_path)
        logger.info("Шаблон успешно загружен")

        # Функция замены заполнителей
        def replace_placeholders(paragraphs):
            for para in paragraphs:
                original_text = para.text
                for key, val in fields.items():
                    para.text = re.sub(r'_+', val or '', para.text, count=1)
                if para.text != original_text:
                    logger.debug(f"Замена в параграфе: '{original_text}' -> '{para.text}'")

        # Замена в параграфах
        replace_placeholders(doc.paragraphs)

        # Замена в таблицах
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    replace_placeholders(cell.paragraphs)

        # Обновление таблиц
        def update_table(table_index, field_data, column_count):
            try:
                if table_index >= len(doc.tables):
                    logger.error(f"Таблица с индексом {table_index} не найдена в документе")
                    return
                table = doc.tables[table_index]
                logger.info(f"Обновление таблицы {table_index}, ожидаемое количество столбцов: {column_count + 1}")

                # Проверка количества столбцов
                if len(table.rows[0].cells) < column_count + 1:
                    logger.error(f"Таблица {table_index} имеет {len(table.rows[0].cells)} столбцов, ожидается {column_count + 1}")
                    return

                # Очищаем строки, кроме заголовка
                if len(table.rows) > 1:
                    for _ in range(len(table.rows) - 1):
                        table._element.remove(table.rows[-1]._element)

                # Добавляем новые строки
                row_count = len(field_data[list(field_data.keys())[0]])
                if row_count == 0:
                    logger.info(f"Нет данных для таблицы {table_index}, оставляем пустой")
                    return

                for i in range(row_count):
                    row = table.add_row()
                    row.cells[0].text = str(i + 1)  # Номер строки
                    for j, key in enumerate(field_data.keys()):
                        if j < column_count and i < len(field_data[key]):
                            row.cells[j + 1].text = field_data[key][i] or ''
                            logger.debug(f"Таблица {table_index}, строка {i}, столбец {j+1}: {field_data[key][i]}")
            except Exception as e:
                logger.error(f"Ошибка при обновлении таблицы {table_index}: {str(e)}")
                raise

        # Обновляем таблицы
        update_table(0, table_fields['objects_on_territory'], 4)  # Таблица 2
        update_table(1, table_fields['objects_nearby'], 4)  # Таблица 3
        update_table(2, table_fields['transport'], 3)  # Таблица 4
        update_table(3, table_fields['service_orgs'], 3)  # Таблица 5
        update_table(4, table_fields['dangerous_sections'], 3)  # Таблица 7
        update_table(5, table_fields['consequences'], 3)  # Таблица 9
        update_table(6, table_fields['patrol_composition'], 3)  # Таблица 10г
        update_table(7, table_fields['critical_elements'], 6)  # Таблица 12

        # Сохраняем документ
        output = io.BytesIO()
        doc.save(output)
        output.seek(0)
        logger.info("Документ успешно сохранен")

        return send_file(output, download_name="Паспорт_безопасности.docx", as_attachment=True)

    except Exception as e:
        logger.error(f"Ошибка при генерации документа: {str(e)}")
        return f"Ошибка при генерации документа: {str(e)}", 500

import os

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
