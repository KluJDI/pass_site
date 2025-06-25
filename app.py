from flask import Flask, request, send_file, render_template
from docx import Document
import io
import re
import os
import logging

app = Flask(__name__)

# Настройка логирования
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
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

        if not os.path.exists(template_path):
            logger.error(f"Шаблон {template_path} не найден")
            return "Ошибка: Шаблон документа не найден", 500

        # Текстовые поля
        text_fields_order = [
            'security_classification', 'copy_number',
            'executive_authority_head', 'executive_authority_name', 'approval_date',
            'security_agency_head', 'security_agency_name', 'security_agency_date',
            'mvd_head', 'mvd_name', 'mvd_date',
            'mchs_head', 'mchs_name', 'mchs_date',
            'rosgvardia_head', 'rosgvardia_name', 'rosgvardia_date',
            'locality',
            'object_name', 'object_address', 'object_affiliation', 'object_boundaries',
            'object_area_perimeter', 'monitoring_results', 'object_category',
            'mvd_territory', 'public_organizations', 'terrain_characteristics',
            'staff_count', 'attendance', 'tenants_info',
            'illegal_actions_a', 'diversion_manifestations_b',
            'security_forces_a', 'patrol_routes_b', 'stationary_posts_b',
            'public_guards_d', 'security_equipment_e',
            'notification_system_zh', 'notification_system_zh_2', 'notification_system_zh_3',
            'notification_system_zh_4', 'notification_system_zh_5', 'notification_system_zh_6',
            'technical_security_a', 'fire_safety_b', 'evacuation_system_v',
            'security_reliability_a', 'urgent_measures_b', 'funding_v',
            'additional_info',
            'recreation_areas', 'communication_schemes', 'evacuation_instructions', 'correction_log',
            'rights_holder', 'rights_holder_name', 'creation_date', 'update_date'
        ]

        # Собираем данные из формы
        fields = {key: request.form.get(key, '') for key in text_fields_order}
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

        logger.debug(f"Данные формы (текстовые поля): {fields}")
        for table_name, data in table_fields.items():
            logger.debug(f"Данные формы (таблица {table_name}): {data}")
            lengths = {key: len(values) for key, values in data.items()}
            logger.debug(f"Длина данных для таблицы {table_name}: {lengths}")

        # Проверка данных таблиц
        for table_name, data in table_fields.items():
            if not any(data[key] for key in data):
                logger.warning(f"Таблица {table_name} пустая")
                table_fields[table_name] = {key: [''] for key in data}
            elif not all(len(data[key]) == len(data[list(data.keys())[0]]) for key in data):
                logger.error(f"Несоответствие длины данных в таблице {table_name}")
                return f"Ошибка: Несоответствие количества данных в таблице {table_name}", 400

        # Загружаем шаблон
        doc = Document(template_path)
        logger.info("Шаблон успешно загружен")

        # Логирование структуры таблиц
        for i, table in enumerate(doc.tables):
            if table.rows:
                columns = len(table.rows[0].cells)
                headers = [cell.text.strip() for cell in table.rows[0].cells]
                logger.info(f"Таблица {i}: {columns} столбцов, заголовки: {headers}")
            else:
                logger.warning(f"Таблица {i} пустая или без строк")

        # Функция замены текстовых заполнителей
        def replace_placeholders(doc, fields):
            placeholder_index = 0
            for para in doc.paragraphs:
                original_text = para.text
                if re.search(r'_+', para.text):
                    if placeholder_index < len(text_fields_order):
                        key = text_fields_order[placeholder_index]
                        value = fields.get(key, '')
                        para.text = re.sub(r'_+', value, para.text)
                        logger.debug(f"Замена в параграфе: '{original_text}' -> '{para.text}' (ключ: {key})")
                        placeholder_index += 1
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            original_text = para.text
                            if re.search(r'_+', para.text):
                                if placeholder_index < len(text_fields_order):
                                    key = text_fields_order[placeholder_index]
                                    value = fields.get(key, '')
                                    para.text = re.sub(r'_+', value, para.text)
                                    logger.debug(f"Таблица: замена '{original_text}' -> '{para.text}' (ключ: {key})")
                                    placeholder_index += 1

        # Замена текстовых данных
        replace_placeholders(doc, fields)

        # Функция обновления таблицы
        def update_table(table_index, field_data, expected_column_count, has_number_column=True):
            try:
                if table_index >= len(doc.tables):
                    logger.error(f"Таблица с индексом {table_index} не найдена")
                    return

                table = doc.tables[table_index]
                actual_columns = len(table.rows[0].cells) if table.rows else 0
                expected_columns = expected_column_count + (1 if has_number_column else 0)
                logger.info(f"Обновление таблицы {table_index}, столбцов в шаблоне: {actual_columns}, ожидается: {expected_columns}")

                # Используем минимальное количество столбцов
                column_count = min(actual_columns - (1 if has_number_column else 0), expected_column_count)
                if column_count < 0:
                    logger.error(f"Таблица {table_index} имеет некорректное количество столбцов: {actual_columns}")
                    return

                # Очищаем строки, кроме заголовка
                while len(table.rows) > 1:
                    table._element.remove(table.rows[-1]._element)
                    logger.debug(f"Удалена строка в таблице {table_index}")

                # Добавляем строки
                row_count = max(len(field_data[key]) for key in field_data if field_data[key]) if any(field_data[key] for key in field_data) else 0
                logger.debug(f"Добавление {row_count} строк в таблицу {table_index}")
                if row_count == 0:
                    logger.warning(f"Нет данных для таблицы {table_index}, добавляем пустую строку")
                    row = table.add_row()
                    for cell in row.cells:
                        cell.text = ''
                    return

                for i in range(row_count):
                    row = table.add_row()
                    if len(row.cells) != actual_columns:
                        logger.error(f"Ошибка: новая строка в таблице {table_index} имеет {len(row.cells)} столбцов, ожидается {actual_columns}")
                        continue
                    cell_offset = 1 if has_number_column else 0
                    if has_number_column and len(row.cells) > 0:
                        row.cells[0].text = str(i + 1)
                        logger.debug(f"Таблица {table_index}, строка {i}, столбец 0: {i + 1}")
                    for j, key in enumerate(field_data.keys()):
                        if j < column_count and j + cell_offset < len(row.cells):
                            value = field_data[key][i] if i < len(field_data[key]) and field_data[key][i] else ''
                            row.cells[j + cell_offset].text = value
                            logger.debug(f"Таблица {table_index}, строка {i}, столбец {j + cell_offset}: {value} (ключ: {key})")
                        elif j + cell_offset >= len(row.cells):
                            logger.warning(f"Пропущен столбец {j + cell_offset} в таблице {table_index}, так как он превышает количество столбцов в шаблоне")
            except Exception as e:
                logger.error(f"Ошибка при обновлении таблицы {table_index}: {str(e)}")
                raise

        # Обновляем таблицы с ожидаемым количеством столбцов (без учета номера)
        update_table(0, table_fields['objects_on_territory'], 4)
        update_table(1, table_fields['objects_nearby'], 4)
        update_table(2, table_fields['transport'], 3)
        update_table(3, table_fields['service_orgs'], 3)
        update_table(4, table_fields['dangerous_sections'], 3)
        update_table(5, table_fields['consequences'], 3)
        update_table(6, table_fields['patrol_composition'], 3, has_number_column=False)
        update_table(7, table_fields['critical_elements'], 6)

        # Сохраняем документ
        output = io.BytesIO()
        doc.save(output)
        output.seek(0)
        logger.info("Документ успешно сохранен")

        return send_file(output, download_name="Паспорт_безопасности.docx", as_attachment=True)

    except Exception as e:
        logger.error(f"Ошибка при генерации документа: {str(e)}")
        return f"Ошибка при генерации документа: {str(e)}", 500

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
