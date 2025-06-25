from flask import Flask, request, send_file, render_template
from docx import Document
import io
import re
import os
import logging

app = Flask(__name__)

# Настройка логирования
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

@app.route('/')
def form():
    logger.info("Открыта страница формы")
    try:
        return render_template('index.html')
    except Exception as e:
        logger.error(f"Ошибка при рендеринге формы: {str(e)}")
        return "Ошибка загрузки формы", 500

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
        text_fields = {
            'security_classification': '_____________________',
            'copy_number': '_________',
            'executive_authority_head': '_______________________________________',
            'executive_authority_name': '___________________',
            'approval_date': '"__" _______________ 20__ г.',
            'security_agency_head': '______________________________________',
            'security_agency_name': '_________________',
            'security_agency_date': '"__" _______________ 20__ г.',
            'mvd_head': '_______________________________________',
            'mvd_name': '___________________',
            'mvd_date': '"__" _______________ 20__ г.',
            'mchs_head': '______________________________________',
            'mchs_name': '_________________',
            'mchs_date': '"__" _______________ 20__ г.',
            'rosgvardia_head': '____________________________________',
            'rosgvardia_name': '__________________',
            'rosgvardia_date': '"__" _______________ 20__ г.',
            'locality': '___________________________________________',
            'object_name': '___________________________________________________________________________',
            'object_address': '___________________________________________________________________________',
            'object_affiliation': '___________________________________________________________________________',
            'object_boundaries': '___________________________________________________________________________',
            'object_area_perimeter': '___________________________________________________________________________',
            'monitoring_results': '___________________________________________________________________________',
            'object_category': '___________________________________________________________________________',
            'mvd_territory': '___________________________________________________________________________',
            'public_organizations': '___________________________________________________________________________',
            'terrain_characteristics': '___________________________________________________________________________',
            'staff_count': '___________________________________________________________________________',
            'attendance': '___________________________________________________________________________',
            'tenants_info': '___________________________________________________________________________',
            'illegal_actions_a': '___________________________________________________________________',
            'diversion_manifestations_b': '____________________________________________________________________',
            'security_forces_a': '___________________________________________________________________',
            'patrol_routes_b': '___________________________________________________________________',
            'stationary_posts_b': '___________________________________________________________________',
            'public_guards_d': '___________________________________________________________________',
            'security_equipment_e': '__________________________________________________________________',
            'notification_system_zh': '___________________________________________________________________________',
            'notification_system_zh_2': '___________________________________________________________________________',
            'notification_system_zh_3': '___________________________________________________________________________',
            'notification_system_zh_4': '___________________________________________________________________________',
            'notification_system_zh_5': '___________________________________________________________________________',
            'notification_system_zh_6': '___________________________________________________________________________',
            'technical_security_a': '__________________________________________________________________',
            'fire_safety_b': '__________________________________________________________________',
            'evacuation_system_v': '___________________________________________________________________________',
            'security_reliability_a': '___________________________________________________________________',
            'urgent_measures_b': '___________________________________________________________________',
            'funding_v': '____________________________________________________________________',
            'additional_info': '___________________________________________________________________________',
            'recreation_areas': '___________________________________________________________________________',
            'communication_schemes': '___________________________________________________________________________',
            'evacuation_instructions': '___________________________________________________________________________',
            'correction_log': '___________________________________________________________________________',
            'rights_holder': '___________________________________________________________________________',
            'rights_holder_name': '________________________________ __________________________________________',
            'creation_date': 'Составлен "__" ____________ 20__ г.',
            'update_date': 'Актуализирован "__" _________ 20__ г.'
        }

        # Собираем данные из формы
        fields = {key: request.form.get(key, '').strip() for key in text_fields}

        # Таблицы
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

        # Загружаем шаблон
        doc = Document(template_path)
        logger.info("Шаблон успешно загружен")

        # Функция замены текстовых заполнителей с сохранением структуры
        def replace_placeholders(doc, fields):
            for para in doc.paragraphs:
                original_text = para.text.strip()
                for key, placeholder in text_fields.items():
                    if placeholder in original_text:
                        value = fields.get(key, '').strip()
                        if value:
                            para.text = original_text.replace(placeholder, value)
                        else:
                            para.text = original_text  # Сохраняем подчёркивания, если данных нет
                        para.alignment = 1  # Выравнивание по центру
                        logger.info(f"Замена в параграфе: '{original_text}' -> '{para.text}' (ключ: {key})")
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            original_text = para.text.strip()
                            for key, placeholder in text_fields.items():
                                if placeholder in original_text:
                                    value = fields.get(key, '').strip()
                                    if value:
                                        para.text = original_text.replace(placeholder, value)
                                    else:
                                        para.text = original_text  # Сохраняем подчёркивания
                                    para.alignment = 1  # Выравнивание по центру
                                    logger.info(f"Замена в таблице: '{original_text}' -> '{para.text}' (ключ: {key})")

        # Замена текстовых заполнителей
        replace_placeholders(doc, fields)

        # Функция обновления таблиц
        def update_table(table_index, field_data, column_count, has_number_column=True):
            try:
                if table_index >= len(doc.tables):
                    logger.error(f"Таблица с индексом {table_index} не найдена")
                    return
                table = doc.tables[table_index]
                expected_columns = column_count + (1 if has_number_column else 0)
                logger.info(f"Обновление таблицы {table_index}, столбцов: {expected_columns}")

                if len(table.rows) == 0 or len(table.rows[0].cells) != expected_columns:
                    logger.error(f"Таблица {table_index} имеет {len(table.rows[0].cells) if table.rows else 0} столбцов, ожидается {expected_columns}")
                    return

                # Очищаем строки, кроме заголовка
                while len(table.rows) > 1:
                    table._element.remove(table.rows[-1]._element)

                # Добавляем строки только если есть данные
                row_count = max(len(field_data[key]) for key in field_data if field_data[key]) if any(field_data[key] for key in field_data) else 0
                if row_count == 0:
                    logger.info(f"Нет данных для таблицы {table_index}, строка не добавлена")
                    return

                for i in range(row_count):
                    row = table.add_row()
                    cell_offset = 1 if has_number_column else 0
                    if has_number_column:
                        row.cells[0].text = str(i + 1)
                        for para in row.cells[0].paragraphs:
                            para.alignment = 1  # Выравнивание по центру
                    for j, key in enumerate(field_data.keys()):
                        if j < column_count:
                            value = field_data[key][i] if i < len(field_data[key]) and field_data[key][i] else ''
                            row.cells[j + cell_offset].text = value
                            for para in row.cells[j + cell_offset].paragraphs:
                                para.alignment = 1  # Выравнивание по центру

            except Exception as e:
                logger.error(f"Ошибка при обновлении таблицы {table_index}: {str(e)}")
                raise

        # Обновляем таблицы
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
