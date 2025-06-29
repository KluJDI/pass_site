from flask import Flask, request, send_file, render_template
from docx import Document
import io
import os
import logging
import re

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

        # Таблицы (синхронизированы с index.html)
        table_fields = {
            'objects_on_territory': {
                'num': request.form.getlist('object_on_territory_num[]') if 'object_on_territory_num[]' in request.form else [''],
                'name': request.form.getlist('object_on_territory_name[]'),
                'details': request.form.getlist('object_on_territory_details[]'),
                'location': request.form.getlist('object_on_territory_location[]'),
                'security': request.form.getlist('object_on_territory_security[]')
            },
            'objects_nearby': {
                'num': request.form.getlist('object_nearby_num[]') if 'object_nearby_num[]' in request.form else [''],
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
        logger.info(f"Шаблон успешно загружен. Количество таблиц: {len(doc.tables)}")

        # Функция для замены текстовых плейсхолдеров
        def replace_text_placeholders(doc, fields):
            def normalize_text(text):
                return re.sub(r'\s+', ' ', text.strip()) if text else ''

            for para in doc.paragraphs:
                for key, placeholder in text_fields.items():
                    normalized_placeholder = normalize_text(placeholder)
                    for run in para.runs:
                        normalized_run_text = normalize_text(run.text)
                        if normalized_placeholder in normalized_run_text:
                            value = fields.get(key, '').strip()
                            run.text = run.text.replace(placeholder, value if value else "")
                            para.alignment = 1  # Выравнивание по центру
                            logger.info(f"Замена в параграфе: '{placeholder}' -> '{value}'")

            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            for key, placeholder in text_fields.items():
                                normalized_placeholder = normalize_text(placeholder)
                                for run in para.runs:
                                    normalized_run_text = normalize_text(run.text)
                                    if normalized_placeholder in normalized_run_text:
                                        value = fields.get(key, '').strip()
                                        run.text = run.text.replace(placeholder, value if value else "")
                                        para.alignment = 1  # Выравнивание по центру
                                        logger.info(f"Замена в таблице: '{placeholder}' -> '{value}'")

        # Функция для заполнения таблиц
        def fill_tables(doc, table_fields):
            table_map = {
                'objects_on_territory': {'index': 0, 'fields': ['num', 'name', 'details', 'location', 'security'], 'expected_columns': 5},
                'objects_nearby': {'index': 1, 'fields': ['num', 'name', 'details', 'side', 'distance'], 'expected_columns': 5},
                'transport': {'index': 2, 'fields': ['type', 'name', 'distance'], 'expected_columns': 3},
                'service_orgs': {'index': 3, 'fields': ['name', 'activity', 'schedule'], 'expected_columns': 3},
                'dangerous_sections': {'index': 4, 'fields': ['name', 'workers', 'risk'], 'expected_columns': 3},
                'consequences': {'index': 5, 'fields': ['name', 'victims', 'scale'], 'expected_columns': 3},
                'patrol_composition': {'index': 6, 'fields': ['type', 'units', 'people'], 'expected_columns': 3},
                'critical_elements': {'index': 7, 'fields': ['name', 'requirements', 'physical_protection', 'terrorism_prevention', 'sufficiency', 'compensation'], 'expected_columns': 6}
            }

            for table_key, table_info in table_map.items():
                table_index = table_info['index']
                expected_columns = table_info['expected_columns']
                if table_index >= len(doc.tables):
                    logger.error(f"Таблица {table_key} с индексом {table_index} не найдена. Всего таблиц: {len(doc.tables)}")
                    continue

                table = doc.tables[table_index]
                if len(table.columns) != expected_columns:
                    logger.error(f"Таблица {table_key} имеет {len(table.columns)} столбцов, ожидалось {expected_columns}")
                    continue

                logger.info(f"Обработка таблицы {table_key} (индекс: {table_index}, столбцов: {len(table.columns)})")
                data = table_fields[table_key]
                row_count = max(len(data[field]) for field in data if data[field]) if any(data[field] for field in data) else 0

                if table_key == 'patrol_composition':
                    row_count = min(row_count, 5)  # Ограничиваем до 5 строк
                    total_units, total_persons = 0, 0

                # Очищаем существующие строки (кроме заголовка)
                while len(table.rows) > 1:
                    table._element.remove(table.rows[-1]._element)

                # Добавляем новые строки
                for i in range(row_count):
                    row = table.add_row()
                    cells = row.cells
                    if len(cells) != expected_columns:
                        logger.error(f"Таблица {table_key}: новая строка имеет {len(cells)} столбцов, ожидалось {expected_columns}")
                        continue

                    for j, field in enumerate(table_info['fields']):
                        value = data[field][i] if i < len(data[field]) else ""
                        cells[j].text = str(value) if value else ""
                        for para in cells[j].paragraphs:
                            para.alignment = 1  # Выравнивание по центру
                        logger.info(f"Добавлена строка в таблицу {table_key}: {field} = {value}")

                    if table_key == 'patrol_composition':
                        try:
                            total_units += int(data['units'][i]) if data['units'][i] else 0
                            total_persons += int(data['people'][i]) if data['people'][i] else 0
                        except ValueError:
                            logger.warning(f"Некорректное значение для units или people в таблице {table_key}, строка {i}")

                # Для таблицы patrol_composition добавляем итоговую строку
                if table_key == 'patrol_composition':
                    row = table.add_row()
                    cells = row.cells
                    if len(cells) != expected_columns:
                        logger.error(f"Таблица {table_key}: итоговая строка имеет {len(cells)} столбцов, ожидалось {expected_columns}")
                        continue
                    cells[0].text = "Всего"
                    cells[1].text = str(total_units)
                    cells[2].text = str(total_persons)
                    for para in cells[0].paragraphs + cells[1].paragraphs + cells[2].paragraphs:
                        para.alignment = 1
                    logger.info(f"Добавлена итоговая строка в таблицу {table_key}: Всего = {total_units} единиц, {total_persons} человек")

        # Загружаем шаблон
        doc = Document(template_path)
        logger.info(f"Шаблон успешно загружен. Количество таблиц: {len(doc.tables)}")

        # Заменяем текстовые плейсхолдеры
        replace_text_placeholders(doc, fields)
        # Заполняем таблицы
        fill_tables(doc, table_fields)

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
