from flask import Flask, request, send_file, render_template
from docx import Document
import io
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
                'num': request.form.getlist('object_on_territory_num[]'),
                'name': request.form.getlist('object_on_territory_name[]'),
                'details': request.form.getlist('object_on_territory_details[]'),
                'location': request.form.getlist('object_on_territory_location[]'),
                'security': request.form.getlist('object_on_territory_security[]')
            },
            'objects_nearby': {
                'num': request.form.getlist('object_nearby_num[]'),
                'name': request.form.getlist('object_nearby_name[]'),
                'characteristics': request.form.getlist('object_nearby_characteristics[]'),
                'location': request.form.getlist('object_nearby_location[]'),
                'distance': request.form.getlist('object_nearby_distance[]')
            },
            'transport_communications': {
                'num': request.form.getlist('transport_num[]'),
                'type': request.form.getlist('transport_type[]'),
                'name': request.form.getlist('transport_name[]'),
                'distance': request.form.getlist('transport_distance[]')
            },
            'service_organizations': {
                'num': request.form.getlist('service_org_num[]'),
                'name': request.form.getlist('service_org_name[]'),
                'activity': request.form.getlist('service_org_activity[]'),
                'schedule': request.form.getlist('service_org_schedule[]')
            },
            'dangerous_areas': {
                'num': request.form.getlist('dangerous_area_num[]'),
                'name': request.form.getlist('dangerous_area_name[]'),
                'worker_count': request.form.getlist('dangerous_area_worker_count[]'),
                'emergency_type': request.form.getlist('dangerous_area_emergency_type[]')
            },
            'terror_consequences': {
                'num': request.form.getlist('terror_consequence_num[]'),
                'threat': request.form.getlist('terror_consequence_threat[]'),
                'victims_count': request.form.getlist('terror_consequence_victims_count[]'),
                'consequence_scale': request.form.getlist('terror_consequence_consequence_scale[]')
            },
            'security_posts': {
                'post_type': request.form.getlist('security_post_type[]'),
                'units': request.form.getlist('security_post_units[]'),
                'persons': request.form.getlist('security_post_persons[]')
            },
            'protection_assessment': {
                'num': request.form.getlist('protection_assessment_num[]'),
                'element_name': request.form.getlist('protection_assessment_element_name[]'),
                'requirements': request.form.getlist('protection_assessment_requirements[]'),
                'physical_protection': request.form.getlist('protection_assessment_physical_protection[]'),
                'terror_prevention': request.form.getlist('protection_assessment_terror_prevention[]'),
                'sufficiency': request.form.getlist('protection_assessment_sufficiency[]'),
                'compensation': request.form.getlist('protection_assessment_compensation[]')
            }
        }

        # Загружаем шаблон
        doc = Document(template_path)
        logger.info(f"Шаблон успешно загружен. Количество таблиц: {len(doc.tables)}")

        # Функция для замены текстовых плейсхолдеров
        def replace_text_placeholders(doc, fields):
            for para in doc.paragraphs:
                for key, placeholder in text_fields.items():
                    if placeholder in para.text:
                        value = fields.get(key, '').strip()
                        para.text = para.text.replace(placeholder, value if value else "")
                        para.alignment = 1  # Выравнивание по центру
                        logger.info(f"Замена в параграфе: '{placeholder}' -> '{value}'")
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            for key, placeholder in text_fields.items():
                                if placeholder in para.text:
                                    value = fields.get(key, '').strip()
                                    para.text = para.text.replace(placeholder, value if value else "")
                                    para.alignment = 1  # Выравнивание по центру
                                    logger.info(f"Замена в таблице: '{placeholder}' -> '{value}'")

        # Функция для заполнения таблиц
        def fill_tables(doc, table_fields):
            table_map = {
                'objects_on_territory': {'index': 0, 'fields': ['num', 'name', 'details', 'location', 'security'], 'expected_columns': 5},
                'objects_nearby': {'index': 1, 'fields': ['num', 'name', 'characteristics', 'location', 'distance'], 'expected_columns': 5},
                'transport_communications': {'index': 2, 'fields': ['num', 'type', 'name', 'distance'], 'expected_columns': 4},
                'service_organizations': {'index': 3, 'fields': ['num', 'name', 'activity', 'schedule'], 'expected_columns': 4},
                'dangerous_areas': {'index': 4, 'fields': ['num', 'name', 'worker_count', 'emergency_type'], 'expected_columns': 4},
                'terror_consequences': {'index': 5, 'fields': ['num', 'threat', 'victims_count', 'consequence_scale'], 'expected_columns': 4},
                'security_posts': {'index': 6, 'fields': ['post_type', 'units', 'persons'], 'expected_columns': 3},
                'protection_assessment': {'index': 7, 'fields': ['num', 'element_name', 'requirements', 'physical_protection', 'terror_prevention', 'sufficiency', 'compensation'], 'expected_columns': 7}
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

                if table_key == 'security_posts':
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

                    if table_key == 'security_posts':
                        try:
                            total_units += int(data['units'][i]) if data['units'][i] else 0
                            total_persons += int(data['persons'][i]) if data['persons'][i] else 0
                        except ValueError:
                            logger.warning(f"Некорректное значение для units или persons в таблице {table_key}, строка {i}")

                # Для таблицы security_posts добавляем итоговую строку
                if table_key == 'security_posts':
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
