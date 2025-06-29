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
        logger.info("Шаблон успешно загружен")

        # Функция для замены плейсхолдеров в тексте
        def replace_text(text, fields, table_fields):
            if not text:
                return text
            
            # Замена простых полей
            for key, placeholder in text_fields.items():
                if placeholder in text:
                    value = fields.get(key, '').strip()
                    text = text.replace(placeholder, value if value else "")
            
            # Замена плейсхолдеров для таблиц
            for table_key, table_data in table_fields.items():
                row_count = max(len(table_data[key]) for key in table_data if table_data[key]) if any(table_data[key] for key in table_data) else 0
                for i in range(row_count):
                    for sub_key in table_data.keys():
                        # Обрабатываем оба формата плейсхолдеров: {table[i].key} и {table[i].key}
                        placeholder1 = f"{{{table_key}[{i}].{sub_key}}}"
                        placeholder2 = f"{{{table_key}\[{i}\].{sub_key}}}"
                        if placeholder1 in text or placeholder2 in text:
                            value = table_data[sub_key][i] if i < len(table_data[sub_key]) and table_data[sub_key][i] else ""
                            text = text.replace(placeholder1, str(value)).replace(placeholder2, str(value))
            return text

        # Функция для замены плейсхолдеров в таблицах
        def replace_table_placeholders(table, fields, table_fields):
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        original_text = para.text
                        new_text = replace_text(original_text, fields, table_fields)
                        if new_text != original_text:
                            para.text = new_text
                            logger.info(f"Замена в таблице: '{original_text}' -> '{new_text}'")

        # Замена плейсхолдеров в параграфах
        for para in doc.paragraphs:
            original_text = para.text
            new_text = replace_text(original_text, fields, table_fields)
            if new_text != original_text:
                para.text = new_text
                para.alignment = 1  # Выравнивание по центру
                logger.info(f"Замена в параграфе: '{original_text}' -> '{new_text}'")

        # Замена плейсхолдеров в таблицах
        for table in doc.tables:
            replace_table_placeholders(table, fields, table_fields)

        # Обработка специальных случаев для таблицы security_posts (итоговая строка)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if "Всего" in cell.text and "{security_posts[5].units}" in cell.text:
                        total_units = sum(int(x) for x in table_fields['security_posts']['units'] if x.isdigit())
                        total_persons = sum(int(x) for x in table_fields['security_posts']['persons'] if x.isdigit())
                        cell.text = cell.text.replace("{security_posts[5].units}", str(total_units))
                        cell.text = cell.text.replace("{security_posts[5].persons}", str(total_persons))

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
