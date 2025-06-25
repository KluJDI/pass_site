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
    # Collect all form data
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

    # Load the Word document template
    doc = Document("pasport_mesta_massovogo_prebuvaniya_lyudei.docx")

    # Replace placeholders in paragraphs
    def replace_placeholders(paragraphs):
        for para in paragraphs:
            for key, val in fields.items():
                para.text = re.sub(r'_+', val, para.text, count=1)

    # Replace placeholders in paragraphs
    replace_placeholders(doc.paragraphs)

    # Replace placeholders in tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_placeholders(cell.paragraphs)

    # Update tables with dynamic data
    def update_table(table_index, field_data, column_count):
        table = doc.tables[table_index]
        # Keep header row, clear other rows
        if len(table.rows) > 1:
            for _ in range(len(table.rows) - 1):
                table._element.remove(table.rows[-1]._element)
        # Add new rows based on form data
        for i in range(len(field_data[list(field_data.keys())[0]])):
            row = table.add_row()
            for j, key in enumerate(field_data.keys()):
                if i < len(field_data[key]):
                    row.cells[j + 1].text = field_data[key][i]  # +1 to skip N п/п column

    # Update each table with corresponding form data
    update_table(0, table_fields['objects_on_territory'], 4)  # Table 2
    update_table(1, table_fields['objects_nearby'], 4)  # Table 3
    update_table(2, table_fields['transport'], 3)  # Table 4
    update_table(3, table_fields['service_orgs'], 3)  # Table 5
    update_table(4, table_fields['dangerous_sections'], 3)  # Table 7
    update_table(5, table_fields['consequences'], 3)  # Table 9
    update_table(6, table_fields['patrol_composition'], 3)  # Table 10г
    update_table(7, table_fields['critical_elements'], 6)  # Table 12

    # Save the document to a BytesIO object
    output = io.BytesIO()
    doc.save(output)
    output.seek(0)

    return send_file(output, download_name="Паспорт_безопасности.docx", as_attachment=True)

import os

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)

