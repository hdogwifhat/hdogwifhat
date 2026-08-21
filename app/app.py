import os
import subprocess
from flask import Flask, request, send_file, render_template
import docx

app = Flask(__name__)

# Template mapping
TEMPLATES = {
    "102_options_shares": {
        "en": "Holding Approval - 102 trustee - options & shares.docx",
        "he": "אישור החזקה- 102 נאמן - מניות ואופציות.docx"
    },
    "102_shares": {
        "en": "Holding Approval - 102 trustee -Shares.docx",
        "he": "אישור החזקה- 102 נאמן - מניות.docx"
    },
    "103K_shares": {
        "en": "Holding Approval - 103K - Shares.docx",
        "he": "אישור החזקה - 103כ - מניות.docx"
    },
    "104B_shares": {
        "en": "Holding Approval - 104B - Shares.docx",
        "he": "אישור החזקה - 104ב - מניות.docx"
    },
    "104H_shares": {
        "en": "Holding Approval - 104H - Shares.docx",
        "he": "אישור החזקה - 104ח - מניות_1.docx"
    },
    "15C_shares": {
        "en": "Holding Approval - 15C - Shares.docx",
        "he": "אישור החזקה - 15ג - מניות.docx"
    },
    "3i_shares": {
        "en": "Holding Approval - 3i - Shares.docx",
        "he": "Holding Approval - 3i - Shares.docx" # No explicit hebrew file for 3i? Let's assume EN for both or check later
    },
    "shareholder": {
        "en": "Holding Approval - Shareholder.docx",
        "he": "אישור החזקה - בעל מניות- מניות.docx"
    }
}

@app.route("/")
def index():
    return render_template("index.html")

def replace_text_in_paragraphs(paragraphs, replacements):
    for p in paragraphs:
        for run in p.runs:
            for key, val in replacements.items():
                if key in run.text:
                    run.text = run.text.replace(key, str(val))

        for key, val in replacements.items():
            if key in p.text:
                text = p.text
                p.text = text.replace(key, str(val))

@app.route("/generate", methods=["POST"])
def generate():
    template_type = request.form.get("template_type")
    language = request.form.get("language", "he")

    company_name = request.form.get("company_name", "")
    shareholder_name = request.form.get("shareholder_name", "")
    id_number = request.form.get("id_number", "")
    today_date = request.form.get("today_date", "")
    holding_date = request.form.get("holding_date", "")
    number_of_shares = request.form.get("number_of_shares", "")
    total_options = request.form.get("total_options", "")
    vested_options = request.form.get("vested_options", "")
    unvested_options = request.form.get("unvested_options", "")

    if language == "en":
        replacements = {
            "[Name of Company]": company_name,
            "[company’s name]": company_name,
            "[Name of Shareholder]": shareholder_name,
            "[ID number]": id_number,
            "[DD MM, YYYY]": today_date,
            "[Number of Shares]": number_of_shares,
            "[Total Number of options]": total_options,
            "[Vested options]": vested_options,
            "[Non Vested options]": unvested_options,
            "[New Version]": "",
            "[Ordinary]": "Ordinary",
            "[Sir/Madam]": "Sir/Madam",
            "August 21, 2026": today_date,
            "July 29, 2026": today_date,
            "as of the date of this letter": f"as of {holding_date}" if holding_date else "as of the date of this letter"
        }
    else:
        replacements = {
            "[שם החברה]": company_name,
            "[שם הניצע\\ת]": shareholder_name,
            "[שם הניצע\\]": shareholder_name,
            "[שם בעל המניות]": shareholder_name,
            "[תעודת זהות \\ ח\"פ]": id_number,
            "[מס' זהות/ דרכון]": id_number,
            "[מס' ת.ז.]": id_number,
            "[DD בMM, YYYY]": today_date,
            "[מספר המניות]": number_of_shares,
            "[כמות מניות]": number_of_shares,
            "[סך הכול אופציות]": total_options,
            "[אופציות מובשלות]": vested_options,
            "[אופציות לא מובשלות]": unvested_options,
            "[נא לעדכן בהתאם לאדון/גברת]": "",
            "[נוסח חדש]": "",
            "[רגילות]": "רגילות",
            "21 August 2026": today_date,
            "29 יולי 2026": today_date,
            "‏29 יולי 2026": today_date,
            "‏‏29 יולי 2026": today_date,
            "יצחק וייס": shareholder_name,
            "014875538": id_number,
            "Kanabo Group PLC": company_name,
            "נכון למועד מכתב זה": f"נכון למועד {holding_date}" if holding_date else "נכון למועד מכתב זה"
        }

    template_file = TEMPLATES.get(template_type, {}).get(language)
    if not template_file:
        return "Template not found", 404

    # Using os.getcwd() to be sure
    base_dir = os.path.dirname(os.path.abspath(__file__))
    template_path = os.path.join(base_dir, "docx_templates", template_file)

    if not os.path.exists(template_path):
        return f"File {template_path} not found", 404

    doc = docx.Document(template_path)
    replace_text_in_paragraphs(doc.paragraphs, replacements)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_text_in_paragraphs(cell.paragraphs, replacements)

    output_docx = f"/tmp/{template_type}_{language}.docx"
    doc.save(output_docx)

    subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", "/tmp", output_docx], check=True)

    output_pdf = f"/tmp/{template_type}_{language}.pdf"

    return send_file(output_pdf, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True, host='0.0.0.0', port=5000)
