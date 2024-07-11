from flask import Flask, request, redirect, url_for, render_template, send_file
import fitz  # PyMuPDF
import os
import pandas as pd
import sqlite3
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
import platform
import tempfile

app = Flask(__name__)

def get_downloads_folder():
    if platform.system() == "Windows":
        return os.path.join(os.environ['USERPROFILE'], 'Downloads')
    elif platform.system() == "Darwin":
        return os.path.join(os.path.expanduser('~'), 'Downloads')
    else:
        return os.path.join(os.path.expanduser('~'), 'Downloads')

def pdf_to_text_pymupdf(pdf_path):
    document = fitz.open(pdf_path)
    text = ''
    for page_num in range(len(document)):
        page = document.load_page(page_num)
        text += page.get_text()
    document.close()
    return text

def create_tables_from_text(text):
    sections = text.split("Sorties des KITS")
    tables = []
    
    for section in sections[1:]:
        words = section.split()
        if len(words) < 4:
            continue
        
        Kit = words[0]
        data = []
        s = 0
        
        for i in range(len(words) - 5):
            if words[i] == Kit:
                s += 1
                if s != 1 and i + 5 < len(words):
                    Constituant = words[i+1]
                    Emplacement = words[i+2]
                    if Emplacement == "Y" or Emplacement == "X":
                        Emplacement = Emplacement + " " + words[i+3]
                        Quantité = words[i+4]
                        m = i + 4
                    else:
                        Quantité = words[i+3]
                        m = i + 3
                    Nom = words[m+1]
                    k = m + 2
                    while k < len(words) and words[k] not in ["CSFAME", "BSFGK"]: 
                        Nom = Nom + " " + words[k]
                        k += 1
                    data.append([Kit, Constituant, Emplacement, Quantité, Nom])
        
        df = pd.DataFrame(data, columns=["Kit", "Constituant", "Emplacement", "Quantité", "Nom"])
        tables.append(df)
    
    return tables

def process_emplacement(df):
    def process_value(value):
        words = value.split()
        if words[0][0].isdigit():
            value = 'S-' + value
        if len(words) == 2:
            value = ''.join(words)
        return value
    
    df['Emplacement'] = df['Emplacement'].apply(process_value)
    return df

def generate_pdf(dataframe, filename, colors_list, kit_color_map):
    pdf = SimpleDocTemplate(filename, pagesize=letter)
    elements = []
    data = [list(dataframe.columns)] + dataframe.values.tolist()
    table = Table(data)
    style = TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.grey), ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke)])
    
    for i, row in enumerate(data[1:], start=1):
        bg_color = kit_color_map.get(row[0], colors.white)
        style.add('BACKGROUND', (0, i), (-1, i), bg_color)
    
    table.setStyle(style)
    elements.append(table)
    pdf.build(elements)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return redirect(request.url)
    file = request.files['file']
    if file.filename == '':
        return redirect(request.url)
    if file:
        with tempfile.TemporaryDirectory() as tempdir:
            file_path = os.path.join(tempdir, file.filename)
            file.save(file_path)
            text = pdf_to_text_pymupdf(file_path)
            
            tables = create_tables_from_text(text)
            tables = [process_emplacement(table) for table in tables]
            
            conn = sqlite3.connect(':memory:')  # Utilisation de la base de données en mémoire
            for i, table in enumerate(tables):
                table_name = f'table_{i+1}'
                table.to_sql(table_name, conn, index=False, if_exists='replace')
            
            combined_df = pd.DataFrame(columns=["Kit", "Constituant", "Emplacement", "Quantité", "Nom"])
            cursor = conn.cursor()
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
            table_names = [table[0] for table in cursor.fetchall()]
            
            for table_name in table_names:
                df = pd.read_sql_query(f"SELECT * FROM {table_name}", conn)
                combined_df = pd.concat([combined_df, df])
            
            combined_df.drop_duplicates(inplace=True)
            
            def sort_key(value):
                import re
                match = re.match(r"([a-zA-Z]+)([0-9]+)", value)
                if match:
                    alpha, num = match.groups()
                    return (alpha, int(num))
                else:
                    return (value, 0)
            
            combined_df = combined_df.sort_values(by='Emplacement', key=lambda col: col.map(sort_key))
            
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS merged_table (
                    Kit TEXT,
                    Constituant TEXT,
                    Emplacement TEXT,
                    Quantité TEXT,
                    Nom TEXT
                )
            ''')
            conn.commit()
            
            combined_df.to_sql('merged_table', conn, index=False, if_exists='replace')
            
            merged_df = pd.read_sql_query("SELECT * FROM merged_table", conn)
            
            kit_color_map = {}
            for i in range(1, (len(request.form) - 1) // 2 + 1):  # Ajustez le calcul pour correspondre au nombre réel d'entrées
                kit_name = request.form.get(f'kit_name_{i}')
                color_hex = request.form.get(f'color_{i}')
                if kit_name and color_hex:
                    kit_color_map[kit_name] = colors.HexColor(color_hex)
            
            downloads_folder = get_downloads_folder()
            output_pdf_path = os.path.join(downloads_folder, 'merged_table_colored.pdf')
            generate_pdf(merged_df, output_pdf_path, list(kit_color_map.values()), kit_color_map)
            
            conn.close()
            
            return send_file(output_pdf_path, as_attachment=True, download_name='merged_table_colored.pdf')

if __name__ == '__main__':
    app.run(debug=True)
