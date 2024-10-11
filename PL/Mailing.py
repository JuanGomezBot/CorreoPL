import base64
import csv
import email.header
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from collections import Counter
import email
from email import encoders
import os
import uuid
import time
import os
import platform
import time
import tempfile
import threading
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from concurrent.futures import ThreadPoolExecutor
from tkinter import messagebox
import mysql.connector

import re


from email.utils import parseaddr

from reportlab import *
# Configuración de la base de datos
db_config = {
    'host': 'NOTTELLINGYOU',
    'user': 'No',
    'password': 'Just No',
    'database': 'NoDatabaseHere'
}

user = 'RandomUserName'
password = 'password123L'
server_address = "21 Baker Street"
port = 2525

def extract_email_from_dict(email_dict):
    """Extract email from dictionary """
    if isinstance(email_dict, dict):
        for key in email_dict:
            if 'Correo' in key or 'Email' in key:
                return email_dict[key]
    return None

def fetch_data_from_database(email_id=None, subject=None, date_value=None, recipient_email=None):
    """Grab all data rows from SQL database"""
    query = "SELECT * FROM email_tracking WHERE 1=1"
    parameters = []

    if email_id:
        query += " AND email_id = %s"
        parameters.append(email_id)

    if subject:
        query += " AND subject = %s"
        parameters.append(subject)
    
    if date_value:
        query += " AND DATE(timestamp) = %s"
        parameters.append(date_value)
    
    if recipient_email:
        query += " AND recipient_email = %s"
        parameters.append(recipient_email)

    conn = mysql.connector.connect(**db_config)
    cursor = conn.cursor(dictionary=True)
    
    try:
        cursor.execute(query, parameters)
        results = cursor.fetchall()
    finally:
        cursor.close()
        conn.close()
    
    return results
#Verifies that email has a valid form
def is_valid_email(email):
    email_regex = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    
    if isinstance(email, dict):
        
        for key in email:
            if 'Correo' in key or 'email' in key:
                email = email[key]
                break
    
    if not isinstance(email, str):
        print("The variable obtained is not a string")
        return False
    
    return re.match(email_regex, email) is not None

def verify_email_address(email):
    """Verify Valid information"""
    if isinstance(email, dict):
        email = extract_email_from_dict(email)
    
    if not isinstance(email, str):
        print("The email is not a string")
        return False

    if '@' not in email:
        print("An email needs to have an @")
        return False

    domain = email.split('@')[-1]

    if not re.match(r'^[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$', domain):
        print("Invalid domain")
        return False

    return True

def check_email_and_send(email, subject, body):
    """Comprobar la validez del correo electrónico y enviar si es válido."""
    if not is_valid_email(email):
        print(f"Formato de dirección de correo electrónico no válido: {email}")
        return False

    if verify_email_address(email):
        print(f"La dirección de correo electrónico es válida: {email}")
        return True
    else:
        print(f"La dirección de correo electrónico no es válida o ha sido rebotada: {email}")
        return False

def create_table_if_not_exists(cursor):
    """Crea la tabla email_tracking si no existe."""
    create_table_query = """
    CREATE TABLE IF NOT EXISTS email_tracking (
        id INT AUTO_INCREMENT PRIMARY KEY,
        email_id VARCHAR(255) NOT NULL,
        recipient_email VARCHAR(255) NOT NULL,
        status ENUM('sent', 'opened', 'failed') NOT NULL,
        timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
    """
    cursor.execute(create_table_query)

def fetch_all_emails():
    """Obtener todos los registros de correo electrónico de la base de datos."""
    query = "SELECT * FROM email_tracking"
    conn = None
    cursor = None
    results = []

    try:
        conn = mysql.connector.connect(**db_config)
        cursor = conn.cursor(dictionary=True)
        cursor.execute(query)
        results = cursor.fetchall()
    
    except mysql.connector.Error as e:
        print(f"Error al obtener datos de la base de datos: {e}")
    
    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()
    
    return results

def generate_pdf(email_id, date_value, recipient_email, pdf_path):
    """Genera un informe en PDF según los parámetros dados."""
    data = fetch_data_from_database(email_id, date_value, recipient_email)
    
    c = canvas.Canvas(pdf_path, pagesize=letter)
    width, height = letter

    c.drawString(1 * inch, height - 1 * inch, f"Informe de correo electrónico para el ID de correo: {email_id}")
    
    y_position = height - 2 * inch
    for record in data:
        c.drawString(1 * inch, y_position, f"ID de correo: {record['email_id']}")
        y_position -= 0.5 * inch
        c.drawString(1 * inch, y_position, f"Correo electrónico del destinatario: {record['recipient_email']}")
        y_position -= 0.5 * inch
        c.drawString(1 * inch, y_position, f"Estado: {record['status']}")
        y_position -= 0.5 * inch
        c.drawString(1 * inch, y_position, f"Marca de tiempo: {record['timestamp']}")
        y_position -= 1 * inch
    
    c.save()
    print(f"Informe en PDF creado en {pdf_path}")

def open_pdf(pdf_path):
    """Abre el archivo PDF usando el visor predeterminado."""
    if platform.system() == 'Darwin':  
        os.system(f'open "{pdf_path}"')
    elif platform.system() == 'Windows':  
        os.system(f'start "" "{pdf_path}"')
    else:  # Linux
        os.system(f'xdg-open "{pdf_path}"')

def is_file_open(pdf_path):
    """Verificar si el archivo está aún abierto."""
    try:
        with open(pdf_path, 'r+b') as file:
            return False
    except IOError:
        return True

def generate_open_and_delete_pdf(email_id, date_value, recipient_email):
    """Genera un informe en PDF según el email_id, fecha y destinatario, lo abre y luego lo elimina una vez cerrado."""
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp_file:
        pdf_path = temp_file.name

    try:
        generate_pdf(email_id, date_value, recipient_email, pdf_path)
        open_pdf(pdf_path)
        

    except Exception as e:
        print(f"Se produjo un error: {str(e)}")
        if os.path.exists(pdf_path):
            os.remove(pdf_path)

def log_email_status(email_id, recipient_email, subject, body=None, attachments=None):
    """Registra el estado del correo electrónico en la base de datos con parámetros opcionales."""
    conn = None
    cursor = None
    try:
        conn = mysql.connector.connect(**db_config)
        cursor = conn.cursor()
        attachs = ""
        if attachments:
            attachs = ",".join(attachments)
        
        email_id_str = str(email_id) if email_id else None

        query = """
        INSERT INTO email_tracking (email_id, subject,recipient_email, status, body, attachments)
        VALUES (%s, %s, %s, %s, %s, %s)
        """
        parameters = (email_id_str, recipient_email, subject, "Enviado", body, attachs)
        
        print(f"Ejecutando consulta: {query} con parámetros: {parameters}")
        cursor.execute(query, parameters)
        conn.commit()
        print("Estado del correo electrónico registrado con éxito.")

    except mysql.connector.Error as err:
        print(f"Error: {err}")
    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()

placeholders = []
cached_data = []

def add_tracking_image(email_body_template, email_id):
    """Agregar una imagen de seguimiento al cuerpo del correo electrónico."""
    tracking_url = f"https://www.proyeccionlaboral.com/track.php?email_id={email_id}&cb={int(time.time())}"
    
    tracking_image = f'<img src={tracking_url} width="100" height="100" style="display:block;" alt="logo"/>'
    
    return email_body_template + tracking_image
     

def detect_delimiter(sample_line):
    """Detect the delimiter used in a CSV file based on the most common delimiter."""
    common_delimiters = [',', ';', '\t', '|', ':']
    delimiter_counts = Counter({delimiter: sample_line.count(delimiter) for delimiter in common_delimiters})
    return delimiter_counts.most_common(1)[0][0]

def read_csv_data(csv_file):
    """Read CSV Data and will be stored on a dictionary"""
    global placeholders, cached_data

    
    cached_data = None

    data = []
    placeholders = []  
    try:
        with open(csv_file, mode='r', encoding='latin1') as file:
            
            sample_line = file.readline()
            delimiter = detect_delimiter(sample_line)
            file.seek(0)  

            reader = csv.DictReader(file, delimiter=delimiter)
            placeholders = reader.fieldnames  
            data = list(reader)
    
       
        cached_data = data
    except FileNotFoundError:
        print(f"File '{csv_file}' not found.")
    except Exception as e:
        
        print(f"Error reading file: {str(e)}")
    return data

def replace_placeholders(email_body_template, row):
    """Will replace placeholder with the CSV Data"""
    for key, value in row.items():
        placeholder = f"${key}"
        if placeholder in email_body_template:
            email_body_template = email_body_template.replace(placeholder, str(value))
        
    return email_body_template

def attach_file_to_msg(msg, file_path):
    """Attaches email file"""
    try:
        with open(file_path, "rb") as file:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(file.read())
        encoders.encode_base64(part)
        filename = email.header.Header(os.path.basename(file_path), "utf-8").encode()
        part.add_header("Content-Disposition", f"attachment; filename={filename}")
        msg.attach(part)
    except Exception as e:
        print(f"Error adjuntando archivo {file_path}. Error: {str(e)}")

def send_email_with_attachments(to_email, subject, body, attachment_files=None, email_id=None):
    """Attaches file with a optional file."""
    msg = MIMEMultipart()
    msg['From'] = user
    msg['To'] = to_email
    msg['Subject'] = subject
    body = body.replace("\n", "<br>")

    msg.attach(MIMEText(body, 'html'))  


    if attachment_files:
        with ThreadPoolExecutor() as executor:
            futures = [executor.submit(attach_file_to_msg, msg, file_path) for file_path in attachment_files if file_path]
            for future in futures:
                future.result()  

    try:
        with smtplib.SMTP(server_address, port) as smtp_server:
            smtp_server.starttls()
            smtp_server.login(user, password)
            if to_email:
                smtp_server.sendmail(user, to_email, msg.as_string())
                print(f"Email successfully sent to {to_email}")
                smtp_server.close()
                return True
    except Exception as e:
        print(f"Error sending email to {to_email}. Error: {str(e)}")
        
        return False

def send_emails(recipients, subject, body_template, global_attachments=None, flag=0, email_id=None):
    """Threading in order to avoid memory consuming"""
    def send_email(recipient):
        to_email = None
        for key in recipient:
            if 'Correo' in key:
                to_email = recipient[key]
                break
        email_id = uuid.uuid4()
        if to_email is None:
            print("Can't be found email")
            return
        
        attachment_files = []
        for key in recipient:
            if key.lower() in ['file', 'Files']:
                archivo_value = recipient.get(key, '')
                if isinstance(archivo_value, str):
                    attachment_files = [file.strip() for file in archivo_value.split(',') if file.strip()]
                break
        
        if flag == 0:
            attachment_files = []
        
        if global_attachments:
            attachment_files.extend(global_attachments)
            print(attachment_files)
        
        attachment_files = [file[0] if isinstance(file, tuple) else file for file in attachment_files]

        replaced_body = replace_placeholders(body_template, recipient)
        cuerpocorreo = add_tracking_image(replaced_body, email_id)
        if check_email_and_send(recipient, subject, cuerpocorreo):
            send_email_with_attachments(to_email, subject, cuerpocorreo, attachment_files)
            log_email_status(email_id, to_email, subject, cuerpocorreo, attachment_files)
    
    with ThreadPoolExecutor() as executor:
        futures = [executor.submit(send_email, recipient) for recipient in recipients]
        for future in futures:
            future.result()  # Esperar a que todos los correos sean enviados

def show_variables():
    """Shows variables that can be used on email"""
    global placeholders
    if placeholders:
        messagebox.showinfo("Available variables on CSV", f"Placeholders found:\n\n{', '.join(placeholders)}")
    else:
        messagebox.showwarning("No Variables", "Not found variables on CSV.")

def send_mass_email(csv_file, subject, body_template, global_attachment=None):
    """Mass email using emails found in csv."""
    recipients = read_csv_data(csv_file)
    if recipients:
        send_emails(recipients, subject, body_template, global_attachment)
    else:
        print("No se encontraron destinatarios válidos.")

def send_personalized_email(csv_file, subject, body_template):
    """Mass personalized emails using a CSV"""
    recipients = read_csv_data(csv_file)
    if recipients:
        send_emails(recipients, subject, body_template, flag=1)
    else:
        print("No se encontraron destinatarios válidos.")

