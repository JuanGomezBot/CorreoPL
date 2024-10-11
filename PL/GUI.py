from collections import Counter
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
import Mailing  # Assuming Mailing.py contains the email sending functions
from PIL import Image, ImageTk  # Ensure Pillow is installed for image handling
import csv
import os
import chardet
import mysql.connector
from tkinter import ttk

# Inicialización global de variables
variable_frame = None
columns_listbox = None
flag = None
huge_frame = None
canvas = None   
file_paths = []
csv_file_path = None
variables_list = []

class HyperlinkButton(tk.Button):
    """Personalized Hyperlink Button"""
    def __init__(self, master=None, **kwargs):
        super().__init__(master=master, **kwargs)
        self.defaultBackground = self["background"]
        self.bind("<Enter>", self.on_enter)
        self.bind("<Leave>", self.on_leave)
        self.config(relief="flat", borderwidth=0, fg="#0066cc", font=("Helvetica", 12, "underline"), bg="#f7f7f7")

    def on_enter(self, e):
        self.config(fg="#003366")

    def on_leave(self, e):
        self.config(fg="#0066cc")
def detect_encoding(file_path):
    """Detects encoding obn email."""
    with open(file_path, 'rb') as f:
        result = chardet.detect(f.read())
    return result['encoding']

def correct_csv_data(file_path):
    """In case of impure or bad data it will be cleaned up"""
    encoding = detect_encoding(file_path)
    temp_file_path = file_path + '_temp'

    try:
        if needs_correction(file_path):
            print("File needs to be fixed Processing...")

            with open(file_path, 'r', encoding=encoding) as infile, \
                 open(temp_file_path, 'w', newline='', encoding=encoding) as outfile:
                reader = csv.reader(infile)
                writer = csv.writer(outfile)

                header = next(reader, None)
                if header:
                    corrected_header = header[0].split(';')
                    writer.writerow(corrected_header)
                else:
                    print("Headers not found")

                for row in reader:
                    if len(row) == 1:
                        row_data = row[0].split(';')
                        if len(row_data) > 1:
                            writer.writerow(row_data)
                        else:
                            writer.writerow(row)
                    else:
                        writer.writerow(row)

            os.replace(temp_file_path, file_path)
            print("File Overwritten")
        else:
            print("File is good")
        
    except PermissionError as e:
        print(f"Seems you are not supposed to do this: {e}")
        print("Make sure this is the only file open")
    except Exception as e:
        print(f"Error Ocurred: {e}")

def needs_correction(file_path):
    """Checks if the file needs to be fixed"""
    encoding = detect_encoding(file_path)
    
    with open(file_path, 'r', encoding=encoding) as file:
        reader = csv.reader(file)
        first_row = next(reader, None)
        file.seek(0)
        first_column = [row[0] for row in reader if row]

    if first_row:
        delimiter_count = len(first_row[0].split(';')) - 1
        if delimiter_count > 5:
            return True
    return False

def select_csv_file(show_variables_flag=True):
    """Allow the selection of CSV file."""
    global csv_file_path
    csv_file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if csv_file_path and needs_correction(csv_file_path):
        correct_csv_data(csv_file_path)
    process_csv_file()
    print(f"CHosen file: {csv_file_path}")
    if show_variables_flag and root.current_page == "personalizado":
        show_variables(csv_file_path)
    if root.current_page == "personalizado":
        show_personalizado_page()

def attach_files():
    """Allows multiple files to be attached."""
    global file_paths
    file_paths = filedialog.askopenfilenames(filetypes=[("All Files", "*.*")])
    if file_paths:
        print(f"Archivos seleccionados: {', '.join(file_paths)}")


def use_csv_file():
    """Returns CSV file Path."""
    global csv_file_path
    return csv_file_path

def send_personalized_email(widget_refs):
    """Sends personalized mass email from CSV"""

    def send_emails():
        csv_file = use_csv_file()
        print(file_paths)
        if csv_file:
            recipients = Mailing.read_csv_data(csv_file)
            print(recipients)
            if recipients:
                subject = widget_refs['subject_entry'].get("1.0", "end-1c").strip()
                body = widget_refs['body_text'].get("1.0", "end-1c").strip()
                print("Filepaths: " + str(file_paths))
                Mailing.send_emails(recipients, subject, body, global_attachments=file_paths, flag=1)
                messagebox.showinfo("Personalized email", "Emails have been sent sucessfully.")
            else:
                messagebox.showwarning("Error", "No valid destinated emails found")
        else:
            messagebox.showwarning("Error", "Select CSV file.")

        # Hide loading animation after emails are sent
        root.after(0, hide_loading)

    # Show loading animation
    show_loading()
    
    # Schedule the email sending function to run after the current event loop
    root.after(100, send_emails)

def show_personalizado_page():
    """Shows personalized email"""
    global widget_refs
    global current_page, csv_file_path, file_paths, widget_refs, flag

    hide_current_widgets()
    show_copyright_image()
    
    root.current_page = "personalizado"
    root.title("JuanRep\CorreoPL\PL\EvaMailer.png")
    root.geometry("800x600")
    widget_refs = {}
    try:
        logo_image = Image.open("JuanRep\CorreoPL\PL\EvaMailer.png")
        logo_image = logo_image.resize((80, 80), Image.Resampling.LANCZOS)
        logo = ImageTk.PhotoImage(logo_image)
        logo_label = tk.Label(root, image=logo, bg="#f7f7f7")
        logo_label.image = logo
        logo_label.grid(row=0, column=1, pady=10, columnspan=2, sticky='n')
    except Exception as e:
        print(f"Error loading image: {e}")
    regresar_button = tk.Button(root, text="return", command=reset_app, bg="#f7f7f7", font=("Helvetica", 10))
    regresar_button.grid(row=0, column=0, padx=10, pady=10, sticky='w')
    global columns_listbox
    columns_listbox = tk.Listbox(root, bg="#ffffff", font=("Helvetica", 10))
    columns_listbox.grid(row=2, column=0, sticky="ns", rowspan=2, padx=10, pady=10)
    widget_refs['columns_listbox'] = columns_listbox
    adjuntar = tk.Button(root, text="Attach", command=lambda: attach_files(), bg="#0066cc", fg="#ffffff", font=("Helvetica", 13), width=8, height=1)
    
    adjuntar.grid(row=4, column=1, columnspan=3, pady=(10, 10))
    process_csv_button = tk.Button(root, text="Import data file", command=lambda: select_csv_file(), bg="#0066cc", fg="#ffffff", font=("Helvetica", 13),  width=20, height=1)
    process_csv_button.grid(row=4, column=1, padx=10, pady=(10, 10), sticky="w")

    subject_label = tk.Label(root, text="Subject", font=('Helvetica', 12), bg="#f7f7f7")
    subject_label.grid(row=1, column=1, padx=10, pady=(10, 2), sticky='w', columnspan=2)
    subject_entry = tk.Text(root, width=40, height=2, font=("Helvetica", 10), bg="#ffffff")
    subject_entry.grid(row=2, column=1, padx=10, pady=(2, 10), sticky='enw', columnspan=2)
    widget_refs['subject_entry'] = subject_entry
    body_label = tk.Label(root, text="Email Body", font=('Helvetica', 12), bg="#f7f7f7")
    body_label.grid(row=2, column=1, padx=10, pady=(15, 2), sticky='w', rowspan=1)
    body_text = tk.Text(root, width=60, height=10, font=("Helvetica", 10), bg="#ffffff")
    body_text.grid(row=3, column=1, padx=10, pady=(15, 2), sticky='nsew', rowspan=1)
    widget_refs['body_text'] = body_text
    sendP_button = tk.Button(root, text="Send", command=lambda: send_personalized_email(widget_refs), bg="#0066cc", fg="#ffffff", font=("Helvetica", 13), width=8, height=1)
    sendP_button.grid(row=4, column=2, padx=10, pady=(10, 10), sticky="e")
    try:
        help_image = tk.PhotoImage(file="JuanRep\CorreoPL\PL\HELP.png")
        help_image = help_image.subsample(10)
    except Exception as e:
        print(f"Error loading image: {e}")
        help_image = None
    if help_image:
        help_button = tk.Button(root, image=help_image, command=ayudaPersonalizado, relief="flat", bg="#f7f7f7")
        help_button.grid(row=0, column=10, padx=10, pady=10, sticky='ne')
        help_button.image = help_image
    
    display_csv_variables(variables_list, body_text)
    root.grid_columnconfigure(0, weight=1)
    root.grid_columnconfigure(1, weight=3)
    root.grid_columnconfigure(2, weight=3)
    root.grid_columnconfigure(10, weight=0)
    root.grid_rowconfigure(0, weight=0)
    root.grid_rowconfigure(1, weight=0)
    root.grid_rowconfigure(2, weight=1)
    root.grid_rowconfigure(3, weight=2)
    root.grid_rowconfigure(4, weight=0)
    root.bind("<Configure>", on_resize)
def show_loading():
    """Displays a loading overlay with an image and text on the root window."""
    global overlay 
    overlay = tk.Toplevel(root)
    overlay.geometry(f"{root.winfo_width()}x{root.winfo_height()}+{root.winfo_x()}+{root.winfo_y()}")
    overlay.overrideredirect(True)  # Removes window decorations
    overlay.configure(bg='gray')

    # Load and place the image
    try:
        loading_image = Image.open("JuanRep\CorreoPL\PL\loading.png")
        loading_image = loading_image.resize((300, 300), Image.Resampling.LANCZOS)
        loading_photo = ImageTk.PhotoImage(loading_image)
        image_label = tk.Label(overlay, image=loading_photo, bg='gray')
        image_label.image = loading_photo  
        image_label.place(relx=0.5, rely=0.4, anchor='center')
    except Exception as e:
        print(f"Error loading image: {e}")

    

    overlay.update_idletasks()

def hide_loading():
    """Hide the loading screen."""
    global overlay  
    if overlay:  
        overlay.destroy()  

def process_csv_file():
    """Process CSV File."""
    global csv_file_path
    if not csv_file_path:
        messagebox.showwarning("Error", "No selected CSV file.")
        return

def display_csv_variables(variables_list, text_field):
    """Shows variables found in CSV file"""
    for widget in root.grid_slaves(row=2, column=0):
        widget.destroy()

    variables_frame = tk.Frame(root, bg="#f7f7f7")
    variables_frame.grid(row=2, column=0, sticky="ns")

    scrollbar = tk.Scrollbar(variables_frame)
    scrollbar.pack(side="right", fill="y")

    canvas = tk.Canvas(variables_frame, yscrollcommand=scrollbar.set, width=200, background="#f7f7f7", bd=0, highlightthickness=0)
    canvas.pack(side="left", fill="both", expand=True)
    
    scrollbar.config(command=canvas.yview)

    button_frame = tk.Frame(canvas, bg="#f7f7f7")
    canvas.create_window((0, 0), window=button_frame, anchor='nw')

    for i, variable in enumerate(variables_list):
        if variable == "ī»¿Correo":
            variable = "Correo"
       
        button = tk.Button(button_frame, text=variable, command=lambda var=variable: insert_variable(text_field, var),
                           bg="#f7f7f7", fg="#000000", relief="flat", font=("Helvetica", 12))
        button.grid(row=i, column=0, padx=5, pady=5, sticky="w")

    button_frame.update_idletasks()
    canvas.config(scrollregion=canvas.bbox("all"))

def insert_variable(text_field, variable_name):
    """Inserts placeholder in text"""
    current_text = text_field.get("1.0", "end-1c")
    text_field.delete("1.0", tk.END)
    text_field.insert(tk.END, current_text + f"${variable_name}")

def hide_current_widgets():
    """Hide current widgets"""
    for widget in root.grid_slaves():
        widget.grid_forget()

def resize_fonts(event=None):
    """Adjusts font size depending of the window state."""
    width = root.winfo_width()
    height = root.winfo_height()
    
    # Determine sizes based on window dimensions
    logo_size = min(width, height) // 12  # Adjust logo size for better fit
    large_font_size = int(width // 50)  # Adjusted to fit better
    medium_font_size = int(width // 60)
    
    # Update label and button fonts
    if 'label' in globals() and label.winfo_exists():
        label.config(font=('Helvetica', large_font_size))
    if 'personalizado_title' in globals() and personalizado_title.winfo_exists():
        personalizado_title.config(font=('Helvetica', large_font_size))
    if 'personalizado_button' in globals() and personalizado_button.winfo_exists():
        personalizado_button.config(font=("Helvetica", large_font_size, 'underline'))  # Adjust font size for button
    if 'pdf_title' in globals() and pdf_title.winfo_exists():
        pdf_title.config(font=('Helvetica', large_font_size))
    if 'pdf_button' in globals() and pdf_button.winfo_exists():
        pdf_button.config(font=("Helvetica", large_font_size, 'underline'))  # Adjust font size for button
    
    # Resize and update the logo image
    try:
        logo = Image.open("JuanRep\CorreoPL\PL\EvaMailer.png")
        logo = logo.resize((logo_size, logo_size))
        logo_img = ImageTk.PhotoImage(logo)
        if 'logo_label' in globals() and logo_label.winfo_exists():
            logo_label.config(image=logo_img)
            logo_label.image = logo_img
    except FileNotFoundError:
        print("Error: Logo not found.")

    # Adjust button width and height based on new font size
    if 'personalizado_button' in globals() and personalizado_button.winfo_exists():
        personalizado_button.config(width=max(15, large_font_size // 2), height=max(2, large_font_size // 20))

def show_pdf_page():
    """Shows PDF Generator."""
    global entry_date, entry_recipient, entry_subject
    root.current_page = "pdf"
    root.title("JuanRep\CorreoPL\PL\EvaMailer.png")
    root.geometry("1500x900")  # Tamaño más grande de la ventana
    root.config(bg="#f7f7f7")


def show_email_data():
    """Shows all data from emails"""
    global canvas, huge_frame, v_scrollbar, h_scrollbar, search_entry, search_button

    root.state("zoomed")
    hide_current_widgets()
    
    # Clear any existing widgets
    for widget in root.winfo_children():
        widget.destroy()
    
    huge_frame = tk.Frame(root, bg="#0066cc")
    huge_frame.grid(sticky="nsew")

    root.grid_rowconfigure(0, weight=1)
    root.grid_columnconfigure(0, weight=1)
    
    canvas = tk.Canvas(huge_frame, bg="white")
    canvas.grid(row=1, column=0, sticky="nsew")

    huge_frame.grid_rowconfigure(1, weight=1)
    huge_frame.grid_columnconfigure(0, weight=1)
    
    # Vertical scrollbar
    v_scrollbar = tk.Scrollbar(huge_frame, orient="vertical", command=canvas.yview)
    v_scrollbar.grid(row=1, column=1, sticky="ns")
    canvas.config(yscrollcommand=v_scrollbar.set)
    
    # Horizontal scrollbar
    h_scrollbar = tk.Scrollbar(huge_frame, orient="horizontal", command=canvas.xview)
    h_scrollbar.grid(row=2, column=0, sticky="ew")
    canvas.config(xscrollcommand=h_scrollbar.set)
    
    inner_frame = tk.Frame(canvas, bg="white")
    inner_frame_id = canvas.create_window((0, 0), window=inner_frame, anchor="nw")

    def resize_canvas(event):
        canvas.config(width=event.width, height=event.height)
        canvas.itemconfig(inner_frame_id, width=event.width)
    
    huge_frame.bind("<Configure>", resize_canvas)

    def update_displayed_emails():
        search_term = search_entry.get().strip().lower()
        filtered_emails = fetch_filtered_emails_from_db(search_term)
        display_emails(filtered_emails)
    
    def fetch_filtered_emails_from_db(search_term):
        conn = mysql.connector.connect(
            host='gator3405.hostgator.com',
            user='plaboral',
            password='PLap2020*!',
            database='plaboral_correos_recibidos'
        )
        cursor = conn.cursor(dictionary=True)
        
        query = """
        SELECT * FROM email_tracking
        WHERE LOWER(email_id) LIKE %s OR 
              LOWER(subject) LIKE %s OR
              LOWER(recipient_email) LIKE %s OR
              LOWER(status) LIKE %s OR
              LOWER(timestamp) LIKE %s OR
              LOWER(body) LIKE %s OR
              LOWER(attachments) LIKE %s;
        """
        search_pattern = f"%{search_term}%"
        cursor.execute(query, (search_pattern, search_pattern, search_pattern, search_pattern, search_pattern, search_pattern, search_pattern))
        rows = cursor.fetchall()
        
        conn.close()
        
        return rows
    
    def display_emails(emails):
        canvas.delete("all")  # Clear previous content

        if not emails:
            canvas.create_text(10, 10, anchor='nw', text="No se encontraron resultados.", font=('Helvetica', 12))
        else:
            y_position = 10
            for email in emails:
                # Draw separator line
                canvas.create_line(10, y_position, canvas.winfo_width() - 10, y_position, fill="black")
                y_position += 20
                
                # Email details
                canvas.create_text(10, y_position, anchor='nw', text=f"ID: {email['email_id']}", font=('Helvetica', 12))
                y_position += 20
                canvas.create_text(10, y_position, anchor='nw', text=f"Correo del destinatario: {email.get('subject', '')}", font=('Helvetica', 12))
                y_position += 20
                canvas.create_text(10, y_position, anchor='nw', text=f"Asunto: {email.get('recipient_email', '')}", font=('Helvetica', 12))
                y_position += 20
                canvas.create_text(10, y_position, anchor='nw', text=f"Estado: {email.get('status', '')}", font=('Helvetica', 12))
                y_position += 20
                canvas.create_text(10, y_position, anchor='nw', text=f"Fecha de envío: {email.get('timestamp', '')}", font=('Helvetica', 12))
                y_position += 20
                canvas.create_text(10, y_position, anchor='nw', text=f"Contenido: {email.get('body', '')}", font=('Helvetica', 12))
                y_position += 20
                canvas.create_text(10, y_position, anchor='nw', text=f"Archivos adjuntos: {email.get('attachments', '')}", font=('Helvetica', 12))
                y_position += 20
                canvas.create_line(10, y_position, canvas.winfo_width() - 10, y_position, fill="white")
                # Add padding before the button
                y_position += 10  # Adjust this value for the desired padding
                
                # Create the action button
                action_button = tk.Button(canvas, text="Generar Certificado", command=lambda email_id=email['email_id']: Mailing.generate_open_and_delete_pdf(email_id, None, None))
                canvas.create_window(10, y_position, anchor='nw', window=action_button)
                y_position += 30  # Space for the button


                # Draw bottom separator line
                canvas.create_line(10, y_position, canvas.winfo_width() - 10, y_position, fill="black")
                y_position += 20

        # Update scroll region after all content is added
        canvas.config(scrollregion=canvas.bbox("all"))

    # Search frame and button setup
    search_frame = tk.Frame(huge_frame, bg="green")
    search_frame.grid(row=0, column=0, sticky='ne', padx=10, pady=10)

    search_entry = tk.Entry(search_frame)
    search_entry.pack(side='left')

    search_button = tk.Button(search_frame, text="Buscar", command=update_displayed_emails)
    search_button.pack(side='left', padx=5)

    # Display emails initially
    update_displayed_emails()

    # Return button setup
    regresar_button = tk.Button(huge_frame, text="Regresar", command=show_main_page)
    regresar_button.grid(row=0, column=0, sticky='nw', padx=10, pady=10)

    # Resize handling
    def on_resize(event):
        canvas.config(width=event.width, height=event.height)
        canvas.itemconfig(inner_frame_id, width=event.width)

    root.bind('<Configure>', on_resize)


def reset_app():
    """Resets app as if it was on menu"""
    global canvas, csv_file_path, file_paths, variables_list, variable_frame, huge_frame

    # Clean variables
    csv_file_path = None
    file_paths = []
    variables_list = []

    # Makes sure no floating variables are here
    if variable_frame:
        variable_frame.destroy()

    # Destroys frame
    if huge_frame:
        huge_frame.destroy()
        huge_frame = None  # Reiniciar la variable huge_frame

    # Destroy and resets canva
    if canvas:
        canvas.destroy()
        canvas = None  # Reiniciar la variable canvas

    # Llamar a la función para mostrar la página principal
    show_main_page()

def show_main_page():
    """Shows main page."""
    root.state('zoomed')
    root.config(bg="#f7f7f7")
    hide_current_widgets()
    
    # Cargar y redimensionar el logo
    logo = Image.open("JuanRep\CorreoPL\PL\EvaMailer.png")
    logo = logo.resize((80, 80))
    logo_img = ImageTk.PhotoImage(logo)
    
    global logo_label
    logo_label = tk.Label(root, image=logo_img, bg="#f7f7f7")
    logo_label.image = logo_img
    logo_label.grid(row=0, column=1, columnspan=2, pady=20, sticky="n")
    
    global label
    label = tk.Label(root, text="WELCOME", font=('Helvetica', 24), bg="#f7f7f7", fg="black")
    label.grid(row=1, column=1, columnspan=2, pady=10, sticky="n")
    
    # Crear y configurar el marco Personalizado
    personalizado_frame = tk.Frame(root, bd=2, relief="ridge", bg="#f7f7f7")
    personalizado_frame.grid(row=2, column=1, padx=20, pady=20, sticky="nsew")
    
    global personalizado_title
    personalizado_title = tk.Label(personalizado_frame, text="PERSONALIZED", font=('Helvetica', 16), bg="#f7f7f7", fg="black")
    personalizado_title.pack(pady=10)
    
    # Crear y configurar el área de texto de descripción
    personalizado_description = tk.Text(personalizado_frame, wrap="word", height=10, width=50, bg="#f7f7f7", fg="black", borderwidth=0, padx=5, pady=5, font=("Helvetica", 10))
    personalizado_description.insert(tk.END, "Our tool allows you to send mass emails, but with a personal touch. Each message is automatically personalized, greeting the recipient by name and providing relevant information extracted from a data file. This means that each email not only reaches many, but also speaks directly to each person, improving the connection and response to our communications.")
    personalizado_description.config(state=tk.DISABLED)
    personalizado_description.pack()
    
    global personalizado_button
    personalizado_button = HyperlinkButton(personalizado_frame, text="Continue", bg="#f7f7f7", fg="#0066cc", activebackground="#f7f7f7", font=("Helvetica", 14), width=15, height=2, command=show_personalizado_page)
    personalizado_button.pack(pady=2, anchor='se')
    
    # Crear y configurar el marco Generar
    Generar_frame = tk.Frame(root, bd=2, relief="ridge", bg="#f7f7f7")
    Generar_frame.grid(row=2, column=2, padx=20, pady=20, sticky="nsew")
    
    global pdf_title
    pdf_title = tk.Label(Generar_frame, text="Generate Report ", font=('Helvetica', 16), bg="#f7f7f7", fg="black")
    pdf_title.pack(pady=10)
    
    # Crear y configurar el área de texto de descripción
    pdf_description = tk.Text(Generar_frame, wrap="word", height=10, width=50, bg="#f7f7f7", fg="black", borderwidth=0, padx=5, pady=5, font=("Helvetica", 10))
    pdf_description.insert(tk.END, "Our PDF download tools shows reports for every email that has been sent, Ensuring the destinatarie has readed the email. These reports, displayed by the click of a button, offer a clear analysis of campaign effectiveness, helping to optimize communication strategies and improve recipient engagement.")
    pdf_description.config(state=tk.DISABLED)
    pdf_description.pack()
    
    global pdf_button
    pdf_button = HyperlinkButton(Generar_frame, text="Continue", bg="#f7f7f7", fg="#0066cc", activebackground="#f7f7f7", font=("Helvetica", 14), width=15, height=2, command=show_email_data)
    pdf_button.pack(pady=2, anchor='se')

    # Configure grid row and column weights to make the layout responsive
    root.grid_rowconfigure(0, weight=1)
    root.grid_rowconfigure(1, weight=1)
    root.grid_rowconfigure(2, weight=1)
    root.grid_rowconfigure(3, weight=1)
    root.grid_columnconfigure(0, weight=1)
    root.grid_columnconfigure(1, weight=1)
    root.grid_columnconfigure(2, weight=1)
    root.grid_columnconfigure(3, weight=1)
    
    # Bind the resize event to adjust fonts dynamically
    root.bind("<Configure>", resize_fonts)
    
    # Display the copyright image
    show_copyright_image()

def replace_placeholders(email_body_template, row):
    """Reeplaces the placeholders woith actual CSV values."""
    for key, value in row.items():
        placeholder = f"${key}"
        email_body_template = email_body_template.replace(placeholder, str(value))
    return email_body_template

def show_variables(csv_file):
    """Shows the placeholders found"""
    global columns_listbox, variables_list, variable_frame, scrollbar
    if csv_file:
        encoding = detect_encoding(csv_file)
        try:
            with open(csv_file, 'r', encoding=encoding) as file:
                # Intentar detectar delimitador dinámicamente
                sample_line = file.readline()
                delimiter = detect_delimiter(sample_line)
                print(f"Detected delimiter: {delimiter}")  # Debugging print
                file.seek(0)  # Resetear el puntero del archivo

                reader = csv.DictReader(file, delimiter=delimiter)
                variables_list = reader.fieldnames
                
                if not variables_list:
                    raise ValueError("Placeholders not found.")
                
                print(f"Variables found: {variables_list}")  # Debugging print
        except Exception as e:
            print(f"Error reading CSV: {e}")
            return

        if variable_frame:
            variable_frame.destroy()
        variable_frame = tk.Frame(root, bg="#f7f7f7")
        variable_frame.grid(row=1, column=3, rowspan=4, padx=10, pady=10, sticky="nsew")

        if 'scrollbar' in globals() and scrollbar:
            scrollbar.destroy()

        scrollbar = tk.Scrollbar(variable_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        canvas = tk.Canvas(variable_frame, yscrollcommand=scrollbar.set, bg="#f7f7f7", bd=0, highlightthickness=0)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar.config(command=canvas.yview)

        button_frame = tk.Frame(canvas, bg="#f7f7f7")
        canvas.create_window((0, 0), window=button_frame, anchor='nw')

        for variable in variables_list:
            button = HyperlinkButton(button_frame, text=variable, command=lambda v=variable: insert_variable(v), bg="#f7f7f7")
            button.pack(anchor='w')

        button_frame.update_idletasks()
        canvas.config(scrollregion=canvas.bbox("all"))

def show_variable_buttons(variables, text_field, row=1, starting_column=2):
    """Muestra botones para insertar variables en el texto."""
    col_num = starting_column
    for var_name in variables:
        button = tk.Button(root, text=var_name, command=lambda name=var_name: insert_variable(text_field, name), bg="#f7f7f7", font=("Helvetica", 10))
        button.grid(row=row, column=col_num, padx=5, pady=5, sticky="w")
        col_num += 1

def ayudaMasivo():
    """Shows massive email tutorial."""
    messagebox.showinfo("General massive email tutorial", "Requirements: \n 1) Ensure that the excel file is saved as .csv \n 2) Ensure that the .csv file has the emails in a header called Mail\n If you want to attach files, click on the Attach general file button . Hold down the ctrl key to add more files\n")

def ayudaPersonalizado():
    """Shows massive personalized email tutorial."""
    messagebox.showinfo("Tutorial Envio Masivo Personalizado", "Requirements \n 1) Ensure that the excel file is saved as .csv \n 2) Ensure that the .csv file has the emails in a header called Mail \n 3) Ensure that the files to be attached are in a header called File \ n 4) That the files are set as an absolute file, example: C:\\Users\\USER\\Documents [This is achieved by right clicking, Location properties: then you put \\ and the name of the file with the extension example: algo.txt], to attach more files, separate it with a , \n When processing the file, some buttons will appear that will allow the mass email to send the unique information to the recipient")

def justify_text(widget, text, width):
    """Text justificator."""
    lines = []
    words = text.split()
    line = ""
    
    for word in words:
        if len(line + word) <= width:
            line += word + " "
        else:
            lines.append(line.strip())
            line = word + " "
    
    if line:
        lines.append(line.strip())
    
    for line in lines:
        words_in_line = line.split()
        if len(words_in_line) == 1:
            widget.insert(tk.END, words_in_line[0] + '\n')
        else:
            total_spaces = width - sum(len(word) for word in words_in_line)
            spaces_between_words = len(words_in_line) - 1
            if spaces_between_words > 0:
                spaces_per_gap = total_spaces // spaces_between_words
                extra_spaces = total_spaces % spaces_between_words
                justified_line = ""
                for i, word in enumerate(words_in_line[:-1]):
                    justified_line += word + " " * (spaces_per_gap + (1 if i < extra_spaces else 0))
                justified_line += words_in_line[-1]
            else:
                justified_line = words_in_line[0]
            widget.insert(tk.END, justified_line + '\n')

def on_variable_select(event):
    """Manages variable selection."""
    selected_index = columns_listbox.curselection()
    if selected_index:
        variable = columns_listbox.get(selected_index[0])
        messagebox.showwarning("Variable Selected", f"You selected variable: {variable}")
        insert_variable(cuerpoP_textfield, variable)

def on_resize(event):
    """Readjusts Screen."""
    global widget_refs

    window_width = root.winfo_width()
    window_height = root.winfo_height()

    screen_width = root.winfo_screenwidth() 
    screen_height = root.winfo_screenheight()

    is_maximized = (window_width == screen_width and window_height == screen_height)

    if 'subject_entry' in widget_refs:
        if is_maximized:
            widget_refs['subject_entry'].config(width=None, height=None)
        else:
            widget_refs['subject_entry'].config(width=max(40, int(window_width / 20)), height=2)
        
    if 'body_text' in widget_refs:
        if is_maximized:
            widget_refs['body_text'].config(width=None, height=None)
        else:
            widget_refs['body_text'].config(width=max(60, int(window_width / 15)), height=10)

    root.update_idletasks()

def detect_delimiter(sample_line):
    """Detects the delimiter used in a CSV file based on the most common delimiter."""
    common_delimiters = [',', ';', '\t', '|', ':']
    delimiter_counts = Counter({delimiter: sample_line.count(delimiter) for delimiter in common_delimiters})
    return delimiter_counts.most_common(1)[0][0]


def show_copyright_image():
    """Shows copyright logo."""
    copyright_img = Image.open("JuanRep\CorreoPL\PL\logopl.png")
    copyright_img = copyright_img.resize((50, 50), Image.Resampling.LANCZOS)
    copyright_photo = ImageTk.PhotoImage(copyright_img)
    copyright_label = tk.Label(root, image=copyright_photo, bg="#f7f7f7")
    copyright_label.image = copyright_photo
    copyright_label.grid(row=5, column=2, sticky="se", padx=10, pady=10)

# Inicialización de la aplicación
root = tk.Tk()
root.title("Email Sender")


show_main_page()
root.mainloop()