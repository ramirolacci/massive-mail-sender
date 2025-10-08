import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from tkinter import colorchooser
from tkinter import ttk
from tkinter import font as tkfont
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
import validators
import threading
import webbrowser
import tempfile
from urllib.parse import quote
from pathlib import Path
import os
import re
import html as htmllib

# --- Opcional: Google Sheets ---
try:
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials
    HAS_GS = True
except ImportError:
    HAS_GS = False
try:
    from PyPDF2 import PdfReader
    HAS_PDF = True
except ImportError:
    HAS_PDF = False

# --- Configuración SMTP por defecto (oculta en UI) ---
DEFAULT_SMTP_SERVER = 'mail.migusto.com.ar'
DEFAULT_SMTP_PORT = 587
DEFAULT_USE_SSL = False
DEFAULT_USE_STARTTLS = True
DEFAULT_SMTP_USER = 'news@migusto.com.ar'
DEFAULT_SMTP_PASS = 'Promociones2025@'

# --- Google Sheets credenciales por defecto ---
GS_CREDENTIALS_PATH = os.path.join(os.path.dirname(__file__), 'credentials.json')

# --- Funciones auxiliares ---
def cargar_emails_desde_excel(path):
    try:
        ext = os.path.splitext(path)[1].lower()
        if ext == '.csv':
            df = pd.read_csv(path)
        elif ext == '.xlsx':
            df = pd.read_excel(path)
        elif ext == '.pdf':
            if not HAS_PDF:
                messagebox.showerror('Error', 'PyPDF2 no está instalado. Agrega PyPDF2 a requirements.txt.')
                return []
            try:
                reader = PdfReader(path)
                full_text = ''.join((page.extract_text() or '') for page in reader.pages)
            except Exception as e:
                messagebox.showerror('Error', f'No se pudo leer el PDF: {e}')
                return []
            # Extraer emails por regex y validar
            raw = re.findall(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", full_text)
            emails = [e for e in dict.fromkeys(raw) if validators.email(e)]
            return emails
        else:
            messagebox.showerror('Error', 'Formato no permitido. Usa .xlsx, .csv o .pdf')
            return []

        if not len(df.columns):
            messagebox.showerror('Error', 'El archivo no contiene columnas.')
            return []
        header = str(df.columns[0]).strip().lower()
        if header != 'email':
            messagebox.showerror('Error', "La primera columna debe llamarse 'email'.")
            return []
        emails = df.iloc[:, 0].dropna().astype(str).tolist()
        return emails
    except Exception as e:
        messagebox.showerror('Error', f'No se pudo leer el archivo: {e}')
        return []

def cargar_emails_desde_gsheet(sheet_id, range_name, creds_path):
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name(creds_path, scope)
        client = gspread.authorize(creds)
        sheet = client.open_by_key(sheet_id)
        ws = sheet.worksheet(range_name) if range_name else sheet.sheet1
        header_row = ws.row_values(1)
        if not header_row or not header_row[0] or header_row[0].strip().lower() != 'email':
            messagebox.showerror('Error', "La primera columna (A1) debe llamarse 'email'.")
            return []
        data = ws.col_values(1)
        data = data[1:]  # omitir encabezado
        return [x for x in data if x and validators.email(x)]
    except Exception as e:
        messagebox.showerror('Error', f'No se pudo leer Google Sheet: {e}')
        return []

def validar_emails(lista):
    return [e for e in lista if validators.email(e)]

def enviar_mails(
    smtp_user,
    smtp_pass,
    emails,
    asunto,
    mensaje_texto,
    status_callback,
    smtp_server,
    smtp_port,
    use_ssl,
    use_starttls,
    html_body=None,
    image_paths=None,
):
    try:
        if use_ssl:
            server = smtplib.SMTP_SSL(smtp_server, smtp_port)
        else:
            server = smtplib.SMTP(smtp_server, smtp_port)
            if use_starttls:
                server.starttls()
        server.login(smtp_user, smtp_pass)
        enviados, fallidos = 0, 0
        total = len(emails)
        for i, email in enumerate(emails, 1):
            status_callback(f'Enviando a {email} ({i}/{total})...')
            try:
                if html_body is None and not image_paths:
                    msg = MIMEText(mensaje_texto, 'plain', 'utf-8')
                else:
                    root = MIMEMultipart('related')
                    alt = MIMEMultipart('alternative')
                    root.attach(alt)
                    alt.attach(MIMEText(mensaje_texto, 'plain', 'utf-8'))
                    cuerpo_html = html_body if html_body is not None else f"<p>{mensaje_texto}</p>"
                    alt.attach(MIMEText(cuerpo_html, 'html', 'utf-8'))
                    if image_paths:
                        for idx, path in enumerate(image_paths, 1):
                            try:
                                with open(path, 'rb') as f:
                                    img = MIMEImage(f.read())
                                cid = f'image{idx}'
                                img.add_header('Content-ID', f'<{cid}>')
                                img.add_header('Content-Disposition', 'inline', filename=os.path.basename(path))
                                root.attach(img)
                            except Exception:
                                pass
                    msg = root
                msg['Subject'] = asunto
                msg['From'] = smtp_user
                msg['To'] = email
                server.sendmail(smtp_user, email, msg.as_string())
                enviados += 1
            except Exception:
                fallidos += 1
        server.quit()
        return enviados, fallidos
    except smtplib.SMTPAuthenticationError:
        messagebox.showerror('Error', 'Error de autenticación SMTP. Verifica email y contraseña.')
        return 0, len(emails)
    except Exception as e:
        messagebox.showerror('Error', f'Error general de envío: {e}')
        return 0, len(emails)

# --- Interfaz gráfica principal ---
class MassiveMailSender(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Massive Mail Sender')
        self.geometry('900x720')
        self.minsize(900, 720)
        self.resizable(True, True)
        self.config(bg='#23272f')
        self.emails = []
        self.image_paths = []
        self.next_image_placeholder = 1
        self.template_var = tk.StringVar(value='Promocional')
        self.grad_start = '#ffb300'  # amarillo anaranjado
        self.grad_end = '#ff3d00'    # rojo
        # Logo por defecto si existe en ./public/@Logo Mi Gusto 2025.png
        default_logo = os.path.join(os.path.dirname(__file__), 'public', '@Logo Mi Gusto 2025.png')
        self.logo_path = default_logo if os.path.isfile(default_logo) else None
        self._setup_styles()
        self.crear_widgets()

    def _apply_template(self, content_html):
        # Varias plantillas: Promocional, Corporativa, Minimalista, Vanilla, Navidad, Halloween
        tpl = (self.template_var.get() or 'Corporativo').strip()
        grad_start = self.grad_start
        grad_end = self.grad_end
        # header content: logo si existe
        logo_html = ''
        if self.logo_path:
            p = Path(self.logo_path).absolute().as_posix()
            logo_src = f"file:///{quote(p, safe='/:')}"
            logo_html = f"<img alt='logo' src='{logo_src}' style='height:28px; display:block'>"
        else:
            logo_html = "<div style='font-weight:700; font-size:20px; color:#fff'>MiGusto</div>"

        # Vanilla → sin envoltorio HTML (solo el contenido tal cual)
        if tpl == 'Vanilla':
            return f"{content_html}"

        # Corporativa → header centrado con logo
        if tpl == 'Corporativa':
            return f"""
            <meta charset='utf-8'>
            <body style='margin:0; background:#ffffff;'>
            <table role='presentation' cellpadding='0' cellspacing='0' border='0' width='100%'>
              <tr>
                <td align='center' style='padding:80px 24px;'>
                  <div style='display:inline-block'>{logo_html}</div>
                </td>
              </tr>
            </table>
            </body>
            """

        # Promocional → plantilla con degradado (antes Corporativo)
        if tpl == 'Promocional':
            return f"""
            <meta charset='utf-8'>
            <body style='margin:0; background:#1f2023;'>
            <table role='presentation' cellpadding='0' cellspacing='0' border='0' width='100%'>
              <tr>
                <td align='center' style='padding:0;'>
                  <table role='presentation' cellpadding='0' cellspacing='0' border='0' width='720' style='width:720px; max-width:720px; background:#23272f; font-family:Segoe UI, Arial, sans-serif; color:#e0e6f0;'>
                    <tr>
                      <td style='background:linear-gradient(90deg,{grad_start},{grad_end}); padding:18px 24px;'>
                        {logo_html}
                      </td>
                    </tr>
                    <tr>
                      <td style='padding:24px'>
                        {content_html}
                      </td>
                    </tr>
                    <tr>
                      <td style='padding:18px 24px; background:#2c313c; font-size:12px; color:#b5c2d6;'>
                        Recibiste este mensaje porque estás suscripto a nuestras promociones.
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>
            </body>
            """

        # Navidad → tonos rojo/verde con header festivo
        if tpl == 'Navidad':
            return f"""
            <meta charset='utf-8'>
            <body style='margin:0; background:#f5f7fa;'>
            <table role='presentation' cellpadding='0' cellspacing='0' border='0' width='100%'>
              <tr>
                <td align='center' style='padding:0;'>
                  <table role='presentation' cellpadding='0' cellspacing='0' border='0' width='720' style='width:720px; max-width:720px; background:#ffffff; font-family:Segoe UI, Arial, sans-serif; color:#1c2a33; border-radius:10px; overflow:hidden;'>
                    <tr>
                      <td style='background:linear-gradient(90deg,#0aa83f,#d32f2f); padding:18px 24px;'>
                        {logo_html}
                      </td>
                    </tr>
                    <tr>
                      <td style='padding:24px; line-height:1.6;'>
                        {content_html}
                      </td>
                    </tr>
                    <tr>
                      <td style='padding:14px 24px; background:#0aa83f0d; color:#2b3a42; font-size:12px;'>
                        ¡Felices Fiestas! ✨
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>
            </body>
            """

        # Halloween → tonos negro/naranja con header
        if tpl == 'Halloween':
            return f"""
            <meta charset='utf-8'>
            <body style='margin:0; background:#0a0a0a;'>
            <table role='presentation' cellpadding='0' cellspacing='0' border='0' width='100%'>
              <tr>
                <td align='center' style='padding:0;'>
                  <table role='presentation' cellpadding='0' cellspacing='0' border='0' width='720' style='width:720px; max-width:720px; background:#121212; font-family:Segoe UI, Arial, sans-serif; color:#f1f1f1; border:1px solid #2a2a2a; border-radius:10px;'>
                    <tr>
                      <td style='background:linear-gradient(90deg,#111,#ff6a00); padding:18px 24px;'>
                        {logo_html}
                      </td>
                    </tr>
                    <tr>
                      <td style='padding:24px; line-height:1.7;'>
                        {content_html}
                      </td>
                    </tr>
                    <tr>
                      <td style='padding:14px 24px; background:#1a1a1a; color:#ffa64d; font-size:12px;'>
                        ¡Feliz Halloween! 🎃
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>
            </body>
            """

        # Minimalista → header centrado con logo
        return f"""
            <meta charset='utf-8'>
            <body style='margin:0; background:#0b0d10;'>
            <table role='presentation' cellpadding='0' cellspacing='0' border='0' width='100%'>
              <tr>
                <td align='center' style='padding:24px;'>
                  <table role='presentation' cellpadding='0' cellspacing='0' border='0' width='720' style='width:720px; max-width:720px; background:#101318; font-family:Segoe UI, Arial, sans-serif; color:#e8ecf3; border:1px solid #2a2f3a; border-radius:10px;'>
                    <tr>
                      <td style='padding:18px 24px; border-bottom:1px solid #2a2f3a; text-align:center;'>{logo_html}</td>
                    </tr>
                    <tr>
                      <td style='padding:24px; line-height:1.6; color:#d4dae4;'>
                        {content_html}
                      </td>
                    </tr>
                    <tr>
                      <td style='padding:14px 24px; font-size:12px; color:#9aa5b1; background:#161a22; border-top:1px solid #2a2f3a; border-bottom-left-radius:10px; border-bottom-right-radius:10px;'>
                        © {Path(__file__).parent.name} · Todos los derechos reservados
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>
            </body>
            """

        # Corporativo (default)
        return f"""
        <meta charset='utf-8'>
        <body style='margin:0; background:#1f2023;'>
        <table role='presentation' cellpadding='0' cellspacing='0' border='0' width='100%'>
          <tr>
            <td align='center' style='padding:0;'>
              <table role='presentation' cellpadding='0' cellspacing='0' border='0' width='720' style='width:720px; max-width:720px; background:#23272f; font-family:Segoe UI, Arial, sans-serif; color:#e0e6f0;'>
                <tr>
                  <td style='background:linear-gradient(90deg,{grad_start},{grad_end}); padding:18px 24px;'>
                    {logo_html}
                  </td>
                </tr>
                <tr>
                  <td style='padding:24px'>
                    {content_html}
                  </td>
                </tr>
                <tr>
                  <td style='padding:18px 24px; background:#2c313c; font-size:12px; color:#b5c2d6;'>
                    Recibiste este mensaje porque estás suscripto a nuestras promociones.
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
        </body>
        """

    def _build_preview_html(self, preview_local=False):
        # Convertir contenido del editor (con tags visuales) a HTML real para la vista previa
        html_body = self._editor_content_to_html(preview_local=preview_local)
        # Reemplazar placeholders {{imageN}}
        for idx, path in enumerate(self.image_paths, start=1):
            if preview_local and path:
                p = Path(path).absolute().as_posix()
                src = f"file:///{quote(p, safe='/:')}"
            else:
                src = f"cid:image{idx}"
            html_body = html_body.replace(f"{{{{image{idx}}}}}", f"<img src='{src}' style='max-width:100%; height:auto; border-radius:8px' alt='img {idx}'>")
        # eliminar placeholders sobrantes sin imagen
        html_body = re.sub(r'\{\{image\d+\}\}', '', html_body)
        # Construir asunto usando el mismo motor de tags
        subject_html = self._apply_inline_tags_widget(self.txt_asunto) or htmllib.escape(self.txt_asunto.get('1.0','end-1c').strip())
        wrapped = self._apply_template(f"<h2 style='margin:0 0 12px'>{subject_html}</h2><div style='margin-bottom:12px'>{html_body}</div>")
        return wrapped

    def _setup_styles(self):
        style = ttk.Style(self)
        try:
            style.theme_use('clam')
        except tk.TclError:
            pass
        # Estilo base para botones
        style.configure('TButton', padding=6)
        # Botón principal
        style.configure('Primary.TButton', foreground='#ffffff', background='#4f8cff')
        style.map('Primary.TButton',
                  background=[('active', '#3d79f2'), ('disabled', '#5f93ff')],
                  foreground=[('disabled', '#e0e6f0')])
        # Botón secundario
        style.configure('Secondary.TButton', foreground='#ffffff', background='#357ae8')
        style.map('Secondary.TButton', background=[('active', '#2d68c7')])
        # Botón de énfasis (Enviar)
        style.configure('Accent.TButton', foreground='#ffffff', background='#ff6a00', font=('Segoe UI', 12, 'bold'), padding=8)
        style.map('Accent.TButton', background=[('active', '#ff3d00')])

    def _load_header_logo(self):
        if not self.logo_path or not os.path.isfile(self.logo_path):
            return None
        try:
            img = tk.PhotoImage(file=self.logo_path)
            h = img.height()
            if h > 28:
                factor = max(1, int(h / 28))
                img = img.subsample(factor, factor)
            return img
        except Exception:
            return None

    def _ranges_for_tag(self, tag):
        # Conservado para compatibilidad: usa el editor de mensaje
        return self._ranges_for_tag_in(self.txt_mensaje, tag)

    def _ranges_for_tag_in(self, widget, tag):
        try:
            ranges = widget.tag_ranges(tag)
        except tk.TclError:
            return []
        pairs = list(zip(ranges[0::2], ranges[1::2]))
        norm = []
        for start, end in pairs:
            s = widget.index(start)
            e = widget.index(end)
            norm.append((s, e))
        return norm

    def _apply_inline_tags(self, text):
        # Conservado para compatibilidad: usa editor de mensaje
        return self._apply_inline_tags_widget(self.txt_mensaje)

    def _apply_inline_tags_widget(self, widget):
        # Genera un mapa de posiciones -> lista de tags que abren/cierran
        positions = {}
        def add(pos, token):
            key = widget.index(pos)
            positions.setdefault(key, []).append(token)

        tag_to_html = {
            'bold': ('<b>', '</b>'),
            'italic': ('<i>', '</i>'),
            'underline': ('<u>', '</u>'),
            'center': ("<div style='text-align:center'>", '</div>'),
        }
        for tag, (open_t, close_t) in tag_to_html.items():
            for start, end in self._ranges_for_tag_in(widget, tag):
                add(start, open_t)
                add(end, close_t)

        for tag in widget.tag_names():
            if tag.startswith('color_'):
                color = tag.split('color_')[-1]
                open_t = f"<span style='color:{color}'>"
                close_t = '</span>'
                for start, end in self._ranges_for_tag_in(widget, tag):
                    add(start, open_t)
                    add(end, close_t)
            if tag.startswith('size_'):
                size = tag.split('size_')[-1]
                open_t = f"<span style='font-size:{size}px'>"
                close_t = '</span>'
                for start, end in self._ranges_for_tag_in(widget, tag):
                    add(start, open_t)
                    add(end, close_t)

        start_idx = widget.index('1.0')
        end_idx = widget.index('end-1c')
        cur = start_idx
        html_parts = []
        while cur != end_idx:
            if cur in positions:
                for token in positions[cur]:
                    html_parts.append(token)
            nxt = widget.index(f"{cur} +1c")
            ch = widget.get(cur, nxt)
            if ch == '\n':
                html_parts.append('<br>')
            else:
                html_parts.append(htmllib.escape(ch))
            cur = nxt
        return ''.join(html_parts)

    def _editor_content_to_html(self, preview_local=False):
        return self._apply_inline_tags_widget(self.txt_mensaje)

    def crear_widgets(self):
        # Layout principal con expansión
        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # Header superior con logo
        header = tk.Frame(self, bg='#2c313c')
        header.grid(row=0, column=0, sticky='ew')
        header.grid_columnconfigure(0, weight=0)
        header.grid_columnconfigure(1, weight=1)  # espacio flexible
        header.grid_columnconfigure(2, weight=0)
        # Título a la izquierda
        tk.Label(header, text='Massive Mail Sender', bg='#2c313c', fg='#e0e6f0', font=('Segoe UI', 16, 'bold')).grid(row=0, column=0, padx=20, pady=10, sticky='w')
        # Logo / MiGusto a la derecha
        self.header_logo_img = self._load_header_logo()
        if self.header_logo_img:
            tk.Label(header, image=self.header_logo_img, bg='#2c313c').grid(row=0, column=2, padx=20, pady=10, sticky='e')
        else:
            tk.Label(header, text='MiGusto', bg='#2c313c', fg='#e0e6f0', font=('Segoe UI', 16, 'bold')).grid(row=0, column=2, padx=20, pady=10, sticky='e')

        # Frame superior para carga de emails
        frame_carga = tk.Frame(self, bg='#2c313c')
        frame_carga.grid(row=1, column=0, sticky='ew', padx=20, pady=10)
        frame_carga.grid_columnconfigure(0, weight=1)

        tk.Label(frame_carga, text='Destinatarios:', bg='#2c313c', fg='#e0e6f0', font=('Segoe UI', 12, 'bold')).grid(row=0, column=0, sticky='w', pady=5)
        self.txt_emails = scrolledtext.ScrolledText(frame_carga, width=60, height=6, bg='#23272f', fg='#e0e6f0', insertbackground='#e0e6f0')
        self.txt_emails.grid(row=1, column=0, columnspan=4, pady=5, sticky='ew')

        btn_excel = ttk.Button(frame_carga, text='Cargar archivo (XLSX/CSV/PDF)', command=self.cargar_excel, style='Primary.TButton')
        btn_excel.grid(row=2, column=0, pady=5, sticky='w')

        btn_limpiar = ttk.Button(frame_carga, text='Limpiar', command=self.limpiar_emails, style='Secondary.TButton')
        btn_limpiar.grid(row=2, column=0, padx=(220,0), pady=5, sticky='w')

        # Panel SMTP a la derecha del frame de carga
        smtp_panel = tk.Frame(frame_carga, bg='#2c313c')
        smtp_panel.grid(row=0, column=2, rowspan=3, sticky='ne', padx=(10,0))
        smtp_panel.grid_columnconfigure(0, weight=0)
        smtp_panel.grid_columnconfigure(1, weight=0)
        # Variables UI
        self.var_smtp_user = tk.StringVar(value=DEFAULT_SMTP_USER)
        self.var_smtp_pass = tk.StringVar(value=DEFAULT_SMTP_PASS)
        self.var_smtp_server = tk.StringVar(value=DEFAULT_SMTP_SERVER)
        # Puerto como StringVar para usar en Combobox (valores 587/465)
        self.var_smtp_port = tk.StringVar(value=str(DEFAULT_SMTP_PORT))
        self.var_use_ssl = tk.BooleanVar(value=DEFAULT_USE_SSL)
        self.var_use_starttls = tk.BooleanVar(value=DEFAULT_USE_STARTTLS)
        # Disposición vertical solicitada
        ttk.Label(smtp_panel, text='Email', background='#2c313c', foreground='#e0e6f0').grid(row=0, column=0, sticky='w')
        ttk.Entry(smtp_panel, textvariable=self.var_smtp_user, width=32).grid(row=0, column=1, sticky='w')
        ttk.Label(smtp_panel, text='Contraseña', background='#2c313c', foreground='#e0e6f0').grid(row=1, column=0, sticky='w', pady=(6,0))
        ttk.Entry(smtp_panel, textvariable=self.var_smtp_pass, width=32, show='•').grid(row=1, column=1, sticky='w', pady=(6,0))
        ttk.Label(smtp_panel, text='Servidor', background='#2c313c', foreground='#e0e6f0').grid(row=2, column=0, sticky='w', pady=(6,0))
        ttk.Entry(smtp_panel, textvariable=self.var_smtp_server, width=32).grid(row=2, column=1, sticky='w', pady=(6,0))
        ttk.Label(smtp_panel, text='Puerto', background='#2c313c', foreground='#e0e6f0').grid(row=3, column=0, sticky='w', pady=(6,0))
        cb_port = ttk.Combobox(smtp_panel, values=['587','465'], state='readonly', width=8, textvariable=self.var_smtp_port)
        cb_port.grid(row=3, column=1, sticky='w', pady=(6,0))
        def _on_port_change(event=None):
            p = (self.var_smtp_port.get() or '').strip()
            if p == '465':
                self.var_use_ssl.set(True)
                self.var_use_starttls.set(False)
            elif p == '587':
                self.var_use_ssl.set(False)
                self.var_use_starttls.set(True)
        cb_port.bind('<<ComboboxSelected>>', _on_port_change)
        _on_port_change()
        ttk.Checkbutton(smtp_panel, text='TLS', variable=self.var_use_starttls).grid(row=4, column=1, sticky='w', pady=(6,0))
        ttk.Checkbutton(smtp_panel, text='SSL', variable=self.var_use_ssl).grid(row=5, column=1, sticky='w', pady=(2,0))

        # Frame para asunto, mensaje e imágenes (sin campos SMTP/credenciales)
        frame_form = tk.Frame(self, bg='#2c313c')
        frame_form.grid(row=2, column=0, sticky='nsew', padx=20, pady=10)
        frame_form.grid_columnconfigure(1, weight=1)
        frame_form.grid_columnconfigure(2, weight=1)
        frame_form.grid_columnconfigure(3, weight=1)
        # permitir que la fila del editor (row=2) se expanda
        frame_form.grid_rowconfigure(2, weight=1)

        tk.Label(frame_form, text='Asunto:', bg='#2c313c', fg='#e0e6f0').grid(row=0, column=0, sticky='e', pady=5)
        # Editor de asunto con soporte de tags (estilo similar al cuerpo)
        subject_font = tkfont.Font(family='Segoe UI', size=14)
        self.txt_asunto = tk.Text(frame_form, height=3, bg='#23272f', fg='#e0e6f0', insertbackground='#e0e6f0', font=subject_font, wrap='word')
        self.txt_asunto.grid(row=0, column=1, columnspan=3, pady=5, padx=5, sticky='ew')
        # Tags para asunto
        self.txt_asunto.tag_configure('bold', font=(subject_font.actual('family'), subject_font.actual('size'), 'bold'))
        self.txt_asunto.tag_configure('italic', font=(subject_font.actual('family'), subject_font.actual('size'), 'italic'))
        self.txt_asunto.tag_configure('underline', underline=1)
        self.txt_asunto.tag_configure('center', justify='center')

        # Barra de herramientas de formato (aplica estilos visuales en el editor)
        toolbar = tk.Frame(frame_form, bg='#2c313c')
        toolbar.grid(row=1, column=1, columnspan=3, sticky='ew', padx=5)
        for i in range(0, 8):
            toolbar.grid_columnconfigure(i, weight=0)
        toolbar.grid_columnconfigure(8, weight=1)

        # Estilos visuales para el editor (no HTML, solo presentación)
        editor_font = tkfont.Font(family='Segoe UI', size=12)
        self.txt_mensaje = scrolledtext.ScrolledText(frame_form, width=70, height=16, bg='#23272f', fg='#e0e6f0', insertbackground='#e0e6f0', font=editor_font, wrap='word')
        # Tags
        self.txt_mensaje.tag_configure('bold', font=editor_font.copy())
        self.txt_mensaje.tag_config('bold', font=(editor_font.actual('family'), editor_font.actual('size'), 'bold'))
        self.txt_mensaje.tag_configure('italic', font=(editor_font.actual('family'), editor_font.actual('size'), 'italic'))
        self.txt_mensaje.tag_configure('underline', underline=1)
        self.txt_mensaje.tag_configure('center', justify='center')

        # Recordar último editor activo (asunto o mensaje)
        self._last_editor = self.txt_mensaje
        self.txt_mensaje.bind('<FocusIn>', lambda e: setattr(self, '_last_editor', self.txt_mensaje))
        self.txt_asunto.bind('<FocusIn>', lambda e: setattr(self, '_last_editor', self.txt_asunto))

        def active_editor():
            return getattr(self, '_last_editor', self.txt_mensaje)

        def toggle_tag(tag):
            ed = active_editor()
            try:
                start = ed.index('sel.first')
                end = ed.index('sel.last')
            except tk.TclError:
                return
            if tag in ed.tag_names('sel.first'):
                ed.tag_remove(tag, start, end)
            else:
                ed.tag_add(tag, start, end)

        def cmd_bold():
            toggle_tag('bold')
        def cmd_italic():
            toggle_tag('italic')
        def cmd_underline():
            toggle_tag('underline')
        def cmd_center():
            toggle_tag('center')
        def cmd_list_ul():
            ed = active_editor()
            try:
                start = ed.index('sel.first linestart')
                end = ed.index('sel.last lineend')
            except tk.TclError:
                return
            lines = ed.get(start, end).split('\n')
            bullet_lines = ['• ' + l if not l.strip().startswith('•') else l for l in lines]
            ed.delete(start, end)
            ed.insert(start, '\n'.join(bullet_lines))
        def cmd_list_ol():
            ed = active_editor()
            try:
                start = ed.index('sel.first linestart')
                end = ed.index('sel.last lineend')
            except tk.TclError:
                return
            lines = [l for l in ed.get(start, end).split('\n')]
            numbered = [f"{i+1}. {l}" for i, l in enumerate(lines)]
            ed.delete(start, end)
            ed.insert(start, '\n'.join(numbered))
        def cmd_link():
            ed = active_editor()
            try:
                start = ed.index('sel.first')
                end = ed.index('sel.last')
            except tk.TclError:
                return
            url = tk.simpledialog.askstring('Agregar enlace', 'URL destino:')
            if not url:
                return
            # Representación visual de link
            ed.tag_configure('link', foreground='#7ecfff', underline=1)
            ed.tag_add('link', start, end)
        def cmd_color():
            ed = active_editor()
            picked = colorchooser.askcolor(title='Elegir color de texto')
            if not picked or not picked[1]:
                return
            try:
                start = ed.index('sel.first')
                end = ed.index('sel.last')
            except tk.TclError:
                return
            tag_name = f"color_{picked[1]}"
            ed.tag_configure(tag_name, foreground=picked[1])
            ed.tag_add(tag_name, start, end)
        def cmd_size():
            ed = active_editor()
            sizes = ['12','14','16','18','20','24','28']
            dlg = tk.Toplevel(self)
            dlg.title('Tamaño de texto')
            cb = ttk.Combobox(dlg, values=sizes, state='readonly')
            cb.current(2)
            cb.pack(padx=10, pady=10)
            def ok():
                sz = cb.get()
                try:
                    start = ed.index('sel.first')
                    end = ed.index('sel.last')
                except tk.TclError:
                    dlg.destroy()
                    return
                tag_name = f"size_{sz}"
                # Ajuste de fuente según editor activo
                base_font = subject_font if ed is self.txt_asunto else editor_font
                ed.tag_configure(tag_name, font=(base_font.actual('family'), int(sz)))
                ed.tag_add(tag_name, start, end)
                dlg.destroy()
            ttk.Button(dlg, text='OK', command=ok).pack(pady=5)

        ttk.Button(toolbar, text='B', width=3, command=cmd_bold).grid(row=0, column=0, padx=2)
        ttk.Button(toolbar, text='I', width=3, command=cmd_italic).grid(row=0, column=1, padx=2)
        ttk.Button(toolbar, text='U', width=3, command=cmd_underline).grid(row=0, column=2, padx=2)
        ttk.Button(toolbar, text='Color', command=cmd_color).grid(row=0, column=3, padx=4)
        ttk.Button(toolbar, text='Tamaño', command=cmd_size).grid(row=0, column=4, padx=4)
        ttk.Button(toolbar, text='Centrar', command=cmd_center).grid(row=0, column=5, padx=4)
        ttk.Button(toolbar, text='Lista •', command=cmd_list_ul).grid(row=0, column=6, padx=4)
        # Botón de lista numerada eliminado por solicitud

        def cmd_preview():
            html = self._build_preview_html(preview_local=True)
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.html')
            tmp.write(html.encode('utf-8'))
            tmp.close()
            webbrowser.open(f'file://{tmp.name}')
        ttk.Button(toolbar, text='Vista previa', command=cmd_preview).grid(row=0, column=8, padx=6, sticky='w')

        # Plantillas y posición de imágenes
        # Renombradas: Promocional (antes corporativo degradado), Corporativa (MI GUSTO centrado), Minimalista (tarjeta blanca)
        opts = ['Vanilla', 'Promocional', 'Corporativa', 'Minimalista', 'Navidad', 'Halloween']
        ttk.Label(toolbar, text='Plantilla:').grid(row=0, column=9, padx=(20,2))
        cb_tpl = ttk.Combobox(toolbar, values=opts, state='readonly', textvariable=self.template_var, width=12)
        cb_tpl.grid(row=0, column=10, padx=2)
        def insert_next_image_placeholder():
            ph = f"{{{{image{self.next_image_placeholder}}}}}"
            cur = self.txt_mensaje.index('insert')
            self.txt_mensaje.insert(cur, ph)
            self.next_image_placeholder += 1
        ttk.Button(toolbar, text='Insertar imagen', command=insert_next_image_placeholder).grid(row=0, column=11, padx=6)

        # Controles de branding (ocultos por pedido del usuario)

        tk.Label(frame_form, text='Mensaje:', bg='#2c313c', fg='#e0e6f0').grid(row=2, column=0, sticky='ne', pady=5)
        self.txt_mensaje.grid(row=2, column=1, columnspan=3, pady=5, padx=5, sticky='nsew')

        # Controles de imágenes, alineados debajo del editor
        imgs_row = 3
        btn_imgs = ttk.Button(frame_form, text='Agregar imágenes…', command=self.cargar_imagenes, style='Primary.TButton')
        btn_imgs.grid(row=imgs_row, column=1, sticky='w', pady=(0,5))
        self.lbl_imgs = tk.Label(frame_form, text='Sin imágenes', bg='#2c313c', fg='#e0e6f0')
        self.lbl_imgs.grid(row=imgs_row, column=2, columnspan=2, sticky='w', pady=(0,5))


        # Botón de envío
        self.btn_enviar = ttk.Button(self, text='Enviar', command=self.enviar, style='Accent.TButton', width=20)
        self.btn_enviar.grid(row=3, column=0, pady=15)

        # Estado
        self.lbl_estado = tk.Label(self, text='', bg='#23272f', fg='#7ecfff', font=('Segoe UI', 11, 'italic'))
        self.lbl_estado.grid(row=4, column=0, pady=5)

    def cargar_excel(self):
        tipos = [('Excel (.xlsx)', '*.xlsx'), ('CSV (.csv)', '*.csv'), ('PDF (.pdf)', '*.pdf')]
        path = filedialog.askopenfilename(title='Seleccionar archivo', filetypes=tipos)
        if not path:
            return
        emails = cargar_emails_desde_excel(path)
        self.txt_emails.delete('1.0', tk.END)
        self.txt_emails.insert('1.0', '\n'.join(emails))
        self.lbl_estado.config(text=f'{len(emails)} emails cargados.')

    def cargar_gsheet(self):
        if not HAS_GS:
            messagebox.showerror('Error', 'gspread no está instalado.')
            return
        top = tk.Toplevel(self)
        top.title('Cargar Google Sheet')
        tk.Label(top, text='ID de la hoja:').pack()
        entry_id = tk.Entry(top, width=50)
        entry_id.pack()
        # Usamos worksheet por defecto y credentials.json por defecto
        def cargar():
            sheet_id = entry_id.get().strip()
            emails = cargar_emails_desde_gsheet(sheet_id, None, GS_CREDENTIALS_PATH)
            self.txt_emails.delete('1.0', tk.END)
            self.txt_emails.insert('1.0', '\n'.join(emails))
            self.lbl_estado.config(text=f'{len(emails)} emails cargados de Google Sheets.')
            top.destroy()
        tk.Button(top, text='Cargar', command=cargar).pack(pady=5)

    def cargar_imagenes(self):
        paths = filedialog.askopenfilenames(title='Seleccionar imágenes', filetypes=[('Imágenes', '*.png;*.jpg;*.jpeg;*.gif')])
        if not paths:
            return
        self.image_paths = list(paths)
        self.lbl_imgs.config(text=f'{len(self.image_paths)} imagen(es) seleccionada(s).')

    def _select_logo(self):
        path = filedialog.askopenfilename(title='Seleccionar logo', filetypes=[('Imágenes', '*.png;*.jpg;*.jpeg;*.gif;*.svg')])
        if path:
            self.logo_path = path

    def _select_gradient(self):
        c1 = colorchooser.askcolor(title='Color inicio (amarillo/naranja)')
        if not c1 or not c1[1]:
            return
        c2 = colorchooser.askcolor(title='Color fin (rojo/naranja)')
        if not c2 or not c2[1]:
            return
        self.grad_start = c1[1]
        self.grad_end = c2[1]

    def limpiar_emails(self):
        self.txt_emails.delete('1.0', tk.END)
        self.lbl_estado.config(text='')

    def enviar(self):
        emails = [e.strip() for e in self.txt_emails.get('1.0', tk.END).splitlines() if e.strip()]
        emails = validar_emails(emails)
        if not emails:
            messagebox.showerror('Error', 'No hay emails válidos para enviar.')
            return
        # Tomar valores actuales desde la UI (panel SMTP)
        server_host = (self.var_smtp_server.get() if hasattr(self, 'var_smtp_server') else DEFAULT_SMTP_SERVER).strip()
        try:
            server_port = int(self.var_smtp_port.get()) if hasattr(self, 'var_smtp_port') else DEFAULT_SMTP_PORT
        except Exception:
            server_port = DEFAULT_SMTP_PORT
        use_ssl = self.var_use_ssl.get() if hasattr(self, 'var_use_ssl') else DEFAULT_USE_SSL
        use_starttls = self.var_use_starttls.get() if hasattr(self, 'var_use_starttls') else DEFAULT_USE_STARTTLS
        user = (self.var_smtp_user.get() if hasattr(self, 'var_smtp_user') else DEFAULT_SMTP_USER).strip()
        pwd = self.var_smtp_pass.get() if hasattr(self, 'var_smtp_pass') else DEFAULT_SMTP_PASS
        # Asunto desde editor enriquecido
        # Colapsar nuevas líneas para el header Subject (no admite saltos)
        asunto_plain = ' '.join(self.txt_asunto.get('1.0','end-1c').split())
        mensaje = self.txt_mensaje.get('1.0', tk.END).strip()
        if not asunto_plain or not mensaje:
            messagebox.showerror('Error', 'Completa todos los campos.')
            return
        if not messagebox.askyesno('Confirmar', f'Se enviará a {len(emails)} destinatarios. ¿Continuar?'):
            return
        self.btn_enviar.config(state=tk.DISABLED)
        def set_estado(msg):
            self.lbl_estado.config(text=msg)
        def run():
            # Generar HTML a partir del contenido con placeholders y plantilla
            html_body = self._editor_content_to_html(preview_local=False)
            for idx, _ in enumerate(self.image_paths, start=1):
                html_body = html_body.replace(f"{{{{image{idx}}}}}", f"<img src='cid:image{idx}' style='max-width:100%; height:auto; border-radius:8px' alt='img {idx}'>")
            html_body = re.sub(r'\{\{image\d+\}\}', '', html_body)
            subject_html = self._apply_inline_tags_widget(self.txt_asunto) or htmllib.escape(asunto_plain)
            html = self._apply_template(f"<h2 style='margin:0 0 12px'>{subject_html}</h2><div style='margin-bottom:12px'>{html_body}</div>")
            enviados, fallidos = enviar_mails(
                smtp_user=user,
                smtp_pass=pwd,
                emails=emails,
                asunto=asunto_plain,
                mensaje_texto=mensaje,
                status_callback=set_estado,
                smtp_server=server_host,
                smtp_port=server_port,
                use_ssl=use_ssl,
                use_starttls=use_starttls,
                html_body=html,
                image_paths=self.image_paths,
            )
            set_estado(f'Envío finalizado. Exitosos: {enviados} | Fallidos: {fallidos}')
            self.btn_enviar.config(state=tk.NORMAL)
        threading.Thread(target=run, daemon=True).start()

if __name__ == '__main__':
    app = MassiveMailSender()
    app.mainloop()
