import  win32com.client
import  tkinter as tk
import  win32timezone
from    tkinter import filedialog
from    tkinter import ttk, messagebox
from    datetime import time, datetime, timedelta, date
import  random
import  webbrowser
import  calendar as cal
from    dateutil.relativedelta import relativedelta
from    getpass import getuser
from    urllib.parse import quote
import  socket

from    reportlab.lib.pagesizes import A4, landscape
from    reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from    reportlab.lib.colors import HexColor
from    reportlab.lib.units import mm
from    reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, Paragraph, BaseDocTemplate, PageTemplate, Frame

import reportlab.platypus as rp

from    reportlab.platypus import Image as RLImage
from    reportlab.lib.enums import TA_CENTER, TA_LEFT
import  tempfile
import  os
import  sys

import io

import base64
from io import BytesIO
from PIL import Image, ImageTk

LOGO_BLACK_BASE64 = """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"""

bool_agenda = 1

def get_server_ip(hostname='UO061M4078592'):
    try:
        return socket.gethostbyname(hostname)
    except:
        return None

constante = '10.165.212.201'

ip = get_server_ip('UO061M4078592') or '10.165.212.201'
ip_zebra = ip

if ip != constante:
    str_alternativa = f'[Conectado ao IP alternativo: {ip}]'
else:
    str_alternativa = ''

print(f"Conectando ao servidor: {ip}")

db = "IST-PGE"

STR_CONN = ( f"Provider=SQLOLEDB;" f"Data Source={ip};" f"Initial Catalog={db};" f"User ID=sa;" f"Password=Wheelp0p3;" )

if bool_agenda == 1:
    APP_NAME = "Assistente de Agenda"
    APP_VERSION = 2
else:
    APP_NAME = "Assistente de Dados"
    APP_VERSION = 2

INFO_VERSAO =          "22/05/2026"

STR_CONN_LINKED = ( f"Provider=SQLOLEDB;" f"Data Source={ip_linked};" f"Initial Catalog={db_linked};" f"User ID={user_linked};" f"Password={password_linked};" )

if db_linked == 'MyLogical':
    instituto = 'Petróleo, Gás e Energia'
elif db_linked == 'Alimentos':
    instituto = 'Química & Meio Ambiente'
elif db_linked == 'Couro':
    instituto = 'Couro & Calçado'
elif db_linked == 'Madeira':
    instituto = 'Madeira e Mobiliário'
elif db_linked == 'Calcado':
    instituto = 'Couro & Calçado'

LABS = {

    'CD': {'name':  'Dimensional',                   'sector_id': 5,     'teams_chat_id': '461bebb602cc4c2f997670af64f1bc50',    'sharepoint_planilha': 'https://sesirs.sharepoint.com/:f:/r/sites/gdms-ISISistemasdeSensoriamento/Documentos%20Compartilhados/ISI%20SIM%20-%20Metrologia/Relat%C3%B3rios%20e%20Registros/Laborat%C3%B3rio%20de%20Dimensional',},
    'CDB': {'name': 'Metrologia por Coordenadas',    'sector_id': 11,    'teams_chat_id': '716ff922fa584a2582ecb48e509edc2e',    'sharepoint_planilha': '',},

    'CE': {'name':  'Eletricidade',                  'sector_id': 2029,  'teams_chat_id': '680cba2db76043939995da4b21cfd11c',    'sharepoint_planilha': 'https://sesirs.sharepoint.com/:f:/r/sites/gdms-ISISistemasdeSensoriamento/Documentos%20Compartilhados/ISI%20SIM%20-%20Metrologia/Relat%C3%B3rios%20e%20Registros/Lab.%20Eletricidade,%20Tempo%20e%20Frequ%C3%AAncia',},
    'TF': {'name':  'Tempo e Frequência',            'sector_id': 2030,  'teams_chat_id': '680cba2db76043939995da4b21cfd11c',    'sharepoint_planilha': 'https://sesirs.sharepoint.com/:f:/r/sites/gdms-ISISistemasdeSensoriamento/Documentos%20Compartilhados/ISI%20SIM%20-%20Metrologia/Relat%C3%B3rios%20e%20Registros/Lab.%20Eletricidade,%20Tempo%20e%20Frequ%C3%AAncia',},

    'CM': {'name':  'Massa',                         'sector_id': 6,     'teams_chat_id': '474353252d224d7caf749a4b5301c4d8',    'sharepoint_planilha': '',},
    'CP': {'name':  'Pressão',                       'sector_id': 7,     'teams_chat_id': 'fc9176a569904180bbfa3eeb1bd52651',    'sharepoint_planilha': 'https://sesirs.sharepoint.com/:f:/r/sites/gdms-ISISistemasdeSensoriamento/Documentos%20Compartilhados/ISI%20SIM%20-%20Metrologia/Relat%C3%B3rios%20e%20Registros/Laborat%C3%B3rio%20de%20Press%C3%A3o',},
    'CT': {'name':  'Temperatura e Umidade',         'sector_id': 25,    'teams_chat_id': '680cba2db76043939995da4b21cfd11c',    'sharepoint_planilha': 'https://sesirs.sharepoint.com/:f:/r/sites/gdms-ISISistemasdeSensoriamento/Documentos%20Compartilhados/ISI%20SIM%20-%20Metrologia/Relat%C3%B3rios%20e%20Registros/Laborat%C3%B3rio%20de%20Temperatura%20e%20Umidade',},
    'CF': {'name':  'Força, Torque e Dureza',        'sector_id': 8,     'teams_chat_id': 'fc9176a569904180bbfa3eeb1bd52651',    'sharepoint_planilha': 'https://sesirs.sharepoint.com/:f:/r/sites/gdms-ISISistemasdeSensoriamento/Documentos%20Compartilhados/ISI%20SIM%20-%20Metrologia/Relat%C3%B3rios%20e%20Registros/Laborat%C3%B3rio%20Ensaios,%20For%C3%A7a%20e%20Torque',},

    'VZ': {'name':  'Vazão',                         'sector_id': 1026,  'teams_chat_id': 'e04abb13e3114d95b399290fe371a391',    'sharepoint_planilha': 'https://sesirs.sharepoint.com/:f:/r/sites/gdms-ISISistemasdeSensoriamento/Documentos%20Compartilhados/ISI%20SIM%20-%20Metrologia/Relat%C3%B3rios%20e%20Registros/Laborat%C3%B3rio%20de%20Vaz%C3%A3o',},
    'CV': {'name':  'Volume e Massa Específica',     'sector_id': 9,     'teams_chat_id': '680cba2db76043939995da4b21cfd11c',    'sharepoint_planilha': 'https://sesirs.sharepoint.com/:f:/r/sites/gdms-ISISistemasdeSensoriamento/Documentos%20Compartilhados/ISI%20SIM%20-%20Metrologia/Relat%C3%B3rios%20e%20Registros/Lab.%20Eletricidade,%20Tempo%20e%20Frequ%C3%AAncia',},
    
}

username00 = getuser()
username0 = username00.split('.')[0]
username = username0.capitalize()

greetings = [
    f' Oi olá ',
    f' Vai almoçar na Bete hoje? ',  
    f' E o churras?? ',
    f' Ora ora ',
    f' Buenas, {username}? ',
    f' {username}?! ',
    f' "Nani??" ',
    f' Tudo bom, {username}? ',
    f' Hi there, {username}! ',
    f' Hello there, {username}! ',
    f' Hellooo! ',
    f' Fala aí {username}, tudo bom? ',
    f' Tudo bão {username}? ',
    f' Bão!? ',
    f' E aí {username}, bão!? ',
    f' {username}!? ',
    f' Alô alô {username}! ',
    f' E aí {username}, tudo certo? ',
    f' Seu nome é Gabriel? ',
    f' Opa opa ',    
    f' Aoba ',    
    f' Tudo certo? ',
    f' ¯|_(ツ)_|¯ ',
    f' *_* ',
    f' Buenas tardes! ', 
    f' Vai um chimas? '
]

buenas = random.choice(greetings)

def check_app_version(str_conn, app_name, current_version):
    """Check if this app's version matches the latest in database"""
    try:
        conn = win32com.client.Dispatch("ADODB.Connection")
        rs = win32com.client.Dispatch("ADODB.Recordset")
        print('str_conn',str_conn)
        conn.Open(str_conn)
        
        rs.Open(f"""
            SELECT version FROM [IST-PGE].dbo.App_Version 
            WHERE app_name = '{app_name}'
            ORDER BY version DESC
        """, conn)
        
        db_version = rs.Fields('version').Value if not rs.EOF else 0
        
        rs.Close()
        conn.Close()
        
        if db_version > current_version:
            return {
                'update_required': True,
                'current': current_version,
                'latest': db_version,
                'message': f"Oii {username}, tudo bommmm?!\n\nSua versão do software {APP_NAME} está obsoleta!\n\nEntre em contato com o Castro para atualizar."

            }
        else:
            return {
                'update_required': False,
                'current': current_version,
                'latest': db_version,
                'message': f"{app_name} - Versão {current_version}"
            }
            
    except Exception as e:
        return {
            'update_required': False,
            'message': f"Não foi possível verificar a versão."
        }

def _from_rgb(rgb): return "#%02x%02x%02x" % rgb

#ef verificar_disponibilidade():
#   try:
#       conn = win32com.client.Dispatch("ADODB.Connection")
#       #conn.ConnectionTimeout = 5
#       print(STR_CONN)
#       conn.Open(STR_CONN)
#       conn.Close()
#       return "ONLINE", "lime"
#   except Exception:
#       return "OFFLINE", "orange"

def verificar_disponibilidade3():
    try:
        conn = win32com.client.Dispatch("ADODB.Connection")
        conn.ConnectionTimeout = 6
        conn.Open(STR_CONN_LINKED)
        conn.Close()
        return "ONLINE", "lime"
    except Exception:
        return "OFFLINE", "orange"

def format_os_code(os_code):
    if len(os_code) == 1:
        return f"000{os_code}"
    elif len(os_code) == 2:
        return f"00{os_code}"
    elif len(os_code) == 3:
        return f"0{os_code}"
    return os_code

# Global variable to track the loading popup
_loading_popup = None


def show_loading_screen(parent_frame, message="Carregando..."):
    """Show a loading overlay - can be called from anywhere"""
    global _loading_popup
    
    # Hide any existing loading screen first
    #hide_loading_screen()
    
    _loading_popup = tk.Toplevel(parent_frame)
    _loading_popup.geometry("550x130")
    
    cor_fundo2 = _from_rgb((0, 0, 0))
    _loading_popup.configure(bg=cor_fundo2)
    _loading_popup.title("")
    _loading_popup.overrideredirect(True)
    _loading_popup.attributes('-topmost', True)
    
    # Center on parent
    _loading_popup.update_idletasks()
    x = parent_frame.winfo_rootx() + (parent_frame.winfo_width() // 2) - 275
    y = parent_frame.winfo_rooty() + (parent_frame.winfo_height() // 2) - 65
    _loading_popup.geometry(f"+{x}+{y}")
    
    frame = tk.Frame(_loading_popup, bg=cor_fundo2, highlightbackground="#FFFFFF", highlightthickness=2)
    frame.pack(fill='both', expand=True, padx=2, pady=2)


    tk.Label(frame, text=f"Oioi, {username}!",font=('Segoe UI', 10), background='black', fg="#FFFFFF").pack(pady=(15, 5))
    
    tk.Label(frame, text=message, font=('Segoe UI', 11),bg='white', fg='black').pack(pady=(5, 5))
    
    _loading_popup.grab_set()
    _loading_popup.update()


def hide_loading_screen():
    """Hide the loading overlay - can be called from anywhere"""
    global _loading_popup
    
    if _loading_popup is not None and _loading_popup.winfo_exists():
        _loading_popup.grab_release()
        _loading_popup.destroy()
        _loading_popup = None

class DatabaseViewer2: ### PADRÕES

    def __init__(self, parent_notebook, str_conn, cor_fundo):
        self.str_conn = str_conn
        self.cor_fundo = cor_fundo
        self.selected_items = []
        self.frame_certificados = ttk.Frame(parent_notebook)
        parent_notebook.add(self.frame_certificados, text=" Padrões ")
        
        self.setup_ui_certificados()
    
    def on_entry_focus_in(self, event):
        if self.search_entry.get() == self.placeholder_text:
            self.search_entry.delete(0, 'end')
            self.search_entry.configure(style='TEntry')

    def on_entry_focus_out(self, event):
        if not self.search_entry.get():
            self.search_entry.insert(0, self.placeholder_text)
            style = ttk.Style()
            self.search_entry.configure(style='Placeholder.TEntry')
        
    def setup_ui_certificados(self):
        search_frame = ttk.Frame(self.frame_certificados)
        search_frame.pack(pady=10, padx=10, fill="x")
        
        self.greetings = ttk.Label(search_frame, text=f"{buenas}", font=("Segoe UI", 10),background='black')
        self.greetings.pack(side="left", padx=5)

        #ttk.Label(search_frame, text="Buscar:").pack(side="left", padx=(0, 15))
        self.search_entry = ttk.Entry(search_frame, width=30)
        self.search_entry.pack(side="left", padx=5)
        self.search_entry.bind('<Return>', lambda e: self.perform_search())
        
        style = ttk.Style()
        style.configure('Placeholder.TEntry', foreground='#7D7D7D')

        self.placeholder_text = "Digite sua busca..."
        self.search_entry.insert(0, self.placeholder_text)
        self.search_entry.configure(style='Placeholder.TEntry')

        self.search_entry.bind('<FocusIn>', self.on_entry_focus_in)
        self.search_entry.bind('<FocusOut>', self.on_entry_focus_out)
        
        ttk.Button(search_frame, text="Buscar", command=self.perform_search).pack(side="left", padx=5)
        ttk.Button(search_frame, text="Mostrar Todos", command=self.load_all_records).pack(side="left", padx=5)
        
        tree_frame = ttk.Frame(self.frame_certificados)
        tree_frame.pack(pady=10, padx=10, expand=True, fill="both")
        
        tree_scroll_y = ttk.Scrollbar(tree_frame)
        tree_scroll_y.pack(side="right", fill="y")
        
        tree_scroll_x = ttk.Scrollbar(tree_frame, orient="horizontal")
        tree_scroll_x.pack(side="bottom", fill="x")
        
        columns = (
                    "Padrão", 
                    "Instrumento", 
                    "Declaração de rastreabilidade metrológica",
                    "Status", 
                    #"Próxima Calibração", 
                    "Setor"
                    #"Ativo/Inativo",
                    #"Certificado",
                    #"Fornecedor")
        )
        
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set, selectmode="extended")

        #self.tree.tag_configure('red_bg', background='red')
        
        tree_scroll_y.config(command=self.tree.yview)
        tree_scroll_x.config(command=self.tree.xview)

        column_widths = {  

                            "Padrão": 5,
                            "Instrumento": 40,
                            "Declaração de rastreabilidade metrológica": 250,
                            "Status": 20,
                            "Setor": 20,
                            "Próxima Calibração": 20,
                            "Ativo/Inativo": 20,
                            "Certificado": 20,
                            "Fornecedor": 20

                        }

        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=column_widths.get(col, 100), minwidth=50)
        
        self.tree.pack(expand=True, fill="both")
        
        #self.tree.tag_configure('red_bg', background='red')
        
        self.tree.bind('<Button-3>', self.show_context_menu)
        
        self.tree.bind('<<TreeviewSelect>>', self.on_select)
        
        button_frame = ttk.Frame(self.frame_certificados)
        button_frame.pack(pady=10, padx=10, fill="x")
        self.btn_visualizar = ttk.Button(button_frame, text="Visualizar Selecionado(s)", command=self.print_selected)
        self.btn_visualizar.pack(side="left", padx=5)
        self.btn_visualizar.config(state="disabled")
        
        ##self.btn_gerar = ttk.Button(button_frame, text="Função 2", command=self.export_selected)
        ##self.btn_gerar.pack(side="left", padx=5)
        ##self.btn_gerar.config(state="disabled")

        self.btn_limpar = ttk.Button(button_frame, text="Limpar Seleção", command=self.clear_selection)
        self.btn_limpar.pack(side="left", padx=5)
        self.btn_limpar.config(state="disabled")

        self.status_label = ttk.Label(self.frame_certificados, text="Pronto para consulta", font=("Segoe UI", 9))
        self.status_label.pack(pady=5)
        
        self.load_all_records()
        
    def show_context_menu(self, event):
        item = self.tree.identify_row(event.y)
        column = self.tree.identify_column(event.x)
        
        if item:
            self.tree.selection_set(item)
            
            context_menu = tk.Menu(self.tree, tearoff=0)
            
            values = self.tree.item(item)['values']
            line = values[2] if values and len(values) > 0 else ""

            if values[2]:
                #context_menu.add_command(label="Abrir link do Sharepoint no navegador", command=lambda: self.open_sharepoint_link(line))
                
                if values[3] == "":
                    context_menu.add_command(label="Copiar declaração de rastreabilidade metrológica", command=lambda: self.copy_to_clipboard(line))
                    #context_menu.add_separator()
                    #context_menu.add_command(label="Copiar linha", command=lambda: self.copy_row_data(item))

            #context_menu.add_separator()
            
            context_menu.post(event.x_root, event.y_root)

    def open_sharepoint_link(self, url):
        try:
            webbrowser.open(url)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao abrir o link:\n{str(e)}")
            self.status_label.config(text="Falha ao abrir link")

    def copy_to_clipboard(self, text):
        
        try:
            self.frame_certificados.clipboard_clear()
            self.frame_certificados.clipboard_append(text)
            self.status_label.config(text="Declaração de rastreabilidade metrológica copiada para a área de transferência!")
            messagebox.showinfo("Sucesso", "Declaração de rastreabilidade metrológica copiada para a área de transferência!")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao copiar:\n{str(e)}")

    #def copy_row_data(self, item):
    #    
    #    try:
    #        values = self.tree.item(item)['values']
    #        if values:
    #            
    #            os_year = values[1][-2:]
    #            
    #            os_parts = values[1].split('/')
    #            
    #            os_code = os_parts[0].lstrip('0')
    #            if len(os_code) == 4:                
    #                values[1] = f"{os_code}/{os_year}"
    #            else:
    #                values[1] = f"0{os_code}/{os_year}"
    #            print(values[1])
#
    #            row_text = "\n".join([f"{val}" for val in values])
    #            
    #            self.frame_certificados.clipboard_clear()
    #            self.frame_certificados.clipboard_append(row_text)
    #            self.status_label.config(text="Linha copiada para a área de transferência!")
    #            messagebox.showinfo("Sucesso", "Dados da linha copiados para a área de transferência!")
    #    except Exception as e:
    #        messagebox.showerror("Erro", f"Falha ao copiar linha:\n{str(e)}")

    def select_all_rows(self):
        
        all_items = self.tree.get_children()
        for item in all_items:
            self.tree.selection_add(item)
        self.status_label.config(text=f"{len(all_items)} linhas selecionadas")
        
        ##self.btn_gerar.config(state="enabled")
        self.btn_visualizar.config(state="enabled")
        self.btn_limpar.config(state="enabled")
    
    def execute_query(self, where_clause=None):
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            rs = win32com.client.Dispatch("ADODB.Recordset")
            
            conn.Open(self.str_conn)
            
            sql = """select
            
            s.code as 'standard',
            it.name as 's_instrument_type',
            --concat(it.name,' ',im.name,' – ',s.code,' – Certificado de Calibração ',last_cal.certificate_code,' emitido por ',sup_cal.name,' – Válido até ',RIGHT(CONVERT(VARCHAR(10), s.next_calibration, 103), 7)) as 'traceability',
            concat(it.name,' – ',s.code,' – certificado de calibração ',last_cal.certificate_code,' emitido por ',sup_cal.name,', válido até ',RIGHT(CONVERT(VARCHAR(10), s.next_calibration, 103), 7)) as 'traceability',
            case when s.next_calibration is null or sec.name is null or last_cal.certificate_code is null or sup_cal.name is null or s.next_calibration < GETDATE() or s.is_active = 0 then 'Fora de uso' end as 'status',
            --SE O PADRÃO TÁ VENCIDO
            --SE O PADRÃO ESTIVER INATIVO
            --SE O PADRÃO NÃO TEM DATA DA PRÓXIMA CALIBRAÇÃO
            --SE O PADRÃO NÃO TEM SETOR RESPONSÁVEL
            --SE O PADRÃO NÃO TEM CERTIFICADO DE CALIBRAÇÃO NENHUM
            -- >>>>FORA DE USO
            CONVERT(VARCHAR(10), s.next_calibration, 103) AS 's_next_calibration',
            sec.name as 's_sector',
            case when s.is_active = 1 then 'Ativo' else 'Inativo' end as 'Status',
            last_cal.certificate_code as 'sc_certificate_code',
            sup_cal.name as 'sc_supplier'

            from standards as s
            left join instrument_types it on it.id = s.id_instrument_type
            left join instruments_models im on im.id = s.id_model
            left join suppliers sup on sup.id = s.id_supplier_calibration
            left join sectors sec on sec.id = s.id_current_sector

            left join (select sc.id_standard,max(sc.id) as max_sc_id from standards_calibrations sc where sc.removed=0 group by sc.id_standard) latest_cal_id on latest_cal_id.id_standard = s.id
            left join standards_calibrations last_cal on last_cal.id = latest_cal_id.max_sc_id
            left join suppliers sup_cal on sup_cal.id = last_cal.id_supplier
            """
            
            if where_clause:
                sql += f" WHERE {where_clause}"
            
            sql += " order by s.code desc"
            
            rs.Open(sql, conn)
            
            results = []
            if not rs.EOF:
                rs.MoveFirst()
                while not rs.EOF:
                    row = []
                    for i in range(rs.Fields.Count):
                        value = rs.Fields(i).Value
                        row.append(value if value is not None else "")
                    #os_year = row[1][-4:]
                    #os_code = row[1].split('/')
                    #os_code = os_code[0].lstrip('0')
                    #os_code = format_os_code(os_code)
                    #row[0] = f"https://sesirs.sharepoint.com/:f:/r/sites/gdms-ISISistemasdeSensoriamento/Documentos%20Compartilhados/ISI%20SIM%20-%20Metrologia/Atendimento%20ao%20Cliente/E%20-%20Ordens%20de%20Servi%C3%A7o/{row[0]}{os_year}/{os_code}"
                    
                    results.append(row)
                    rs.MoveNext()
            
            rs.Close()
            conn.Close()
            
            return results
            
        except Exception as e:
            messagebox.showerror("Erro", f"Falha na consulta:\n{str(e)}")
            return []
        
    def load_all_records(self):
        self.status_label.config(text="Carregando certificados...", font=("Segoe UI", 12))
        self.frame_certificados.update()
        
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        where_clause = "s.removed = 0"
        
        results = self.execute_query(where_clause)
        
        for row in results:
            if row[3] == "Fora de uso":
                #self.tree.insert("", "end", values=row, tags=('red_bg',))
                
                item_id = self.tree.insert("", "end", values=row)
                self.tree.item(item_id, tags=('fora_de_uso',))
                
                self.tree.tag_configure('fora_de_uso', background='#FFD9D9')
            else:
                self.tree.insert("", "end", values=row)
        
        self.status_label.config(text=f"{len(results)} padrões encontrados")

    def perform_search(self):
        
        search_term = self.search_entry.get().strip()
        
        if search_term == "Digite sua busca...":
            return

        if not search_term:
            self.load_all_records()
            return
        
        self.status_label.config(text=f"Buscando por '{search_term}'...")
        self.frame_certificados.update()
        
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        where_clause = f"(it.name LIKE '%{search_term}%' or s.code LIKE '%{search_term}%' or sup.name LIKE '%{search_term}%' or sec.name LIKE '%{search_term}%') and s.removed = 0 "
        #where_clause = f"(it.name LIKE '%{search_term}%' or s.code LIKE '%{search_term}%' and s.removed = 0 and it.removed = 0 and im.removed = 0 and sup.removed = 0 and sec.removed = 0)"
        
        results = self.execute_query(where_clause)
        
        for row in results:
            if row[3] == "Fora de uso":
                item_id = self.tree.insert("", "end", values=row)
                self.tree.item(item_id, tags=('fora_de_uso',))
                
                self.tree.tag_configure('fora_de_uso', background='#FFD9D9')
            else:
                self.tree.insert("", "end", values=row)

        if len(results) == 0:
            self.status_label.config(text=f"Nenhum padrão encontrado para '{search_term}'")
        if len(results) == 1:
            self.status_label.config(text=f"{len(results)} padrões encontrados para '{search_term}'")
        if len(results) > 1:
            self.status_label.config(text=f"{len(results)} padrões para '{search_term}'")
    
    def on_select(self, event):
        selected = self.tree.selection()
        if len(selected) > 1:
            self.status_label.config(text=f"{len(selected)} padrões selecionados")
            #self.btn_gerar.config(state="enabled")
            self.btn_visualizar.config(state="enabled")
            self.btn_limpar.config(state="enabled")

        if len(selected) == 1:
            self.status_label.config(text=f"{len(selected)} padrão selecionado")
            self.btn_visualizar.config(state="enabled")
            self.btn_limpar.config(state="enabled")

        if len(selected) == 0:
            self.status_label.config(text=f"Nenhum padrão selecionado")
            self.btn_visualizar.config(state="disabled")
            self.btn_limpar.config(state="disabled")

    def get_selected_data(self):
        selected_items = self.tree.selection()
        selected_data = []
        
        for item in selected_items:
            values = self.tree.item(item)['values']
            selected_data.append({
                'standard':             values[0],
                'instrument_type':      values[1],
                'traceability':         values[2],
                'status':               values[3],
                's_next_calibration':   values[4],
                's_sector':             values[5],
                'Status':               values[6],
                'sc_certificate_code':  values[7],
                'sc_supplier':          values[8]
            })

        return selected_data
    
    def print_selected(self):
        selected_data = self.get_selected_data()
        
        if not selected_data:
            messagebox.showinfo("Aviso", "Nenhum item selecionado para visualizar.")
            return
        
        print_window = tk.Toplevel(self.frame_certificados)
        print_window.title("Padrões selecionados")
        print_window.geometry("900x500")
        print_window.configure(bg=self.cor_fundo)
        
        #title_label = tk.Label(print_window, text="VISUALIZAÇÃO DE DADOS", font=("Segoe UI", 12, "bold"), bg=self.cor_fundo, fg="white")
        #title_label.pack(pady=35)
        
        text_frame = ttk.Frame(print_window)
        text_frame.pack(pady=10, padx=10, expand=True, fill="both")
        
        text_widget = tk.Text(text_frame, wrap="word", font=("Consolas", 10))
        scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side="right", fill="y")
        text_widget.pack(side="left", expand=True, fill="both")
        
        output_text = f"{'='*70}\n"
        output_text += f"Total de itens selecionados: {len(selected_data)}\n"
        output_text += f"{'='*70}\n\n"
        
        for i, item in enumerate(selected_data, 1):
            output_text += f"Padrão {item['standard']}:\n"
             
            output_text += f"{item['traceability']}\n"
            
            #output_text += f"Instrumento: {item['instrument_type']}\n"
            output_text += f"\n\n"
            if item['status'] == "":
                output_text += f">>> Status da Declaração de rastreabilidade metrológica: OK\n"
            else:
                output_text += f">>> Status da Declaração de Rastreabilidade metrológica: {item['status']}\n"
            
            if item['s_sector'] == "":
                output_text += f">>> Não há setor responsável cadastrado neste padrão\n"
            else:
                output_text += f">>> Setor responsável: {item['s_sector']}\n"
                
            if item['Status'] == "Inativo":
                output_text += f">>> Padrão inativo\n"

            if item['sc_certificate_code'] == "":
                output_text += f">>> Não há certificado cadastrado neste padrão\n"
            if item['sc_supplier'] == "":
                output_text += f">>> Não há fornecedor cadastrado no certificado deste padrão\n"

            output_text += f"{'-'*40}\n\n"

        text_widget.insert("1.0", output_text)
        text_widget.config(state="disabled")
        
    def export_selected(self):
        selected_data = self.get_selected_data()
        
        if not selected_data:
            messagebox.showinfo("Aviso", "Nenhum item selecionado para exportar.")
            return
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"certificados SELECIONADOS{timestamp}.txt"
        
        try:
            with open(filename, "w", encoding="utf-8") as f:
                f.write("="*70 + "\n")
                f.write("EXPORTAÇÃO DE DADOS - SISTEMA DE CALIBRAÇÕES\n")
                f.write(f"Data/Hora da exportação: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
                f.write(f"Total de certificados exportados: {len(selected_data)}\n")
                f.write("="*70 + "\n\n")
                
                for i, item in enumerate(selected_data, 1):
                    f.write(f"REGISTRO {i}:\n")
                    f.write(f"  Ordem de Serviço: {item['ordem_servico']}\n")
                    f.write(f"  Certificado: {item['certificado']}\n")
                    f.write(f"  Item: {item['item']}\n")
                    f.write(f"  Data da Calibração: {item['data_calibracao']}\n")
                    f.write(f"  Data da Calibração: {item['cliente']}\n")
                    f.write("-"*40 + "\n\n")
            
            messagebox.showinfo("Sucesso", f"Dados exportados com sucesso para:\n{filename}")
            self.status_label.config(text=f"Exportado para {filename}")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao exportar dados:\n{str(e)}")
    
    def clear_selection(self):
        for item in self.tree.selection():
            self.tree.selection_remove(item)
        self.status_label.config(text="Seleção limpa")


def insert_protocol_and_get_id(str_conn, user_name, client_name='', budget_code='', service_order_code='', pdf_filename=''):
    """
    Insert a new protocol record and return the generated ID as the protocol code
    Returns: dict with 'id', 'protocol_code', or None if failed
    """
    try:
        conn = win32com.client.Dispatch("ADODB.Connection")
        rs = win32com.client.Dispatch("ADODB.Recordset")
        
        conn.Open(str_conn)
        print('conn.Open(str_conn)', str_conn)
        
        # Escape single quotes in strings
        safe_client = client_name.replace("'", "''") if client_name else ''
        safe_budget = budget_code.replace("'", "''") if budget_code else ''
        safe_order = service_order_code.replace("'", "''") if service_order_code else ''
        safe_filename = pdf_filename.replace("'", "''") if pdf_filename else ''
        # Insert the record with a temporary protocol code

        insert_sql = f"""
            INSERT INTO Receipt_Protocols 
            (protocol_code, insert_date, user_name, client_name, budget_code, service_order_code, pdf_filename)
            VALUES (
                'TEMP',
                GETDATE(),
                '{user_name}',
                '{safe_client}',
                '{safe_budget}',
                '{safe_order}',
                '{safe_filename}'
            )
        """

        conn.Execute(insert_sql)

        # Now get the last inserted ID
        rs.Open("SELECT SCOPE_IDENTITY() as new_id", conn)

        new_id = None
        if not rs.EOF:
            new_id = rs.Fields('new_id').Value
        rs.Close()

        if new_id is None:
            conn.Close()
            print("Erro: Não foi possível obter o ID do registro inserido")
            return None

        # Format protocol code using the ID with leading zeros
        protocol_code = f"{int(new_id)}"

        # Update the record with the proper protocol code
        update_sql = f"""
            UPDATE Receipt_Protocols 
            SET protocol_code = '{protocol_code}'
            WHERE id = {int(new_id)}
        """
        conn.Execute(update_sql)
        
        
        # Now get the last inserted ID
        rs.Open("SELECT SCOPE_IDENTITY() as new_id", conn)
        
        new_id = None
        if not rs.EOF:
            new_id = rs.Fields('new_id').Value
        rs.Close()
        
        if new_id is None:
            conn.Close()
            print("Erro: Não foi possível obter o ID do registro inserido")
            return None
        
        # Format protocol code using the ID with leading zeros
        protocol_code = f"{int(new_id)}"
        
        # Update the record with the formatted protocol code
        update_sql = f"""
            UPDATE Receipt_Protocols 
            SET protocol_code = '{protocol_code}'
            WHERE id = {int(new_id)}
        """
        conn.Execute(update_sql)
        
        conn.Close()
        
        print(f"Protocolo inserido: ID={new_id}, Código={protocol_code}")
        
        return {
            'id': int(new_id),
            'protocol_code': protocol_code
        }
        
    except Exception as e:
        print(f"Erro ao inserir protocolo: {str(e)}")
        messagebox.showerror("Erro", f"Falha ao registrar protocolo:\n{str(e)}")
        try:
            conn.Close()
        except:
            pass
        return None

class ProtocolViewer: ### VISUALIZADOR DE PROTOCOLOS
    
    def __init__(self, parent_notebook, str_conn, cor_fundo):
        self.str_conn = str_conn
        self.cor_fundo = cor_fundo
        self.frame_protocolos = ttk.Frame(parent_notebook)
        parent_notebook.add(self.frame_protocolos, text=" Lista de Recebimento ")
        
        self.setup_ui()
    
    def setup_ui(self):
        # Search frame
        search_frame = ttk.Frame(self.frame_protocolos)
        search_frame.pack(pady=10, padx=10, fill="x")
        
        ttk.Label(search_frame, text="Buscar DRs:").pack(side="left", padx=5)
        
        self.search_entry = ttk.Entry(search_frame, width=30)
        self.search_entry.pack(side="left", padx=5)
        self.search_entry.bind('<Return>', lambda e: self.perform_search())
        
        # Placeholder
        self.placeholder_text = "Buscar por código, cliente, OS..."
        self.search_entry.insert(0, self.placeholder_text)
        style = ttk.Style()
        style.configure('Placeholder.TEntry', foreground='#7D7D7D')
        self.search_entry.configure(style='Placeholder.TEntry')
        self.search_entry.bind('<FocusIn>', self.on_entry_focus_in)
        self.search_entry.bind('<FocusOut>', self.on_entry_focus_out)
        
        ttk.Button(search_frame, text="Buscar", command=self.perform_search).pack(side="left", padx=5)
        ttk.Button(search_frame, text="Mostrar Todos", command=self.load_all_records).pack(side="left", padx=5)
        ttk.Button(search_frame, text="Atualizar", command=self.load_all_records).pack(side="left", padx=5)
        
        # Date filter
        #ttk.Label(search_frame, text="De:").pack(side="left", padx=(20, 2))
        #self.date_from = ttk.Entry(search_frame, width=12)
        #self.date_from.pack(side="left", padx=2)
        #self.date_from.insert(0, (datetime.now() - timedelta(days=30)).strftime('%d/%m/%Y'))
        
        #ttk.Label(search_frame, text="Até:").pack(side="left", padx=2)
        #self.date_to = ttk.Entry(search_frame, width=12)
        #self.date_to.pack(side="left", padx=2)
        #self.date_to.insert(0, datetime.now().strftime('%d/%m/%Y'))
        
        #ttk.Button(search_frame, text="Filtrar por Data", command=self.filter_by_date).pack(side="left", padx=5)
        
        # Tree frame
        tree_frame = ttk.Frame(self.frame_protocolos)
        tree_frame.pack(pady=10, padx=10, expand=True, fill="both")
        
        tree_scroll_y = ttk.Scrollbar(tree_frame)
        tree_scroll_y.pack(side="right", fill="y")
        
        tree_scroll_x = ttk.Scrollbar(tree_frame, orient="horizontal")
        tree_scroll_x.pack(side="bottom", fill="x")
        
        columns = (
            "Número",
            "Data de Inserção",
            "Usuário",
            "Cliente",
            "Orçamento",
            "Ordem de Serviço",
            "Arquivo PDF"
        )
        
        self.tree = ttk.Treeview(
            tree_frame, 
            columns=columns, 
            show="headings", 
            yscrollcommand=tree_scroll_y.set, 
            xscrollcommand=tree_scroll_x.set, 
            selectmode="extended"
        )
        
        tree_scroll_y.config(command=self.tree.yview)
        tree_scroll_x.config(command=self.tree.xview)
        
        column_widths = {
            "Número": 150,
            "Data de Inserção": 130,
            "Usuário": 100,
            "Cliente": 200,
            "Orçamento": 100,
            "Ordem de Serviço": 100,
            "Arquivo PDF": 150
        }
        
        for col in columns:
            self.tree.heading(col, text=col, command=lambda c=col: self.sort_column(c))
            self.tree.column(col, width=column_widths.get(col, 100), minwidth=50)
        
        self.tree.pack(expand=True, fill="both")
        
        # Right-click menu
        self.tree.bind('<Button-3>', self.show_context_menu)
        
        # Status bar
        self.status_label = ttk.Label(
            self.frame_protocolos, 
            text="Pronto para consulta", 
            font=("Segoe UI", 9)
        )
        self.status_label.pack(pady=5)
        
        # Load all records
        self.load_all_records()
    
    def on_entry_focus_in(self, event):
        if self.search_entry.get() == self.placeholder_text:
            self.search_entry.delete(0, 'end')
            self.search_entry.configure(style='TEntry')
    
    def on_entry_focus_out(self, event):
        if not self.search_entry.get():
            self.search_entry.insert(0, self.placeholder_text)
            self.search_entry.configure(style='Placeholder.TEntry')
    
    def execute_query(self, where_clause=None):
        """Execute query and return results"""
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            rs = win32com.client.Dispatch("ADODB.Recordset")
            
            conn.Open(self.str_conn)
            
            sql = """
                SELECT 
                
                    protocol_code,
                    insert_date,
                    user_name,
                    client_name,
                    budget_code,
                    service_order_code,
                    pdf_filename
                FROM Receipt_Protocols
                WHERE removed = 0
            """
            
            if where_clause:
                sql += f" AND {where_clause}"
            
            sql += " ORDER BY insert_date DESC"
            
            rs.Open(sql, conn)
            
            results = []
            if not rs.EOF:
                rs.MoveFirst()
                while not rs.EOF:
                    row = []
                    for i in range(rs.Fields.Count):
                        value = rs.Fields(i).Value
                        # Format date
                        if i == 1 and value:  # insert_date
                            try:
                                value = value.strftime('%d/%m/%Y %H:%M:%S')
                            except:
                                pass
                        row.append(value if value is not None else "")
                    results.append(row)
                    rs.MoveNext()
            
            rs.Close()
            conn.Close()
            
            return results
            
        except Exception as e:
            messagebox.showerror("Erro", f"Falha na consulta:\n{str(e)}")
            return []
    
    def load_all_records(self):
        """Load all protocol records"""
        self.status_label.config(text="Carregando protocolos...")
        self.frame_protocolos.update()
        
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        results = self.execute_query()
        
        for row in results:
            self.tree.insert("", "end", values=row)
        
        #self.status_label.config(text=f"{len(results)} protocolos encontrados")
    
    def perform_search(self):
        """Search protocols by various fields"""
        search_term = self.search_entry.get().strip()
        
        if search_term == self.placeholder_text or not search_term:
            self.load_all_records()
            return
        
        self.status_label.config(text=f"Buscando por '{search_term}'...")
        self.frame_protocolos.update()
        
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        where_clause = f"""
            (protocol_code LIKE '%{search_term}%' 
            OR client_name LIKE '%{search_term}%' 
            OR user_name LIKE '%{search_term}%'
            OR budget_code LIKE '%{search_term}%'
            OR service_order_code LIKE '%{search_term}%')
        """
        
        results = self.execute_query(where_clause)
        
        for row in results:
            self.tree.insert("", "end", values=row)
        
        if len(results) == 0:
            self.status_label.config(text=f"Nenhum protocolo encontrado para '{search_term}'")
        elif len(results) == 1:
            self.status_label.config(text=f"{len(results)} protocolo encontrado")
        else:
            self.status_label.config(text=f"{len(results)} protocolos encontrados")
    
    def filter_by_date(self):
        """Filter protocols by date range"""
        date_from_str = self.date_from.get()
        date_to_str = self.date_to.get()
        
        try:
            date_from = datetime.strptime(date_from_str, '%d/%m/%Y')
            date_to = datetime.strptime(date_to_str, '%d/%m/%Y')
            
            # Set time to end of day for date_to
            date_to = date_to.replace(hour=23, minute=59, second=59)
            
            self.status_label.config(text=f"Filtrando de {date_from_str} até {date_to_str}...")
            self.frame_protocolos.update()
            
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            where_clause = f"""
                insert_date >= '{date_from.strftime('%Y-%m-%d %H:%M:%S')}'
                AND insert_date <= '{date_to.strftime('%Y-%m-%d %H:%M:%S')}'
            """
            
            results = self.execute_query(where_clause)
            
            for row in results:
                self.tree.insert("", "end", values=row)
            
            self.status_label.config(text=f"{len(results)} protocolos no período")
            
        except ValueError:
            messagebox.showerror("Erro", "Formato de data inválido. Use DD/MM/AAAA")
    
    def show_context_menu(self, event):
        """Show right-click context menu"""
        item = self.tree.identify_row(event.y)
        
        if item:
            self.tree.selection_set(item)
            
            context_menu = tk.Menu(self.tree, tearoff=0)
            
            values = self.tree.item(item)['values']
            
            context_menu.add_command(
                label="Copiar Protocolo", 
                command=lambda: self.copy_to_clipboard(values[1])
            )
            
            if values[6]:  # If PDF filename exists
                context_menu.add_command(label="Copiar Nome do Arquivo", command=lambda: self.copy_to_clipboard(values[7]))
            
            context_menu.add_separator()
            context_menu.add_command(label="Copiar Todos os Dados", command=lambda: self.copy_all_data(item))
            
            context_menu.post(event.x_root, event.y_root)
    
    def copy_to_clipboard(self, text):
        """Copy text to clipboard"""
        try:
            self.frame_protocolos.clipboard_clear()
            self.frame_protocolos.clipboard_append(str(text))
            #self.status_label.config(text="Texto copiado para a área de transferência!")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao copiar:\n{str(e)}")
    
    def copy_all_data(self, item):
        """Copy all row data to clipboard"""
        try:
            values = self.tree.item(item)['values']
            columns = ["Número", "Data de Inserção", "Usuário", 
                "Cliente", "Orçamento", "Ordem de Serviço", "Arquivo PDF"
            ]
            
            row_text = "\n".join([f"{col}: {val}" for col, val in zip(columns, values)])
            
            self.frame_protocolos.clipboard_clear()
            self.frame_protocolos.clipboard_append(row_text)
            #self.status_label.config(text="Dados copiados para a área de transferência!")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao copiar:\n{str(e)}")
    
    def sort_column(self, col):
        """Sort tree by column"""
        # Get all items
        items = [(self.tree.set(item, col), item) for item in self.tree.get_children('')]
        
        # Sort (handles numeric sorting for ID)
        try:
            items.sort(key=lambda x: int(x[0]) if x[0].isdigit() else x[0])
        except:
            items.sort(key=lambda x: x[0])
        
        # Rearrange items
        for index, (val, item) in enumerate(items):
            self.tree.move(item, '', index)

class CertificateRequestForm: ### SOLICITAÇÃO DE CERTIFICADO
    
    def __init__(self, parent_notebook, str_conn, cor_fundo):
        self.str_conn = str_conn
        self.cor_fundo = cor_fundo
        self.frame_form = ttk.Frame(parent_notebook)
        parent_notebook.add(self.frame_form, text=" Formulário de Recebimento ")
        
        # Load customer list from database
        self.customers_list = []
        self.filtered_customers = []
        self.instruments_list = []
        self.instrument_types_list = []
        self.filtered_instrument_types = []
        self.load_customers()
        
        # Budgets and service orders lists
        self.budgets_list = []
        self.filtered_budgets = []
        self.service_orders_list = []
        self.filtered_service_orders = []
        
        # Form variables
        self.form_data = {
            'client_name': tk.StringVar(),
            'budget_code': tk.StringVar(),
            'service_order_code': tk.StringVar(),
            'instrument_name': tk.StringVar(),
            'instrument_model': tk.StringVar(),
            'instrument_serial': tk.StringVar(),
            'instrument_tag': tk.StringVar(),
            'calibration_type': tk.StringVar(),
            'required_points': tk.StringVar(),
            'special_requirements': tk.StringVar(),
            'request_date': tk.StringVar(value=datetime.now().strftime('%d/%m/%Y')),
            'requested_by': tk.StringVar(value=username),
            'priority': tk.StringVar(value='Normal')
        }
        
        # Track the currently selected customer ID
        self.selected_customer_id = None
        
        # Track suggestion windows
        self.suggestion_listbox = None
        self.suggestion_window = None
        self.current_suggestion_field = None  # Track which field is showing suggestions
        
        self.setup_form_ui()

    def load_instrument_types(self):
        """Load instrument types from database"""
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            rs = win32com.client.Dispatch("ADODB.Recordset")
            
            conn.Open(self.str_conn)
            
            sql = """
                SELECT DISTINCT 
                    it.name
                FROM instrument_types it
                WHERE it.name IS NOT NULL 
                    AND it.name != ''
                    AND it.removed = 0
                ORDER BY it.name
            """
            
            rs.Open(sql, conn)
            
            self.instrument_types_list = []
            if not rs.EOF:
                rs.MoveFirst()
                while not rs.EOF:
                    name = rs.Fields('name').Value
                    if name:
                        self.instrument_types_list.append(name)
                    rs.MoveNext()
            
            rs.Close()
            conn.Close()
            
            self.filtered_instrument_types = self.instrument_types_list.copy()
            return self.instrument_types_list
            
        except Exception as e:
            print(f"Erro ao carregar tipos de instrumento: {str(e)}")
            self.instrument_types_list = []
            self.filtered_instrument_types = []
            return []

    def add_instrument(self):
        """Add instrument to the list"""
        quantidade = self.entry_quantidade.get().strip()
        tag = self.entry_tag.get().strip()
        tipo_item = self.entry_tipo_item.get().strip()

        #if not quantidade:
        #    messagebox.showwarning("Aviso", "Preencha a quantidade.")
        #    return
        #
        #if not tipo_item:
        #    messagebox.showwarning("Aviso", "Selecione o tipo de item.")
        #    return
        
        #try:
        #    quantidade_int = int(quantidade)
        #    if quantidade_int <= 0:
        #        raise ValueError
        #except ValueError:
        #    messagebox.showwarning("Aviso", "Quantidade deve ser um número inteiro positivo.")
        #    return
        
        # Add to list
        instrument = { 'quantidade': quantidade if quantidade else '|', 'tag': tag if tag else '|', 'tipo_item': tipo_item if tipo_item else '|' }

        self.instruments_list.append(instrument)
        
        self.instruments_tree.insert("", "end", values=(quantidade if quantidade else '|', tag if tag else '|', tipo_item if tipo_item else '|'))
        
        # Clear input fields
        self.entry_quantidade.delete(0, 'end')
        self.entry_tag.delete(0, 'end')
        self.entry_tipo_item.delete(0, 'end')
        
        # Update status
        self.update_instruments_count()

    def remove_instrument(self):
        """Remove selected instrument from the list"""
        selected = self.instruments_tree.selection()
        if not selected:
            messagebox.showwarning("Aviso", "Selecione um item para remover.")
            return
        
        for item in selected:
            # Get the index of the selected item
            all_items = self.instruments_tree.get_children()
            index = all_items.index(item)
            
            # Remove from treeview
            self.instruments_tree.delete(item)
            
            # Remove from list
            if 0 <= index < len(self.instruments_list):
                self.instruments_list.pop(index)
        
        self.update_instruments_count()

    def show_instruments_context_menu(self, event):
        """Show right-click context menu for instruments tree"""
        item = self.instruments_tree.identify_row(event.y)
        
        if item:
            self.instruments_tree.selection_set(item)
            
            context_menu = tk.Menu(self.instruments_tree, tearoff=0)
            context_menu.add_command(label="Remover Item", command=self.remove_instrument)
            context_menu.post(event.x_root, event.y_root)

    def update_instruments_count(self):
        """Update the count label for instruments"""
        count = len(self.instruments_list)
        if count == 0:
            self.lbl_instruments_count.config(text="Nenhum item adicionado")
        elif count == 1:
            self.lbl_instruments_count.config(text="1 item adicionado")
        else:
            self.lbl_instruments_count.config(text=f"{count} itens adicionados")

    def load_customers(self):
        """Load customer aliases from database"""
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            rs = win32com.client.Dispatch("ADODB.Recordset")
            
            conn.Open(self.str_conn)
            
            sql = """
                SELECT DISTINCT 
                    c.id,
                    c.alias
                FROM  customers c
                WHERE c.alias IS NOT NULL 
                    AND c.alias != ''
                    AND c.removed = 0
                ORDER BY c.alias
            """
            
            rs.Open(sql, conn)
            
            self.customers_list = []
            if not rs.EOF:
                rs.MoveFirst()
                while not rs.EOF:
                    customer = {
                        'id': rs.Fields('id').Value if rs.Fields('id').Value else None,
                        'alias': rs.Fields('alias').Value if rs.Fields('alias').Value else ''
                    }
                    self.customers_list.append(customer)
                    rs.MoveNext()
            
            rs.Close()
            conn.Close()
            
            self.filtered_customers = self.customers_list.copy()
            #print(f"Carregados {len(self.customers_list)} clientes do banco de dados")
            
        except Exception as e:
            print(f"Erro ao carregar clientes: {str(e)}")
            self.customers_list = []
            self.filtered_customers = []
    
    def load_budgets_and_orders(self, customer_id):
        """Load budgets and service orders for a specific customer"""
        if not customer_id:
            return
        
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            rs = win32com.client.Dispatch("ADODB.Recordset")
            
            conn.Open(self.str_conn)
            
            # Load budgets for this customer
            sql_budgets = f"""
                SELECT DISTINCT 
                    b.code
                FROM budgets b
                WHERE b.id_customer = {customer_id}
                    AND b.code IS NOT NULL 
                    AND b.code != ''
                    AND b.removed = 0
                ORDER BY b.code
            """
            
            rs.Open(sql_budgets, conn)
            
            self.budgets_list = []
            if not rs.EOF:
                rs.MoveFirst()
                while not rs.EOF:
                    budget_code = rs.Fields('code').Value if rs.Fields('code').Value else ''
                    if budget_code:
                        self.budgets_list.append(budget_code)
                    rs.MoveNext()
            
            rs.Close()
            
            # Load service orders for this customer
            sql_orders = f"""
                SELECT DISTINCT 
                    os.code
                FROM orders_services os
                INNER JOIN budgets b ON b.id_order_service = os.id
                WHERE b.id_customer = {customer_id}
                    AND os.code IS NOT NULL 
                    AND os.code != ''
                    AND os.removed = 0
                ORDER BY os.code
            """
            
            rs.Open(sql_orders, conn)
            
            self.service_orders_list = []
            if not rs.EOF:
                rs.MoveFirst()
                while not rs.EOF:
                    order_code = rs.Fields('code').Value if rs.Fields('code').Value else ''
                    if order_code:
                        self.service_orders_list.append(order_code)
                    rs.MoveNext()
            
            rs.Close()
            conn.Close()
            
            #print(f"Carregados {len(self.budgets_list)} orçamentos e {len(self.service_orders_list)} ordens de serviço")
            
        except Exception as e:
            print(f"Erro ao carregar orçamentos e ordens: {str(e)}")
            self.budgets_list = []
            self.service_orders_list = []
    
    def filter_customers(self, search_text):
        """Filter customers based on search text"""
        if not search_text or search_text.strip() == '':
            self.filtered_customers = self.customers_list.copy()
        else:
            search_text = search_text.lower().strip()
            self.filtered_customers = [
                customer for customer in self.customers_list 
                if search_text in customer['alias'].lower()
            ]
        
        # Limit results to avoid performance issues
        if len(self.filtered_customers) > 50:
            self.filtered_customers = self.filtered_customers[:50]
    
    def filter_items(self, items_list, search_text):
        """Generic filter for budgets and service orders"""
        if not search_text or search_text.strip() == '':
            return items_list.copy()
        else:
            search_text = search_text.lower().strip()
            filtered = [item for item in items_list if search_text in str(item).lower()]
            
            # Limit results
            if len(filtered) > 50:
                filtered = filtered[:50]
            return filtered
    
    def setup_form_ui(self):
        # Main container with scrollbar
        main_canvas = tk.Canvas(self.frame_form, bg=self.cor_fundo, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.frame_form, orient="vertical", command=main_canvas.yview)
        scrollable_frame = ttk.Frame(main_canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all"))
        )
        
        main_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        main_canvas.configure(yscrollcommand=scrollbar.set)
        
        main_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        
                # Bind mouse wheel for scrolling
        def _on_mousewheel(event):
            main_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        main_canvas.bind_all("<MouseWheel>", _on_mousewheel)

        # Store the binding for later unbinding
        self._mousewheel_binding = _on_mousewheel

        # Title
        title_label = tk.Label(
            scrollable_frame, 
            text="ANEXO I - Recebimento e Devolução de Materiais", 
            font=("Segoe UI", 14, "bold"), 
            bg=self.cor_fundo, 
            fg="white"
        )
        title_label.pack(pady=(20, 30))
        
        # Client Information Section
        self.create_section(scrollable_frame, "DADOS DO CLIENTE")
        
        form_frame = ttk.Frame(scrollable_frame)
        form_frame.pack(pady=10, padx=20, fill="x")
        
        # Row 1: Company Name with Autocomplete
        ttk.Label(form_frame, text="Nome da empresa:").grid(row=0, column=0, sticky="w", pady=5, padx=5)
        
        # Frame for the autocomplete entry
        name_frame = ttk.Frame(form_frame)
        name_frame.grid(row=0, column=1, sticky="ew", pady=5, padx=5)
        
        # Entry widget
        self.entry_name = ttk.Entry(name_frame, textvariable=self.form_data['client_name'], width=58)
        self.entry_name.pack(fill="x")
        
        # Bind events for autocomplete
        self.entry_name.bind('<KeyRelease>', lambda e: self.on_field_keyrelease(e, 'client_name'))
        self.entry_name.bind('<FocusOut>', self.on_field_focusout)
        self.entry_name.bind('<Down>', self.on_arrow_down)
        self.entry_name.bind('<Up>', self.on_arrow_up)
        self.entry_name.bind('<Return>', self.on_return_key)
        self.entry_name.bind('<Escape>', self.hide_suggestions)
        
        # Row 2: Budget Code with Autocomplete
        ttk.Label(form_frame, text="Orçamento/Contrato:").grid(row=1, column=0, sticky="w", pady=5, padx=5)
        
        budget_frame = ttk.Frame(form_frame)
        budget_frame.grid(row=1, column=1, sticky="ew", pady=5, padx=5)
        
        self.entry_budget = ttk.Entry(budget_frame, textvariable=self.form_data['budget_code'], width=30)
        self.entry_budget.pack(fill="x")
        
        # Bind events for autocomplete
        self.entry_budget.bind('<KeyRelease>', lambda e: self.on_field_keyrelease(e, 'budget'))
        self.entry_budget.bind('<FocusOut>', self.on_field_focusout)
        self.entry_budget.bind('<Down>', self.on_arrow_down)
        self.entry_budget.bind('<Up>', self.on_arrow_up)
        self.entry_budget.bind('<Return>', self.on_return_key)
        self.entry_budget.bind('<Escape>', self.hide_suggestions)
        self.entry_budget.bind('<FocusIn>', self.on_budget_focusin)
        
        # Row 3: Service Order Code with Autocomplete
        ttk.Label(form_frame, text="Ordem de Serviço:").grid(row=2, column=0, sticky="w", pady=5, padx=5)
        
        order_frame = ttk.Frame(form_frame)
        order_frame.grid(row=2, column=1, sticky="ew", pady=5, padx=5)
        
        self.entry_order = ttk.Entry(order_frame, textvariable=self.form_data['service_order_code'], width=30)
        self.entry_order.pack(fill="x")
        
        # Bind events for autocomplete
        self.entry_order.bind('<KeyRelease>', lambda e: self.on_field_keyrelease(e, 'order'))
        self.entry_order.bind('<FocusOut>', self.on_field_focusout)
        self.entry_order.bind('<Down>', self.on_arrow_down)
        self.entry_order.bind('<Up>', self.on_arrow_up)
        self.entry_order.bind('<Return>', self.on_return_key)
        self.entry_order.bind('<Escape>', self.hide_suggestions)
        self.entry_order.bind('<FocusIn>', self.on_order_focusin)
        
        instrument_frame = ttk.Frame(scrollable_frame)
        instrument_frame.pack(pady=10, padx=20, fill="x")
        
        self.create_section(scrollable_frame, "INFORMAÇÕES ADICIONAIS")
        
        additional_frame = ttk.Frame(scrollable_frame)
        additional_frame.pack(pady=10, padx=20, fill="x")
        
        ttk.Label(additional_frame, text="Informações Adicionais:").grid(row=0, column=0, sticky="w", pady=5, padx=5)
        
        self.text_requirements = tk.Text(additional_frame, height=4, width=72)
        self.text_requirements.grid(row=1, column=0, columnspan=2, sticky="ew", pady=5, padx=5)
        
        def update_requirements(*args):
            self.form_data['special_requirements'].set(self.text_requirements.get("1.0", "end-1c"))
        
        self.text_requirements.bind('<KeyRelease>', update_requirements)
                
                # ===== INSTRUMENTS SECTION =====
        self.create_section(scrollable_frame, "DADOS DO MATERIAL/SERVIÇO RECEBIDO")

        instruments_form_frame = ttk.Frame(scrollable_frame)
        instruments_form_frame.pack(pady=10, padx=20, fill="x")

        # Row 0: Labels
        ttk.Label(instruments_form_frame, text="Quantidade:").grid(row=0, column=0, sticky="w", pady=(0, 2), padx=0)
        ttk.Label(instruments_form_frame, text="Tag:").grid(row=0, column=1, sticky="w", pady=(0, 2), padx=5)
        ttk.Label(instruments_form_frame, text="Tipo de item:").grid(row=0, column=2, sticky="w", pady=(0, 2), padx=5)

        # Row 1: Entry fields
        self.entry_quantidade = ttk.Entry(instruments_form_frame, width=10)
        self.entry_quantidade.grid(row=1, column=0, sticky="w", pady=2, padx=5)

        self.entry_tag = ttk.Entry(instruments_form_frame, width=20)
        self.entry_tag.grid(row=1, column=1, sticky="w", pady=2, padx=5)
        
        self.load_instrument_types()

        # Use Entry with autocomplete instead of Combobox
        self.entry_tipo_item = ttk.Entry(instruments_form_frame, width=42)
        self.entry_tipo_item.grid(row=1, column=2, sticky="w", pady=2, padx=5)

        # Bind events for autocomplete
        self.entry_tipo_item.bind('<KeyRelease>', lambda e: self.on_field_keyrelease(e, 'instrument_type'))
        self.entry_tipo_item.bind('<FocusOut>', self.on_field_focusout)
        self.entry_tipo_item.bind('<Down>', self.on_arrow_down)
        self.entry_tipo_item.bind('<Up>', self.on_arrow_up)
        self.entry_tipo_item.bind('<Return>', self.on_return_key)
        self.entry_tipo_item.bind('<Escape>', self.hide_suggestions)

        # Add button
        self.btn_add_instrument = ttk.Button(instruments_form_frame, text="Adicionar", command=self.add_instrument)
        self.btn_add_instrument.grid(row=1, column=3, sticky="w", pady=2, padx=10)

        # Bind Enter key in quantidade field to add instrument
        self.entry_quantidade.bind('<Return>', lambda e: self.add_instrument())
        self.entry_tag.bind('<Return>', lambda e: self.add_instrument())
        self.entry_tipo_item.bind('<Return>', self.on_instrument_return_key)

        # Instruments list (Treeview)
        instruments_list_frame = ttk.Frame(scrollable_frame)
        instruments_list_frame.pack(pady=5, padx=20, fill="x")

        # Treeview frame with scrollbar
        tree_container = ttk.Frame(instruments_list_frame)
        tree_container.pack(fill="x", expand=True)

        instruments_scroll_y = ttk.Scrollbar(tree_container)
        instruments_scroll_y.pack(side="right", fill="y")

        self.instruments_tree = ttk.Treeview(
            tree_container, 
            columns=("Quantidade", "Tag", "Tipo de Item"), 
            show="headings", 
            yscrollcommand=instruments_scroll_y.set,
            height=5
        )
        instruments_scroll_y.config(command=self.instruments_tree.yview)

        # Configure columns
        self.instruments_tree.heading("Quantidade", text="Quantidade")
        self.instruments_tree.heading("Tag", text="Tag")
        self.instruments_tree.heading("Tipo de Item", text="Tipo de Item")

        self.instruments_tree.column("Quantidade", width=80, minwidth=60)
        self.instruments_tree.column("Tag", width=150, minwidth=100)
        self.instruments_tree.column("Tipo de Item", width=250, minwidth=150)

        self.instruments_tree.pack(fill="x", expand=True)

        # Right-click context menu
        self.instruments_tree.bind('<Button-3>', self.show_instruments_context_menu)

        # Delete key to remove
        self.instruments_tree.bind('<Delete>', lambda e: self.remove_instrument())

        # Count label
        self.lbl_instruments_count = ttk.Label(
            instruments_list_frame, 
            text="Nenhum item adicionado", 
            font=("Segoe UI", 9, "italic")
        )
        self.lbl_instruments_count.pack(pady=(5, 0))

        request_frame = ttk.Frame(scrollable_frame)
        request_frame.pack(pady=10, padx=20, fill="x")
        
        ttk.Label(request_frame, text="Data da emissão:").grid(row=0, column=0, sticky="w", pady=5, padx=5)
        ttk.Label(request_frame, textvariable=self.form_data['request_date'], font=("Segoe UI", 10, "bold")).grid(row=0, column=1, sticky="w", pady=5, padx=5)
        ttk.Label(request_frame, text="Usuário emissor:").grid(row=1, column=0, sticky="w", pady=5, padx=5)
        ttk.Label(request_frame, textvariable=self.form_data['requested_by'], font=("Segoe UI", 10, "bold")).grid(row=1, column=1, sticky="w", pady=5, padx=5)
        
        button_frame = ttk.Frame(scrollable_frame)
        button_frame.pack(pady=20, padx=20, fill="x")
        
        self.btn_generate = ttk.Button( button_frame,  text="Criar documento",  command=self.generate_pdf, style='Accent.TButton' )
        self.btn_generate.pack(side="left", padx=5)
        
        self.btn_clear = ttk.Button( button_frame,  text="Limpar Formulário",  command=self.clear_form )
        self.btn_clear.pack(side="left", padx=5)
                
        # NEW BUTTON - Open list in popup window
        self.btn_open_list = ttk.Button( button_frame,  text="Abrir Lista de Recebimento",  command=self.open_protocol_list )
        self.btn_open_list.pack(side="left", padx=5)

        self.status_label = ttk.Label(scrollable_frame, text="", font=("Segoe UI", 9, "italic"))
        self.status_label.pack(pady=10)
    
    def on_budget_focusin(self, event):
        
        self.current_suggestion_field = 'budget'
        # If there's a customer selected but budgets haven't been loaded yet
        if self.selected_customer_id and not self.budgets_list:
            self.load_budgets_and_orders(self.selected_customer_id)

    def open_protocol_list(self):
        """Open the protocol list in a popup window"""
        popup = tk.Toplevel(self.frame_form)
        popup.title("Lista de Recebimento")
        popup.geometry("1000x550")
        popup.configure(bg=self.cor_fundo)
        popup.minsize(1000, 550)
        popup.maxsize(1000, 550)
        
        # Unbind mouse wheel from main window while popup is open
        self.frame_form.unbind_all("<MouseWheel>")
        
        def on_popup_close():
            """Re-bind mouse wheel to main window when popup closes"""
            self.frame_form.bind_all("<MouseWheel>", self._mousewheel_binding)
            popup.destroy()
        
        # Handle window close (X button)
        popup.protocol("WM_DELETE_WINDOW", on_popup_close)
        
        # Create a ProtocolViewer instance inside the popup
        temp_notebook = ttk.Notebook(popup)
        temp_notebook.pack(expand=True, fill="both", padx=5, pady=5)
        
        # Create the ProtocolViewer in the popup
        protocol_viewer = ProtocolViewer(temp_notebook, STR_CONN, self.cor_fundo)
        
        # Add a close button at the bottom
        close_frame = ttk.Frame(popup)
        close_frame.pack(pady=10)
        ttk.Button(close_frame, text="Fechar", command=on_popup_close).pack()
        
        # Center the popup on screen
        popup.update_idletasks()
        #x = popup.winfo_screenwidth() // 2 - 450
        #y = popup.winfo_screenheight() // 2 - 275
        #popup.geometry(f"+{x}+{y}")

    def on_order_focusin(self, event):
        
        self.current_suggestion_field = 'order'
        # If there's a customer selected but orders haven't been loaded yet
        if self.selected_customer_id and not self.service_orders_list:
            self.load_budgets_and_orders(self.selected_customer_id)
    
    def on_field_keyrelease(self, event, field_type):
        """Handle key release in any autocomplete field"""
        # Ignore special keys
        if event.keysym in ('Up', 'Down', 'Left', 'Right', 'Return', 'Escape', 'Tab', 'Shift_L', 'Shift_R', 'Control_L', 'Control_R', 'Alt_L', 'Alt_R'):
            return
        
        # Set current field type
        self.current_suggestion_field = field_type
        
        # Hide previous suggestions
        self.hide_suggestions()
        
        if field_type == 'client_name':
            search_text = self.form_data['client_name'].get()
            self.filter_customers(search_text)
            
            if self.filtered_customers and search_text.strip():
                self.show_suggestions(self.filtered_customers, 'client_name')
                
        elif field_type == 'budget':
            search_text = self.form_data['budget_code'].get()
            self.filtered_budgets = self.filter_items(self.budgets_list, search_text)
            
            if self.filtered_budgets and search_text.strip():
                self.show_suggestions(self.filtered_budgets, 'budget')
                
        elif field_type == 'order':
            search_text = self.form_data['service_order_code'].get()
            self.filtered_service_orders = self.filter_items(self.service_orders_list, search_text)
            
            if self.filtered_service_orders and search_text.strip():
                self.show_suggestions(self.filtered_service_orders, 'order')
                
        elif field_type == 'instrument_type':
            search_text = self.entry_tipo_item.get()
            self.filtered_instrument_types = self.filter_items(self.instrument_types_list, search_text)
            
            if self.filtered_instrument_types and search_text.strip():
                self.show_suggestions(self.filtered_instrument_types, 'instrument_type')

    def on_field_focusout(self, event):
        """Hide suggestions when entry loses focus, with delay"""
        self.frame_form.after(200, self.hide_suggestions)
    
    def on_arrow_down(self, event):
        """Navigate down in suggestion list"""
        if self.suggestion_listbox and self.suggestion_window:
            try:
                current_selection = self.suggestion_listbox.curselection()
                if current_selection:
                    next_index = min(current_selection[0] + 1, self.suggestion_listbox.size() - 1)
                else:
                    next_index = 0
                
                self.suggestion_listbox.selection_clear(0, tk.END)
                self.suggestion_listbox.selection_set(next_index)
                self.suggestion_listbox.activate(next_index)
                self.suggestion_listbox.see(next_index)
            except:
                pass
        return "break"
    
    def on_arrow_up(self, event):
        """Navigate up in suggestion list"""
        if self.suggestion_listbox and self.suggestion_window:
            try:
                current_selection = self.suggestion_listbox.curselection()
                if current_selection:
                    prev_index = max(current_selection[0] - 1, 0)
                else:
                    prev_index = self.suggestion_listbox.size() - 1
                
                self.suggestion_listbox.selection_clear(0, tk.END)
                self.suggestion_listbox.selection_set(prev_index)
                self.suggestion_listbox.activate(prev_index)
                self.suggestion_listbox.see(prev_index)
            except:
                pass
        return "break"
    
    def on_return_key(self, event):
        """Select highlighted suggestion on Enter"""
        if self.suggestion_listbox and self.suggestion_window:
            try:
                selection = self.suggestion_listbox.curselection()
                if selection:
                    selected_index = selection[0]
                    self.select_suggestion(selected_index)
                    self.hide_suggestions()
                    return "break"
            except:
                pass
        return
    
    def show_suggestions(self, items, field_type):
        """Display the suggestion dropdown for any field"""
        if not items:
            return
        
        # Determine which entry widget to position under
        if field_type == 'client_name':
            entry_widget = self.entry_name
        elif field_type == 'budget':
            entry_widget = self.entry_budget
        elif field_type == 'order':
            entry_widget = self.entry_order
        elif field_type == 'instrument_type':
            entry_widget = self.entry_tipo_item
        else:
            return
        
        # Create popup window for suggestions
        self.suggestion_window = tk.Toplevel(self.frame_form)
        self.suggestion_window.wm_overrideredirect(True)
        
        # Position the window below the entry widget
        entry_x = entry_widget.winfo_rootx()
        entry_y = entry_widget.winfo_rooty() + entry_widget.winfo_height()
        entry_width = entry_widget.winfo_width()
        
        self.suggestion_window.geometry(f"{entry_width}x200+{entry_x}+{entry_y}")
        
        # Create listbox with scrollbar
        list_frame = tk.Frame(self.suggestion_window, bg='white')
        list_frame.pack(fill='both', expand=True)
        
        scrollbar = tk.Scrollbar(list_frame, orient='vertical')
        scrollbar.pack(side='right', fill='y')
        
        self.suggestion_listbox = tk.Listbox(
            list_frame, 
            yscrollcommand=scrollbar.set,
            bg='white',
            fg='black',
            font=('Segoe UI', 9),
            selectmode='single',
            activestyle='none',
            height=10
        )
        self.suggestion_listbox.pack(side='left', fill='both', expand=True)
        scrollbar.config(command=self.suggestion_listbox.yview)
        
        # Populate listbox
        if field_type == 'client_name':
            for customer in items[:50]:
                self.suggestion_listbox.insert(tk.END, customer['alias'])
        else:
            for item in items[:50]:
                self.suggestion_listbox.insert(tk.END, str(item))
        
        # Bind events
        self.suggestion_listbox.bind('<ButtonRelease-1>', lambda e: self.on_suggestion_click(e, field_type))
        self.suggestion_listbox.bind('<Double-Button-1>', lambda e: self.on_suggestion_doubleclick(e, field_type))
        
        # Set focus back to entry
        entry_widget.focus_set()

    def hide_suggestions(self, event=None):
        """Hide the suggestion dropdown"""
        if self.suggestion_window:
            try:
                self.suggestion_window.destroy()
            except:
                pass
            finally:
                self.suggestion_window = None
                self.suggestion_listbox = None
    
    def on_suggestion_click(self, event, field_type):
        """Handle single click on suggestion"""
        if self.suggestion_listbox:
            selection = self.suggestion_listbox.curselection()
            if selection:
                selected_index = selection[0]
                self.select_suggestion(selected_index)
                self.hide_suggestions()
    
    def on_suggestion_doubleclick(self, event, field_type):
        """Handle double click on suggestion"""
        if self.suggestion_listbox:
            selection = self.suggestion_listbox.curselection()
            if selection:
                selected_index = selection[0]
                self.select_suggestion(selected_index)
                self.hide_suggestions()
    
    def select_suggestion(self, index):
        """Select a suggestion based on current field type"""
        try:
            if self.current_suggestion_field == 'client_name':
                if 0 <= index < len(self.filtered_customers):
                    customer = self.filtered_customers[index]
                    self.form_data['client_name'].set(customer['alias'])
                    self.selected_customer_id = customer['id']
                    self.form_data['budget_code'].set('')
                    self.form_data['service_order_code'].set('')
                    self.load_budgets_and_orders(customer['id'])
            
            elif self.current_suggestion_field == 'budget':
                if 0 <= index < len(self.filtered_budgets):
                    self.form_data['budget_code'].set(str(self.filtered_budgets[index]))
                    self.status_label.config(text=f"Orçamento selecionado: {self.filtered_budgets[index]}")
            
            elif self.current_suggestion_field == 'order':
                if 0 <= index < len(self.filtered_service_orders):
                    self.form_data['service_order_code'].set(str(self.filtered_service_orders[index]))
                    self.status_label.config(text=f"Ordem de Serviço selecionada: {self.filtered_service_orders[index]}")
            
            elif self.current_suggestion_field == 'instrument_type':
                if 0 <= index < len(self.filtered_instrument_types):
                    self.entry_tipo_item.delete(0, 'end')
                    self.entry_tipo_item.insert(0, str(self.filtered_instrument_types[index]))
                    self.status_label.config(text=f"Tipo de item selecionado: {self.filtered_instrument_types[index]}")
        
        except Exception as e:
            print(f"Erro ao selecionar sugestão: {str(e)}")

    def create_section(self, parent, title):
        """Create a section header"""
        section_frame = ttk.Frame(parent)
        section_frame.pack(pady=(5, 5), padx=20, fill="x")
        
        #ttk.Separator(section_frame, orient='horizontal').pack(fill="x")
        tk.Label(
            section_frame, 
            text=title, 
            font=("Segoe UI", 10, "bold"), 
            bg=self.cor_fundo, 
            fg="white"
        ).pack(pady=5)
    
    def validate_form(self):
        """Validate required fields"""
        required_fields = {
            #'client_name': 'Razão Social/Nome',
            #'instrument_name': 'Instrumento',
            #'instrument_serial': 'Nº Série',
            }
        
        missing_fields = []
        for field, label in required_fields.items():
            if not self.form_data[field].get().strip():
                missing_fields.append(label)
        
        if missing_fields:
            messagebox.showerror( "Campos Obrigatórios",  f"Preencha os seguintes campos:\n\n" + "\n".join(f"• {f}" for f in missing_fields) )
            return False
        return True
    
    def generate_pdf(self):
        """Generate PDF with form data"""
        if not self.validate_form():
            return
        
        # Ask for save location
        filename = filedialog.asksaveasfilename( defaultextension=".pdf", filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")], initialfile=f"DR {datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf" )
        
        if not filename:
            return
        
        try:
            self._create_pdf(filename)
            #self.status_label.config(text=f"PDF gerado com sucesso: {os.path.basename(filename)}")
            messagebox.showinfo("Sucesso", f"PDF gerado com sucesso!\n\n{filename}")
            
            os.startfile(filename)
                
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao gerar PDF:\n{str(e)}")
            self.status_label.config(text="Erro ao gerar PDF")
    
    def _create_pdf(self, filename):
        
        # ===== INSERT PROTOCOL INTO DATABASE =====
        protocol_result = insert_protocol_and_get_id(
            STR_CONN,
            username,
            self.form_data['client_name'].get(),
            self.form_data['budget_code'].get(),
            self.form_data['service_order_code'].get(),
            os.path.basename(filename)
        )
        
        if protocol_result:
            protocol_code = protocol_result['protocol_code']
            print(f"Protocolo registrado: ID {protocol_result['id']} - {protocol_code}")
        else:
            protocol_code = f"{datetime.now().strftime('%Y%m%d')}-{random.randint(1000, 9999)}"
            print("Usando protocolo fallback")

        # ===== CONFIGURABLE TABLE POSITION AND SIZE =====
        first_col_width = 55 * mm
        second_col_width = 115 * mm
        
        # ===== DEFINE HEADER AND FOOTER FUNCTIONS =====
        def add_header(canvas, doc):
            """Add header to each page"""
            canvas.saveState()
            
            # Solicitation type
            canvas.setFont('Helvetica', 9)
            canvas.setFillColor(HexColor("#000000"))
            canvas.drawString(20*mm, A4[1] - 21*mm, "SERVIÇO NACIONAL DE APRENDIZAGEM INDUSTRIAL - SENAI")
            
            right_header = f"SENAI-UO631-PR-00004 rev. 5"
            canvas.drawString(A4[0] - 65*mm, A4[1] - 21*mm, right_header)

            # Horizontal line below header
            canvas.setStrokeColor(HexColor("#000000"))
            canvas.setLineWidth(1)
            canvas.line(20*mm, A4[1] - 24*mm, A4[0] - 20*mm, A4[1] - 24*mm)
            
            canvas.restoreState()
        
        def add_footer(canvas, doc):
            """Add footer to each page"""
            canvas.saveState()
            
            # Horizontal line above footer
            canvas.setStrokeColor(HexColor("#000000"))
            canvas.setLineWidth(1)
            canvas.line(20*mm, 20*mm, A4[0] - 20*mm, 20*mm)
            
            # Footer text
            canvas.setFont('Helvetica', 7)
            canvas.setFillColor(HexColor("#000000"))
            
            left_footer = f"Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}"
            canvas.drawString(20*mm, 13*mm, left_footer)
            
            right_footer = f"Página {doc.page}"
            canvas.drawRightString(A4[0] - 20*mm, 13*mm, right_footer)
            
            canvas.restoreState()
        
        # ===== CREATE DOCUMENT =====
        doc = BaseDocTemplate(
            filename,
            pagesize=A4,
            rightMargin=20*mm,
            leftMargin=20*mm,
            topMargin=30*mm,
            bottomMargin=25*mm,
            title='ANEXO I - Recebimento e Devolução de Materiais',
            author=username
        )
        
        # Create a frame that covers the full content area
        frame = Frame(
            doc.leftMargin,
            doc.bottomMargin,
            doc.width,
            doc.height,
            id='normal'
        )
        
        template = PageTemplate(
            id='main',
            frames=[frame],
            onPage=add_header,
            onPageEnd=add_footer
        )
        
        doc.addPageTemplates([template])
        
        # ===== STYLES =====
        styles = getSampleStyleSheet()
        
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=12,
            textColor=HexColor("#000000"),
            spaceAfter=10,
            alignment=TA_CENTER
        )
        
        label_style = ParagraphStyle(
            'LabelStyle',
            parent=styles['Normal'],
            fontSize=9,
            fontName='Helvetica-Bold',
            textColor=HexColor('#333333'),
            leading=12,
            alignment=TA_LEFT
        )

        label_style2 = ParagraphStyle(
            'LabelStyle',
            parent=styles['Normal'],
            fontSize=11,
            fontName='Helvetica-Bold',
            textColor=HexColor('#333333'),
            leading=12,
            alignment=TA_CENTER
        )
        
        value_style = ParagraphStyle(
            'ValueStyle',
            parent=styles['Normal'],
            fontSize=9,
            fontName='Helvetica',
            textColor=HexColor('#000000'),
            leading=12,
            alignment=TA_LEFT
        )
        
        section_style = ParagraphStyle(
            'SectionStyle',
            parent=styles['Normal'],
            fontSize=10,
            fontName='Helvetica-Bold',
            textColor=HexColor('#1B4B9F'),
            leading=14,
            alignment=TA_LEFT
        )
        
        centered_style = ParagraphStyle(
            'CenteredStyle',
            parent=styles['Normal'],
            fontSize=10,
            fontName='Helvetica',
            textColor=HexColor('#000000'),
            leading=14,
            alignment=TA_CENTER
        )
        
        # ===== BUILD CONTENT =====
        story = []
        
        # ===== ADD LOGO AS A FLOWABLE =====
        if LOGO_BLACK_BASE64:
            try:
                base64_str = LOGO_BLACK_BASE64
                
                if "," in base64_str:
                    base64_str = base64_str.split(",")[1]
                    
                image_bytes = base64.b64decode(base64_str)
                image_buffer = io.BytesIO(image_bytes)
                
                h = 30
                w = int(h * 3.74703357)
                img_flowable = rp.Image(image_buffer, width=w, height=h)
                # Create a single-cell table that spans the full page width and centers the image
                logo_table = Table(
                    [[img_flowable]], 
                    colWidths=[doc.width]  # Full page width
                )

                logo_table.setStyle(TableStyle([
                    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),  # Center the image horizontally
                    ('LEFTPADDING', (0, 0), (-1, -1), 0),
                    ('RIGHTPADDING', (0, 0), (-1, -1), 0),
                    ('TOPPADDING', (0, 0), (-1, -1), 0),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
                ]))
                
                story.append(logo_table)
                story.append(Spacer(1, 5 * mm))
                
            except Exception as e:
                print(f"Erro ao adicionar logo: {str(e)}")
        
        # Title
        story.append(Paragraph("ANEXO I - Recebimento e Devolução de Materiais", title_style))
        story.append(Spacer(1, 0.5 * mm))
        
        # ===== CREATE TABLE DATA =====
        table_data = []
        
        table_data.append([Paragraph(f"<b>N°:</b> {protocol_code}", label_style), Paragraph(f"<b>Data de chegada:</b> {datetime.now().strftime('%d/%m/%Y')}", label_style)])
        
        table_data.append([Paragraph(f"<b>RECEBIMENTO</b>", centered_style)])
        
        table_data.append([Paragraph("<b>Nome da empresa:</b>", label_style), Paragraph(self.form_data['client_name'].get(), value_style)])
        
        table_data.append([Paragraph("<b>Orçamento/Contrato:</b>", label_style), 
        Paragraph(self.form_data['budget_code'].get(), value_style)])
        
        table_data.append([Paragraph("<b>Ordem de Serviço:</b>", label_style), Paragraph(self.form_data['service_order_code'].get(), value_style)])
        
                # Calculate total quantity
        total_quantidade = sum(int(item['quantidade']) for item in self.instruments_list if item['quantidade'].isdigit())
        quantidade_text = str(total_quantidade) if total_quantidade > 0 else ""

        table_data.append([
            Paragraph("<b>Quantidade de itens:</b>", label_style), 
            Paragraph(quantidade_text, value_style)
        ])

        table_data.append([Paragraph("<b>Assinatura:</b>", label_style), Paragraph("", value_style)])
        
        # Separator row
        table_data.append([Paragraph("<b>DEVOLUÇÃO</b>", centered_style)])
        
        table_data.append([Paragraph("<b>Data de devolução:</b>", label_style),Paragraph(self.form_data['instrument_name'].get(), value_style)])
        
        table_data.append([
            Paragraph("<b>Entregue para:</b>", label_style),
            Paragraph(self.form_data['instrument_serial'].get(), value_style)
        ])
        
        table_data.append([
            Paragraph("<b>Assinatura:</b>", label_style),
            Paragraph("", value_style)
        ])
        
        # Additional Information
        if self.form_data['special_requirements'].get().strip():
            table_data.append([Paragraph("<b>INFORMAÇÕES ADICIONAIS</b>", centered_style)])
            additional_info = self.form_data['special_requirements'].get().replace('\n', '<br/>')
            table_data.append([Paragraph(additional_info, value_style)])
        
        # ===== BUILD TABLE =====
        formatted_rows = []
        for row in table_data:
            formatted_rows.append(list(row))
        
        table = Table(formatted_rows, colWidths=[first_col_width, second_col_width])
        
        # ===== APPLY TABLE STYLE =====
        table_style = TableStyle([
            ('GRID', (0, 0), (-1, -1), 0.5, HexColor("#000000")),
            ('BOX', (0, 0), (-1, -1), 1, HexColor("#000000")),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('LEFTPADDING', (0, 0), (-1, -1), 6),
            ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ])
        
        # Apply span and background for single-column rows
        row_index = 0
        for row in formatted_rows:
            if len(row) == 1:
                table_style.add('SPAN', (0, row_index), (1, row_index))
                table_style.add('BACKGROUND', (0, row_index), (-1, row_index), HexColor('#E8F0FE'))
                table_style.add('TEXTCOLOR', (0, row_index), (-1, row_index), HexColor("#000000"))
            row_index += 1
        
        table.setStyle(table_style)
        story.append(table)
                
                # ===== INSTRUMENTS TABLE IN PDF =====
        if self.instruments_list:
            story.append(Spacer(1, 5 * mm))
            story.append(Paragraph("<b>DADOS DO MATERIAL/SERVIÇO RECEBIDO</b>", label_style2))
            story.append(Spacer(1, 2 * mm))
            
            # Create instruments table data
            inst_table_data = []
            
            # Header row
            inst_table_data.append([
                Paragraph("<b>Quantidade</b>", label_style),
                Paragraph("<b>Tag</b>", label_style),
                Paragraph("<b>Tipo de Item</b>", label_style)
            ])
            
            # Data rows
            for item in self.instruments_list:
                inst_table_data.append([
                    Paragraph(item['quantidade'], value_style),
                    Paragraph(item['tag'], value_style),
                    Paragraph(item['tipo_item'], value_style)
                ])
            
            # Create instruments table
            inst_table = Table(inst_table_data, colWidths=[30*mm, 40*mm, 100*mm])
            
            inst_table_style = TableStyle([
                ('GRID', (0, 0), (-1, -1), 0.5, HexColor("#000000")),
                ('BOX', (0, 0), (-1, -1), 1, HexColor("#000000")),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('TOPPADDING', (0, 0), (-1, -1), 4),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
                ('LEFTPADDING', (0, 0), (-1, -1), 6),
                ('RIGHTPADDING', (0, 0), (-1, -1), 6),
                ('BACKGROUND', (0, 0), (-1, 0), HexColor('#E8F0FE')),  # Header row
            ])
            
            inst_table.setStyle(inst_table_style)
            story.append(inst_table)

        doc.build(story)

    def _draw_rectangle(self, canvas, x, y, width, height, line_width, line_color, fill_color=None):
        """Helper method to draw a rectangle on the canvas"""
        
        # Draw fill if specified
        if fill_color:
            canvas.setFillColor(fill_color)
            canvas.rect(x, y, width, height, fill=1, stroke=0)
        
        # Draw border
        canvas.setStrokeColor(line_color)
        canvas.setLineWidth(line_width)
        canvas.rect(x, y, width, height, fill=0, stroke=1)


    def clear_form(self):
        """Clear all form fields"""
        for key in self.form_data:
            if key in ['request_date', 'requested_by', 'priority']:
                continue
            self.form_data[key].set('')
        
        # Clear text widget
        self.text_requirements.delete("1.0", "end")
        
        # Clear instruments
        for item in self.instruments_tree.get_children():
            self.instruments_tree.delete(item)
        self.instruments_list = []
        self.update_instruments_count()
        
        # Clear instrument input fields
        self.entry_quantidade.delete(0, 'end')
        self.entry_tag.delete(0, 'end')
        self.entry_tipo_item.delete(0, 'end')
        
        self.selected_customer_id = None
        self.budgets_list = []
        self.service_orders_list = []
        self.filtered_budgets = []
        self.filtered_service_orders = []
        self.hide_suggestions()

    def on_instrument_return_key(self, event):
        """Handle Return key in instrument type field"""
        # If suggestions are visible, select the highlighted one
        if self.suggestion_listbox and self.suggestion_window and self.current_suggestion_field == 'instrument_type':
            try:
                selection = self.suggestion_listbox.curselection()
                if selection:
                    selected_index = selection[0]
                    self.select_suggestion(selected_index)
                    self.hide_suggestions()
                    return "break"
            except:
                pass
        
        # If no suggestion selected, add the instrument
        self.add_instrument()
        return "break"
# ==================== MAIN ====================

version_check = check_app_version(STR_CONN, APP_NAME, APP_VERSION)

if version_check['update_required']:
    messagebox.showerror("Versão Obsoleta", version_check['message'])
    sys.exit()
else:
    print(version_check['message'])

status_servidor, cor_status = verificar_disponibilidade3()
status_servidor3, cor_status3 = status_servidor, cor_status

root = tk.Tk()

if bool_agenda == 1:
    root.title("Assistente de Agenda [beta 1]")
elif bool_agenda == 0:
    root.title("Assistente de Dados [beta 1]")
    
root.geometry("550x675")
root.minsize(650, 675)
show_loading_screen(root,f"O {APP_NAME} está inicializando...")

cor_fundo = _from_rgb((27, 75, 159))
root.configure(bg=cor_fundo)

style = ttk.Style()
style.theme_use('vista')

style.configure("TNotebook", background=cor_fundo, borderwidth=0, padding=0)
style.configure("TNotebook.Tab", font=("Segoe UI", 9, "bold"), padding=[5, 5],borderwidth=0)
style.configure("TFrame", background=cor_fundo, borderwidth=0, relief='flat')
style.configure("TLabel", background=cor_fundo, foreground="white", font=("Segoe UI", 10))

notebook = ttk.Notebook(root)
notebook.pack(expand=True, fill="both")

if bool_agenda == 1:
    db_viewer2 = DatabaseViewer2(notebook, STR_CONN_LINKED, cor_fundo)         ### PADRÕES
    cert_form = CertificateRequestForm(notebook, STR_CONN_LINKED, cor_fundo)   ### SOLICITAÇÃO DE CERTIFICADO
    #protocol_viewer = ProtocolViewer(notebook, STR_CONN, cor_fundo)            ### PROTOCOLOS

hide_loading_screen()
root.tkraise()
#db_viewer_orcamentos = DatabaseViewerOrcamentos(notebook, STR_CONN_LINKED, cor_fundo)

#for lab_code, lab_config in LABS.items():
#    
#    ServiceScheduler(notebook, STR_CONN, cor_fundo,id_sector=lab_config['sector_id'],lab_name=lab_config['name'],lab_config=lab_config)
#    ServiceSchedulerLinkedDirect(notebook, STR_CONN_LINKED, STR_CONN, cor_fundo,lab_code=lab_code,lab_config=lab_config)
    
frame_info = ttk.Frame(notebook)
notebook.add(frame_info, text=" Dados de Conexão & Informações ")

main_canvas = tk.Canvas(frame_info, bg=cor_fundo, highlightthickness=0)
main_scrollbar = ttk.Scrollbar(frame_info, orient="vertical", command=main_canvas.yview)
main_canvas.configure(yscrollcommand=main_scrollbar.set)
main_canvas.pack(side="left", expand=True, fill="both")
main_scrollbar.pack(side="right", fill="y")

content_frame = ttk.Frame(main_canvas)
main_canvas.create_window((0, 0), window=content_frame, anchor="nw", width=main_canvas.winfo_width())

def configure_scroll_region(event):
    main_canvas.configure(scrollregion=main_canvas.bbox("all"))
    
content_frame.bind("<Configure>", configure_scroll_region)
main_canvas.bind("<Configure>", lambda e: main_canvas.itemconfig("all", width=e.width))

ttk.Label(content_frame, text=f"Status do Servidor {db} {str_alternativa}").pack(pady=(20, 0))
label_status = tk.Label(content_frame, text=status_servidor, fg=cor_status, bg=cor_fundo, font=("Segoe UI", 10, "bold"))
label_status.pack(pady=5)

ttk.Label(content_frame, text=f"Status do Servidor {db_linked}").pack(pady=(0, 0))
label_status3 = tk.Label(content_frame, text=status_servidor3, fg=cor_status3, bg=cor_fundo, font=("Segoe UI", 10, "bold"))
label_status3.pack(pady=5)

info_texto = f"Driver: SQLOLEDB\nVersão: {INFO_VERSAO}"
label_info = ttk.Label(content_frame, text=info_texto, justify="center", font=("Consolas", 9))
label_info.pack(pady=15)

footer_frame = ttk.Frame(content_frame)
footer_frame.pack(side="bottom", pady=10, fill="x")

credits = ttk.Label(footer_frame, text="Castro", font=("Segoe UI", 8, "italic"), cursor="hand2")
credits.pack()
ttk.Label(footer_frame, text="2026 IST PGE - Laboratório de Metrologia", font=("Segoe UI", 8, "italic")).pack()

def open_link(event):
    webbrowser.open("https://teams.microsoft.com/l/chat/0/0?users=guilherme.castro@senairs.org.br&message=E%20aí%20meu%20chapa")
credits.bind("<Button-1>", open_link)


root.mainloop()

#PIL PIL._imagingtk PIL._tkinter_finder

# 5e6ee1b4689b5261d9f9fafe2f83aad4082df04a10ab6d2846094b97ad33a2e7 1701198326
# cd2303597b614f903daa0dcf2be479b6a53cb045c1c4acd9071ea9ab95fb706b 198326
