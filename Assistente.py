import win32com.client
import tkinter as tk
import win32timezone
from tkinter import ttk, messagebox
from datetime import datetime, timedelta, date
import random
import webbrowser
import calendar as cal
from dateutil.relativedelta import relativedelta        
from getpass import getuser
from urllib.parse import quote
import socket
from datetime import time, datetime, timedelta, date


def get_server_ip(hostname='UO061M4118173'):
    try:
        return socket.gethostbyname(hostname)
    except:
        return None

constante = '10.165.212.74'

ip = get_server_ip('UO061M4118173') or '10.165.212.74'
ip_zebra = ip

if ip != constante:
    str_alternativa = f'[Conectado ao IP alternativo: {ip}]'
else:
    str_alternativa = ''

print(f"Conectando ao servidor: {ip}")

db = "IST-PGE"

STR_CONN = (
    f"Provider=SQLOLEDB;"
    f"Data Source={ip};"
    f"Initial Catalog={db};"
    f"User ID=sa;"
    f"Password=Wheelp0p2;"
)

STR_CONN_LINKED = (
    f"Provider=SQLOLEDB;"
    f"Data Source={ip_linked};"
    f"Initial Catalog={db_linked};"
    f"User ID={user_linked};"
    f"Password={password_linked};"
)


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
    f' Oi olá  ',
    f' Ora ora  ',
    f' Buenas, {username}?  ',
    f' {username}?!  ',
    f' "Nani??"  ',
    f' Tudo bom, {username}?  ',
    f' Hi there, {username}!  ',
    f' Hello there, {username}!  ',
    f' Hellooo!  ',
    f' Fala aí {username}, tudo bom?  ',
    f' Tudo bão {username}?  ',
    f' Bão!?  ',
    f' E aí {username}, bão!?  ',
    f' {username}!?  ',
    f' Alô alô {username}!  ',
    f' E aí {username}, tudo certo?  ',    
    f' Seu nome é Gabriel?  ',    
    f' Opa opa  ',    
    f' Aoba  ',    
    f' Tudo certo?  ',
    f'  ¯|_(ツ)_|¯   ',    
    f'  *_*   ',
    f'  Buenas tardes!  ',
    f'  Vai um chimas?  '
]

buenas = random.choice(greetings)

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
        conn.ConnectionTimeout = 12
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


# ==================== SERVICE SCHEDULER (MULTI-LAB) ====================
class ServiceScheduler:
    """ ABA: Agendamento de Serviços (Calibração FIFO) - Multi-Laboratório """
    
    def __init__(self, parent_notebook, str_conn, cor_fundo, id_sector=None, lab_name="", lab_config=None):
        self.str_conn = str_conn
        self.cor_fundo = cor_fundo
        self.id_sector = id_sector
        self.lab_name = lab_name
        self.lab_config = lab_config or {}
        
        tab_name = f" Calendário de {self.lab_name} "
        self.frame_certificados = ttk.Frame(parent_notebook)
        parent_notebook.add(self.frame_certificados, text=tab_name)
        
        parent_notebook.bind('<<NotebookTabChanged>>', self.on_tab_changed)
        
        self.current_month_offset = 0
        self.calendar_data = {}
        self.selected_date = None
        
        self.setup_ui()
        self.load_services()
        self.load_calendar_data()
        self.render_calendar()
        self.update_queue_count()

    def on_tab_changed(self, event):
        notebook = event.widget
        current_tab = notebook.select()
        current_tab_text = notebook.tab(current_tab, "text")
        
        if self.lab_name in current_tab_text:
            self.refresh_calendar()

    def show_loading_screen(self, message="Carregando..."):
        """Show a loading overlay with dynamic info"""
        self.loading_popup = tk.Toplevel(self.frame_certificados)
        self.loading_popup.geometry("350x130")
        self.loading_popup.configure(bg=self.cor_fundo)
        self.loading_popup.title("")
        self.loading_popup.overrideredirect(True)
        self.loading_popup.attributes('-topmost', True)
        
        # Center on parent
        self.loading_popup.update_idletasks()
        x = self.frame_certificados.winfo_rootx() + (self.frame_certificados.winfo_width() // 2) - 175
        y = self.frame_certificados.winfo_rooty() + (self.frame_certificados.winfo_height() // 2) - 65
        self.loading_popup.geometry(f"+{x}+{y}")
        
        frame = tk.Frame(self.loading_popup, bg=self.cor_fundo, highlightbackground="#FFFFFF", highlightthickness=2)
        frame.pack(fill='both', expand=True, padx=2, pady=2)

        tk.Label(frame, text=f"Oioi, {username}! O app está carregando:", font=('Segoe UI', 10), background='black', fg="#FFFFFF").pack(pady=(15, 5))

        #tk.Label(frame, text=f"Laboratório de {self.lab_name}", font=('Segoe UI', 13, 'bold'), bg=self.cor_fundo, fg="#FFFFFF").pack(pady=(5, 5))
        
        # Message
        tk.Label(frame, text=message, font=('Segoe UI', 11),bg=self.cor_fundo, fg='white').pack(pady=(5, 5))
        
        self.loading_popup.grab_set()
        self.loading_popup.update()

    def hide_loading_screen(self):
        """Hide the loading overlay"""
        if hasattr(self, 'loading_popup') and self.loading_popup.winfo_exists():
            self.loading_popup.grab_release()
            self.loading_popup.destroy()

    def setup_ui(self):
        main_container = tk.Frame(self.frame_certificados, bg=self.cor_fundo)
        main_container.pack(fill='both', expand=True, padx=10, pady=10)
        
        self.calendar_frame = tk.Frame(main_container, bg=self.cor_fundo)
        self.calendar_frame.pack(side='left', fill='both', expand=True)
        
        self.status_label = ttk.Label(self.frame_certificados, text="Pronto para agendamento", font=("Segoe UI", 12))
        self.status_label.pack(pady=5)
        

    def open_teams_chat(self):
        try:
            chat_id = self.lab_config.get('teams_chat_id', '')
            if chat_id:
                link = f"https://teams.cloud.microsoft/l/chat/19:{chat_id}@thread.v2/conversations?context=%7B%22contextType%22%3A%22chat%22%7D"
                webbrowser.open(link)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao abrir chat:\n{str(e)}")

    def open_teams_link_planilha(self, item):
        try:
            link = self.lab_config.get('sharepoint_planilha', '')
            if link:
                webbrowser.open(link)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao abrir link de diretório:\n\n{str(e)}")

    def extract_code_from_notes(self, notes):
        if not notes:
            return ""
        parts = notes.split(' - ')
        if len(parts) >= 2:
            return parts[1]
        return ""

    def extract_os_from_notes(self, notes):
        if not notes:
            return ""
        parts = notes.split(' - ')
        if len(parts) >= 1:
            return parts[0].replace("OS ", "").strip()
        return ""

    def extract_item_from_notes(self, notes):
        if not notes:
            return notes
        parts = notes.split(' - ')
        if len(parts) >= 2:
            return parts[0], parts[1], parts[2]
        return notes[:15]

    def load_exceptions_for_date(self, check_date):
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            rs = win32com.client.Dispatch("ADODB.Recordset")
            conn.Open(self.str_conn)
            
            date_str = check_date.strftime('%Y-%m-%d') if isinstance(check_date, date) else str(check_date)
            
            sql = f"""
                SELECT exception_type, start_time, end_time, notes
                FROM [IST-PGE].dbo.Calendar_Exceptions
                WHERE exception_date = '{date_str}'
                AND is_available = 0
                AND id_sector = {self.id_sector}
            """
            rs.Open(sql, conn)
            
            exceptions = []
            if not rs.EOF:
                rs.MoveFirst()
                while not rs.EOF:
                    start_t = rs.Fields('start_time').Value
                    end_t = rs.Fields('end_time').Value
                    notes = rs.Fields('notes').Value if rs.Fields('notes').Value else ""
                    
                    if start_t and isinstance(start_t, str):
                        start_t = str(start_t)[:5]
                    if end_t and isinstance(end_t, str):
                        end_t = str(end_t)[:5]
                    
                    if start_t and end_t:
                        exceptions.append(f"{start_t}-{end_t}: {notes}")
                    else:
                        exceptions.append(f"08:00-17:00: {notes}")
                    rs.MoveNext()
            
            rs.Close()
            conn.Close()
            return exceptions
        except Exception as e:
            print(f"Error loading exceptions: {e}")
            return []

    def load_exceptions_for_date_raw(self, check_date):
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            rs = win32com.client.Dispatch("ADODB.Recordset")
            conn.Open(self.str_conn)
            
            date_str = check_date.strftime('%Y-%m-%d') if isinstance(check_date, date) else str(check_date)
            
            sql = f"""
                SELECT id, exception_type, start_time, end_time, notes
                FROM [IST-PGE].dbo.Calendar_Exceptions
                WHERE exception_date = '{date_str}'
                AND is_available = 0
                AND id_sector = {self.id_sector}
            """
            rs.Open(sql, conn)
            
            exceptions = []
            if not rs.EOF:
                rs.MoveFirst()
                while not rs.EOF:
                    start_t = rs.Fields('start_time').Value
                    end_t = rs.Fields('end_time').Value
                    exceptions.append({
                        'id': rs.Fields('id').Value,
                        'exception_type': rs.Fields('exception_type').Value,
                        'start_time': str(start_t)[:5] if start_t else None,
                        'end_time': str(end_t)[:5] if end_t else None,
                        'notes': rs.Fields('notes').Value if rs.Fields('notes').Value else ""
                    })
                    rs.MoveNext()
            
            rs.Close()
            conn.Close()
            return exceptions
        except Exception as e:
            print(f"Error loading exceptions: {e}")
            return []

    def _get_exceptions(self):
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            rs = win32com.client.Dispatch("ADODB.Recordset")
            conn.Open(self.str_conn)
            
            sql = f"""
                SELECT exception_date, exception_type, start_time, end_time, is_available
                FROM [IST-PGE].dbo.Calendar_Exceptions
                WHERE id_sector = {self.id_sector}
            """
            rs.Open(sql, conn)
            
            result = []
            if not rs.EOF:
                rs.MoveFirst()
                while not rs.EOF:
                    exc_date = rs.Fields('exception_date').Value
                    if isinstance(exc_date, str):
                        exc_date = datetime.strptime(exc_date, '%Y-%m-%d').date()
                    elif isinstance(exc_date, datetime):
                        exc_date = exc_date.date()
                    
                    start_t = rs.Fields('start_time').Value
                    end_t = rs.Fields('end_time').Value
                    
                    if start_t and isinstance(start_t, str):
                        parts = start_t.split(':')
                        start_t = time(int(parts[0]), int(parts[1]))
                    if end_t and isinstance(end_t, str):
                        parts = end_t.split(':')
                        end_t = time(int(parts[0]), int(parts[1]))
                    
                    result.append({
                        'exception_date': exc_date,
                        'exception_type': rs.Fields('exception_type').Value,
                        'start_time': start_t,
                        'end_time': end_t,
                        'is_available': rs.Fields('is_available').Value
                    })
                    rs.MoveNext()
            
            rs.Close()
            conn.Close()
            return result
        except Exception as e:
            print(f"Error loading exceptions: {e}")
            return []

    def load_calendar_data(self):
        self.calendar_data.clear()
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            conn.Open(self.str_conn)
            
            sql = f"""
                SELECT 
                    ts.slot_date, ts.start_time, ts.end_time,
                    s.code, s.specification, s.description,
                    s.execution_time_minutes, ss.notes, ss.status,
                    ss.priority, ss.id as schedule_id
                FROM [IST-PGE].dbo.Time_Slots ts
                JOIN [IST-PGE].dbo.Service_Schedule ss ON ts.schedule_id = ss.id
                JOIN [IST-PGE].dbo.Service_Modes_Local s ON ts.service_id = s.id
                WHERE ts.id_sector = {self.id_sector}
                ORDER BY ts.slot_date, ts.start_time
            """
            
            rs = win32com.client.Dispatch("ADODB.Recordset")
            rs.Open(sql, conn)
            
            if not rs.EOF:
                rs.MoveFirst()
                while not rs.EOF:
                    slot_date = rs.Fields('slot_date').Value
                    if isinstance(slot_date, str):
                        slot_date = datetime.strptime(slot_date, '%Y-%m-%d').date()
                    elif isinstance(slot_date, datetime):
                        slot_date = slot_date.date()
                    
                    calendar_key = (slot_date.day, slot_date.month, slot_date.year)
                    
                    order_data = {
                        'start_time': rs.Fields('start_time').Value,
                        'end_time': rs.Fields('end_time').Value,
                        'code': rs.Fields('code').Value,
                        'specification': rs.Fields('specification').Value,
                        'description': rs.Fields('description').Value,
                        'execution_time_minutes': rs.Fields('execution_time_minutes').Value,
                        'notes': rs.Fields('notes').Value,
                        'status': rs.Fields('status').Value,
                        'priority': rs.Fields('priority').Value,
                        'schedule_id': rs.Fields('schedule_id').Value
                    }
                    
                    if calendar_key not in self.calendar_data:
                        self.calendar_data[calendar_key] = []
                    self.calendar_data[calendar_key].append(order_data)
                    rs.MoveNext()
            
            rs.Close()
            conn.Close()
        except Exception as e:
            print(f"Error loading calendar data: {e}")

    def render_calendar(self):
        for widget in self.calendar_frame.winfo_children():
            widget.destroy()
        
        today = datetime.now()
        target_date = today + relativedelta(months=self.current_month_offset)
        target_date = target_date.replace(day=1)
        year, month = target_date.year, target_date.month
        
        meses_pt = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
                    'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
        
        # Preload exceptions
        exceptions_by_date = {}
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            rs = win32com.client.Dispatch("ADODB.Recordset")
            conn.Open(self.str_conn)
            
            first_day = f"{year}-{month:02d}-01"
            if month == 12:
                last_day = f"{year+1}-01-01"
            else:
                last_day = f"{year}-{month+1:02d}-01"
            
            sql = f"""
                SELECT exception_date, exception_type, start_time, end_time, notes, is_available
                FROM [IST-PGE].dbo.Calendar_Exceptions
                WHERE exception_date >= '{first_day}'
                AND exception_date < '{last_day}'
                AND id_sector = {self.id_sector}
            """
            rs.Open(sql, conn)
            
            if not rs.EOF:
                rs.MoveFirst()
                while not rs.EOF:
                    exc_date = rs.Fields('exception_date').Value
                    if isinstance(exc_date, str):
                        exc_date = datetime.strptime(exc_date, '%Y-%m-%d').date()
                    elif isinstance(exc_date, datetime):
                        exc_date = exc_date.date()
                    
                    exc_key = (exc_date.day, exc_date.month, exc_date.year)
                    
                    start_t = rs.Fields('start_time').Value
                    end_t = rs.Fields('end_time').Value
                    is_avail = rs.Fields('is_available').Value
                    notes = rs.Fields('notes').Value if rs.Fields('notes').Value else ""
                    
                    if exc_key not in exceptions_by_date:
                        exceptions_by_date[exc_key] = []
                    
                    exceptions_by_date[exc_key].append({
                        'start_time': str(start_t)[:5] if start_t else None,
                        'end_time': str(end_t)[:5] if end_t else None,
                        'is_available': is_avail,
                        'notes': notes
                    })
                    rs.MoveNext()
            rs.Close()
            conn.Close()
        except Exception as e:
            print(f"Error loading exceptions: {e}")
        
        # Navigation Header
        nav_frame = tk.Frame(self.calendar_frame, bg=self.cor_fundo)
        nav_frame.pack(fill='x', pady=(0, 5))
        
        self.greetings = ttk.Label(nav_frame, text=f"{buenas}", font=("Segoe UI", 10), background='black')
        self.greetings.pack(side="left", padx=5, pady=(0, 0))
        
        nav_row = tk.Frame(nav_frame, bg=self.cor_fundo)
        nav_row.pack(fill='x')

        btn_prev = tk.Label(nav_row, text="◀", anchor='center', justify='center',font=('Segoe UI', 12, 'bold'), bg=self.cor_fundo, fg='white',cursor='hand2', padx=0)
        btn_prev.pack(side='left', padx=5)
        btn_prev.bind('<Button-1>', lambda e: self.change_month(-1))

        month_label = tk.Label(nav_row, text=f"{meses_pt[month-1]} {year}",font=('Segoe UI', 12, 'bold'), bg=self.cor_fundo, fg='white')
        month_label.pack(side='left', padx=5)

        btn_next = tk.Label(nav_row, text="▶", anchor='center', justify='center',font=('Segoe UI', 12, 'bold'), bg=self.cor_fundo, fg='white',cursor='hand2', padx=0)
        btn_next.pack(side='left', padx=5)
        btn_next.bind('<Button-1>', lambda e: self.change_month(1))

        nav_row2 = tk.Frame(nav_frame, bg=self.cor_fundo)
        nav_row2.pack(fill='x')

        lab_label = tk.Label(nav_row2, text=f"Calendário de Serviços - Laboratório de {self.lab_name}",font=('Segoe UI', 12, 'bold'), bg=self.cor_fundo, fg='white',anchor='w')
        lab_label.pack(side='left', padx=5)

        header_frame = tk.Frame(self.calendar_frame, bg=self.cor_fundo)
        header_frame.pack(fill='x')
        
        dias_semana = ['Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb', 'Dom']
        for dia in dias_semana:
            lbl = tk.Label(header_frame, text=dia, font=('Segoe UI', 9, 'bold'),bg='#1B4B9F', fg='white', width=16, pady=4)
            lbl.pack(side='left', padx=1, pady=1)
        
        # Calendar Grid
        grid_frame = tk.Frame(self.calendar_frame, bg=self.cor_fundo)
        grid_frame.pack(fill='both', expand=True)
        
        cal_matrix = cal.monthcalendar(year, month)
        today_date = today.date()
        
        for week in cal_matrix:
            week_frame = tk.Frame(grid_frame, bg=self.cor_fundo)
            week_frame.pack(fill='x')
            
            for day in week:
                if day == 0:
                    day_frame = tk.Frame(week_frame, bg='#908F8F', width=120, height=80)
                    day_frame.pack(side='left', padx=1, pady=1)
                    day_frame.pack_propagate(False)
                else:
                    current_date = datetime(year, month, day).date()
                    calendar_key = (day, month, year)
                    
                    is_today = (current_date == today_date)
                    is_weekend = current_date.weekday() >= 5
                    has_orders = calendar_key in self.calendar_data
                    has_exceptions = calendar_key in exceptions_by_date
                    
                    if is_today:
                        bg_color = '#B7D5F5'
                    elif is_weekend:
                        bg_color = "#908F8F"
                    elif has_orders:
                        bg_color = '#FFFFFF'
                    else:
                        bg_color = '#908F8F'
                    
                    is_clickable = not is_weekend
                    
                    if is_clickable:
                        day_frame = tk.Frame(week_frame, bg=bg_color, width=120, height=80, bd=1, cursor='hand2')
                    else:
                        day_frame = tk.Frame(week_frame, bg=bg_color, width=120, height=80, bd=1)
                    
                    day_frame.pack(side='left', padx=1, pady=1)
                    day_frame.pack_propagate(False)
                    
                    day_num_frame = tk.Frame(day_frame, bg=bg_color)
                    day_num_frame.pack(fill='x', padx=2, pady=1)
                    
                    fg_day = 'white' if is_today else 'black'
                    day_label = tk.Label(day_num_frame, text=str(day), font=('Segoe UI', 9, 'bold'),bg=bg_color, fg=fg_day, anchor='w')
                    day_label.pack(side='left')
                    
                    if has_orders:
                        count = len(self.calendar_data[calendar_key])
                        badge = tk.Label(day_num_frame, text=str(count), font=('Segoe UI', 8, 'bold'),bg='#2462D2', fg='white', width=3, height=1)
                        badge.pack(side='right')
                    
                    if has_exceptions:
                        exc_list = exceptions_by_date[calendar_key]
                        has_real_exception = any('intervalinho' not in str(exc.get('notes', '')).lower() for exc in exc_list)
                        if has_real_exception:
                            exc_indicator = tk.Label(day_num_frame, text="⚠", font=('Segoe UI', 8),bg=bg_color, fg="#E64848")
                            exc_indicator.pack(side='right', padx=(0, 2))

                    if has_orders:
                        unique_items = []
                        seen = set()
                        for order in self.calendar_data[calendar_key]:
                            result = self.extract_item_from_notes(order.get('notes', ''))
                            if isinstance(result, tuple) and len(result) >= 2:
                                item_code = result[0]
                            else:
                                item_code = str(result)[:15] if result else ""
                            
                            if item_code and item_code not in seen:
                                unique_items.append(item_code)
                                seen.add(item_code)
                        
                        preview_text = "\n".join(unique_items[:3])
                        preview = tk.Label(day_frame, text=preview_text, font=('Segoe UI', 7), bg=bg_color,fg='#333333', anchor='w', justify='left', cursor='hand2')
                        preview.pack(fill='x', padx=3)

                        if len(unique_items) > 3:
                            more_label = tk.Label(day_frame, text=f"+{len(unique_items) - 3} mais...",font=('Segoe UI', 6), bg=bg_color, fg="#4A4949", anchor='w')
                            more_label.pack(fill='x', padx=3)
                    
                    if is_clickable:
                        day_frame.bind('<Button-1>', lambda e, key=calendar_key: self.show_day_orders(key))
                        day_frame.bind('<Button-3>', lambda e, key=calendar_key: self.show_day_context_menu(e, key))
                        for child in day_frame.winfo_children():
                            child.bind('<Button-1>', lambda e, key=calendar_key: self.show_day_orders(key))
                            child.bind('<Button-3>', lambda e, key=calendar_key: self.show_day_context_menu(e, key))
        
        btn_today = ttk.Button(grid_frame, text="Hoje", command=self.go_to_today, cursor='hand2')
        btn_today.pack(side="left", padx=5)
        
        btn_refresh = ttk.Button(grid_frame, text="Atualizar", command=self.refresh_calendar, cursor='hand2')
        btn_refresh.pack(side="left", padx=5)
        
        self.buscar = ttk.Button(grid_frame, text=" Limpar agenda ", command=self.clear_all_scheduled, cursor='hand2')
        self.buscar.pack(side="left", padx=5)

    def show_day_context_menu(self, event, calendar_key):
        context_menu = tk.Menu(self.calendar_frame, tearoff=0)
        
        count = len(self.calendar_data.get(calendar_key, []))
        
        if count == 1:
            context_menu.add_command(label=f"Ver serviço", command=lambda k=calendar_key: self.show_day_orders(k))
        elif count > 1:
            context_menu.add_command(label=f"Ver serviços", command=lambda k=calendar_key: self.show_day_orders(k))
        
        context_menu.add_separator()
        context_menu.add_command(label=f"Iniciar chat do Teams com {self.lab_name}", command=self.open_teams_chat)
        
        exc_date = datetime(calendar_key[2], calendar_key[1], calendar_key[0]).date()
        exceptions = self.load_exceptions_for_date(exc_date)
        
        if exceptions:
            context_menu.add_separator()
            if len(exceptions) == 1:
                context_menu.add_command(label=f"Remover bloqueio de agenda deste dia", command=lambda k=calendar_key: self.remove_day_exceptions(k))
            if len(exceptions) > 1:
                context_menu.add_command(label=f"Remover bloqueios de agenda deste dia", command=lambda k=calendar_key: self.remove_day_exceptions(k))
        
        context_menu.add_separator()
        if count == 1:
            context_menu.add_command(label=f"Reagendar o serviço para um dia específico...", command=lambda k=calendar_key: self.reschedule_day_services_manual(k))
            context_menu.add_command(label=f"Reagendar o serviço automaticamente", command=lambda k=calendar_key: self.reschedule_day_services(k))
        elif count > 1:
            context_menu.add_command(label=f"Reagendar todos os serviços para um dia específico...", command=lambda k=calendar_key: self.reschedule_day_services_manual(k))
            context_menu.add_command(label=f"Reagendar todos os serviços automaticamente", command=lambda k=calendar_key: self.reschedule_day_services(k))
        
        context_menu.post(event.x_root, event.y_root)

    # ===== SCHEDULING ENGINE =====
    def find_next_available_slot(self, from_time, duration_minutes):
        check_date = from_time.date()
        max_days = 365
        
        availability = self._get_availability()
        exceptions = self._get_exceptions()
        
        for _ in range(max_days):
            sql_day_of_week = (check_date.weekday() + 1) % 7
            day_windows = [w for w in availability if w['day_of_week'] == sql_day_of_week]
            
            if not day_windows:
                check_date += timedelta(days=1)
                from_time = datetime.combine(check_date, time(8, 0))
                continue
            
            full_day_blocked = any(
                ex['exception_date'] == check_date and ex['is_available'] == False and ex['start_time'] is None
                for ex in exceptions
            )
            
            if full_day_blocked:
                check_date += timedelta(days=1)
                from_time = datetime.combine(check_date, time(8, 0))
                continue
            
            for window in sorted(day_windows, key=lambda w: w['start_time']):
                window_start_dt = datetime.combine(check_date, window['start_time'])
                window_end_dt = datetime.combine(check_date, window['end_time'])
                proposed_start = max(from_time, window_start_dt)
                
                while proposed_start + timedelta(minutes=duration_minutes) <= window_end_dt:
                    proposed_end = proposed_start + timedelta(minutes=duration_minutes)
                    
                    blocked_by_exception = False
                    for ex in exceptions:
                        if (ex['exception_date'] == check_date and ex['is_available'] == False
                            and ex['start_time'] is not None and ex['end_time'] is not None):
                            ex_start_dt = datetime.combine(check_date, ex['start_time'])
                            ex_end_dt = datetime.combine(check_date, ex['end_time'])
                            if proposed_start < ex_end_dt and proposed_end > ex_start_dt:
                                blocked_by_exception = True
                                proposed_start = ex_end_dt
                                break
                    
                    if blocked_by_exception:
                        continue
                    
                    if self._has_time_slot_conflict(check_date, proposed_start.time(), proposed_end.time()):
                        next_free = self._get_next_free_time(check_date, proposed_start.time())
                        if next_free is None:
                            break
                        proposed_start = datetime.combine(check_date, next_free)
                        continue
                    
                    return proposed_start, proposed_end
            
            check_date += timedelta(days=1)
            if day_windows:
                earliest = min(w['start_time'] for w in day_windows)
                from_time = datetime.combine(check_date, earliest)
            else:
                from_time = datetime.combine(check_date, time(8, 0))
        
        return None, None

    def _get_availability(self):
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            rs = win32com.client.Dispatch("ADODB.Recordset")
            conn.Open(self.str_conn)
            
            sql = """
                SELECT day_of_week, start_time, end_time
                FROM [IST-PGE].dbo.Calendar_Availability
                WHERE is_active = 1
                ORDER BY day_of_week, start_time
            """
            rs.Open(sql, conn)
            
            result = []
            if not rs.EOF:
                rs.MoveFirst()
                while not rs.EOF:
                    start_t = rs.Fields('start_time').Value
                    end_t = rs.Fields('end_time').Value
                    if isinstance(start_t, str):
                        parts = start_t.split(':')
                        start_t = time(int(parts[0]), int(parts[1]))
                    if isinstance(end_t, str):
                        parts = end_t.split(':')
                        end_t = time(int(parts[0]), int(parts[1]))
                    
                    result.append({
                        'day_of_week': rs.Fields('day_of_week').Value,
                        'start_time': start_t,
                        'end_time': end_t
                    })
                    rs.MoveNext()
            
            rs.Close()
            conn.Close()
            return result
        except Exception as e:
            print(f"Error loading availability: {e}")
            return []

    def _has_time_slot_conflict(self, check_date, start_time, end_time):
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            rs = win32com.client.Dispatch("ADODB.Recordset")
            conn.Open(self.str_conn)
            
            start_str = start_time.strftime('%H:%M:%S') if isinstance(start_time, time) else str(start_time)
            end_str = end_time.strftime('%H:%M:%S') if isinstance(end_time, time) else str(end_time)
            date_str = check_date.strftime('%Y-%m-%d')
            
            sql = f"""
                SELECT COUNT(*) as cnt
                FROM [IST-PGE].dbo.Time_Slots
                WHERE slot_date = '{date_str}'
                AND start_time < '{end_str}'
                AND end_time > '{start_str}'
                AND id_sector = {self.id_sector}
            """
            rs.Open(sql, conn)
            count = rs.Fields('cnt').Value if not rs.EOF else 0
            rs.Close()
            conn.Close()
            return count > 0
        except Exception as e:
            print(f"Error checking conflicts: {e}")
            return True

    def _get_next_free_time(self, check_date, after_time):
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            rs = win32com.client.Dispatch("ADODB.Recordset")
            conn.Open(self.str_conn)
            
            date_str = check_date.strftime('%Y-%m-%d')
            time_str = after_time.strftime('%H:%M:%S') if isinstance(after_time, time) else str(after_time)
            
            sql = f"""
                SELECT TOP 1 end_time
                FROM [IST-PGE].dbo.Time_Slots
                WHERE slot_date = '{date_str}'
                AND start_time <= '{time_str}'
                AND end_time > '{time_str}'
                AND id_sector = {self.id_sector}
                ORDER BY end_time ASC
            """
            rs.Open(sql, conn)
            
            if not rs.EOF:
                end_time = rs.Fields('end_time').Value
                result = self._parse_time_value(end_time)
                rs.Close()
                conn.Close()
                return result
            
            rs.Close()
            
            sql2 = f"""
                SELECT TOP 1 end_time
                FROM [IST-PGE].dbo.Time_Slots
                WHERE slot_date = '{date_str}'
                AND start_time > '{time_str}'
                AND id_sector = {self.id_sector}
                ORDER BY start_time ASC
            """
            rs.Open(sql2, conn)
            
            if not rs.EOF:
                end_time = rs.Fields('end_time').Value
                result = self._parse_time_value(end_time)
                rs.Close()
                conn.Close()
                return result
            
            rs.Close()
            conn.Close()
            return None
        except Exception as e:
            print(f"Error getting next free time: {e}")
            return None

    def _parse_time_value(self, value):
        if value is None:
            return None
        if isinstance(value, time):
            return value
        if isinstance(value, str):
            value = value.split('.')[0]
            if ':' in value:
                parts = value.split(':')
                return time(int(parts[0]), int(parts[1]), 0)
            else:
                return time(int(value), 0, 0)
        if isinstance(value, float):
            total_hours = value * 24
            hours = int(total_hours)
            minutes = int((total_hours - hours) * 60)
            return time(hours, minutes, 0)
        if isinstance(value, datetime):
            return time(value.hour, value.minute, 0)
        return None

    def schedule_all_pending(self, from_date=None):
        
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            conn.Open(self.str_conn)
            
            rs = win32com.client.Dispatch("ADODB.Recordset")
            sql = f"""
                SELECT ss.id as schedule_id, ss.service_id,
                       s.execution_time_minutes, ss.priority
                FROM [IST-PGE].dbo.Service_Schedule ss
                JOIN [IST-PGE].dbo.Service_Modes_Local s ON ss.service_id = s.id
                WHERE ss.status = 'PENDING' AND ss.id_sector = {self.id_sector}
                ORDER BY ss.priority ASC, ss.requested_at ASC
            """
            rs.Open(sql, conn)
            
            pending_services = []
            if not rs.EOF:
                rs.MoveFirst()
                while not rs.EOF:
                    pending_services.append({
                        'schedule_id': rs.Fields('schedule_id').Value,
                        'service_id': rs.Fields('service_id').Value,
                        'execution_time_minutes': rs.Fields('execution_time_minutes').Value,
                        'priority': rs.Fields('priority').Value
                    })
                    rs.MoveNext()
            rs.Close()
            
            if len(pending_services) == 0:
                conn.Close()
                return None, None
            
            if from_date:
                current_time = from_date if isinstance(from_date, datetime) else datetime.combine(from_date, time(8, 0))
            else:
                current_time = datetime.now()
            
            first_start = None
            last_end = None
            
            for service in pending_services:
                next_start, next_end = self.find_next_available_slot(current_time, service['execution_time_minutes'])
                
                if next_start is None:
                    continue
                
                if first_start is None:
                    first_start = next_start
                last_end = next_end
                
                conn.Execute(f"""
                    UPDATE [IST-PGE].dbo.Service_Schedule
                    SET scheduled_start = '{next_start.strftime('%Y-%m-%d %H:%M:%S')}',
                        scheduled_end = '{next_end.strftime('%Y-%m-%d %H:%M:%S')}',
                        status = 'SCHEDULED', updated_at = GETDATE()
                    WHERE id = {service['schedule_id']}
                """)
                
                conn.Execute(f"""
                    INSERT INTO [IST-PGE].dbo.Time_Slots (slot_date, start_time, end_time, schedule_id, service_id, id_sector)
                    VALUES ('{next_start.strftime('%Y-%m-%d')}', '{next_start.strftime('%H:%M:%S')}',
                            '{next_end.strftime('%H:%M:%S')}', {service['schedule_id']},
                            {service['service_id']}, {self.id_sector})
                """)
                
                current_time = next_end
            
            conn.Close()
            
            return first_start, last_end
        except Exception as e:
            print(f"Error in schedule_all_pending: {e}")
            return None, None

    # ===== RESCHEDULE METHODS =====
    def reschedule_day_services(self, calendar_key):
        if calendar_key not in self.calendar_data:
            messagebox.showinfo("Aviso", "Nenhum serviço neste dia!")
            return
        
        count = len(self.calendar_data[calendar_key])
        date_str = f"{calendar_key[0]:02d}/{calendar_key[1]:02d}/{calendar_key[2]}"
        confirm = messagebox.askyesno("Info agendamento", f"Deseja reagendar {'o serviço' if count == 1 else f'os {count} serviços'} do dia {date_str}?")
        if not confirm:
            return
        
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            conn.Open(self.str_conn)
            
            reschedule_date = datetime(calendar_key[2], calendar_key[1], calendar_key[0]).date()
            date_str_sql = reschedule_date.strftime('%Y-%m-%d')
            
            rs = win32com.client.Dispatch("ADODB.Recordset")
            rs.Open(f"SELECT ts.schedule_id FROM [IST-PGE].dbo.Time_Slots ts WHERE ts.slot_date = '{date_str_sql}' AND ts.id_sector = {self.id_sector}", conn)
            
            schedule_ids = []
            if not rs.EOF:
                rs.MoveFirst()
                while not rs.EOF:
                    schedule_ids.append(rs.Fields('schedule_id').Value)
                    rs.MoveNext()
            rs.Close()
            
            if not schedule_ids:
                conn.Close()
                return
            
            conn.Execute(f"DELETE FROM [IST-PGE].dbo.Time_Slots WHERE slot_date = '{date_str_sql}' AND id_sector = {self.id_sector}")
            
            id_list = ','.join(str(sid) for sid in schedule_ids)
            conn.Execute(f"""
                UPDATE [IST-PGE].dbo.Service_Schedule
                SET status = 'PENDING', scheduled_start = NULL, scheduled_end = NULL, updated_at = GETDATE()
                WHERE id IN ({id_list})
            """)
            conn.Close()
            
            self.show_loading_screen(f"Carregando {len(schedule_ids)} {"itens..." if len(schedule_ids) > 1 else "item..."}")

            self.next_start, self.next_end = self.schedule_all_pending(from_date=reschedule_date + timedelta(days=1))
            
            self.load_calendar_data()
            self.refresh_calendar()
            self.update_queue_count()
            
            self.hide_loading_screen()
            
            if count == 1:
                msg = f"Serviço reagendado para {self.next_start.strftime('%d/%m/%Y')} das {self.next_start.strftime('%H:%M')} às {self.next_end.strftime('%H:%M')}"
            else:
                msg = f"{count} serviços reagendados para a partir de {self.next_start.strftime('%d/%m/%Y')}"
            self.status_label.config(text=msg)
            messagebox.showinfo('Info agendamento', msg)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao reagendar:\n{str(e)}")

    def reschedule_day_services_manual(self, calendar_key):
        if calendar_key not in self.calendar_data:
            messagebox.showinfo("Aviso", "Nenhum serviço neste dia!")
            return
        
        count = len(self.calendar_data[calendar_key])
        date_str = f"{calendar_key[0]:02d}/{calendar_key[1]:02d}/{calendar_key[2]}"
        confirm = messagebox.askyesno("Info agendamento", 
            f"Deseja reagendar manualmente {'o serviço' if count == 1 else f'os {count} serviços'} do dia {date_str}?")
        if not confirm:
            return
        
        self._open_manual_reschedule_popup(calendar_key)

    def _open_manual_reschedule_popup(self, calendar_key):
        popup = tk.Toplevel(self.frame_certificados)
        popup.geometry("380x280")
        popup.configure(bg=self.cor_fundo)
        popup.title("Reagendamento Manual")
        
        count = len(self.calendar_data[calendar_key])
        date_str = f"{calendar_key[0]:02d}/{calendar_key[1]:02d}/{calendar_key[2]}"
        
        tk.Label(popup, text=f"Reagendar {count} serviço(s) de {date_str}",
                font=('Segoe UI', 10, 'bold'), bg=self.cor_fundo, fg='white').pack(pady=(15, 10))
        
        tk.Label(popup, text="Nova data (DD/MM/AAAA):", font=('Segoe UI', 9),
                bg=self.cor_fundo, fg='white').pack(anchor='w', padx=30, pady=(10, 0))
        
        date_frame = tk.Frame(popup, bg=self.cor_fundo)
        date_frame.pack(fill='x', padx=30, pady=(2, 5))
        
        tomorrow = datetime.now() + timedelta(days=1)
        day_var = tk.StringVar(value=tomorrow.strftime('%d'))
        month_var = tk.StringVar(value=tomorrow.strftime('%m'))
        year_var = tk.StringVar(value=tomorrow.strftime('%Y'))
        
        ttk.Entry(date_frame, textvariable=day_var, width=3).pack(side='left')
        tk.Label(date_frame, text="/", bg=self.cor_fundo, fg='white').pack(side='left')
        ttk.Entry(date_frame, textvariable=month_var, width=3).pack(side='left')
        tk.Label(date_frame, text="/", bg=self.cor_fundo, fg='white').pack(side='left')
        ttk.Entry(date_frame, textvariable=year_var, width=5).pack(side='left')
        
        tk.Label(popup, text="Horário de início (HH:MM):", font=('Segoe UI', 9),
                bg=self.cor_fundo, fg='white').pack(anchor='w', padx=30, pady=(10, 0))
        
        time_var = tk.StringVar(value="08:00")
        ttk.Entry(popup, textvariable=time_var, width=10).pack(anchor='w', padx=30, pady=(2, 5))
        
        tk.Label(popup, text="⚠ Os serviços serão agendados em sequência\n   a partir desta data e horário.",
                font=('Segoe UI', 8), bg=self.cor_fundo, fg='#FFD700').pack(pady=(10, 5))
        
        btn_frame = tk.Frame(popup, bg=self.cor_fundo)
        btn_frame.pack(pady=15)
        
        ttk.Button(btn_frame, text=" Confirmar Reagendamento ",
                command=lambda: self._execute_manual_reschedule(calendar_key, popup, day_var, month_var, year_var, time_var)
                ).pack(side='left', padx=5)
        ttk.Button(btn_frame, text=" Cancelar ", command=popup.destroy).pack(side='left', padx=5)

    def _execute_manual_reschedule(self, calendar_key, popup, day_var, month_var, year_var, time_var):
        try:
            day = day_var.get().strip().zfill(2)
            month = month_var.get().strip().zfill(2)
            year = year_var.get().strip()
            new_date = datetime.strptime(f"{year}-{month}-{day}", '%Y-%m-%d').date()
            
            time_str = time_var.get().strip()
            parts = time_str.split(':')
            if len(parts) != 2:
                messagebox.showwarning("Aviso", "Formato de horário inválido!")
                return
            new_time = time(int(parts[0]), int(parts[1]))
            new_datetime = datetime.combine(new_date, new_time)
            
            reschedule_date_orig = datetime(calendar_key[2], calendar_key[1], calendar_key[0]).date()
            if new_date == reschedule_date_orig:
                messagebox.showwarning("Aviso", f"A nova data deve ser diferente da data do agendamento ({reschedule_date_orig.strftime('%d/%m/%Y')})!")
                return
            #if new_date <= reschedule_date_orig:
            #    messagebox.showwarning("Aviso", f"A nova data deve ser posterior a {reschedule_date_orig.strftime('%d/%m/%Y')}!")
            #    return
            
            popup.destroy()
            
            count = len(self.calendar_data[calendar_key])
            
            conn = win32com.client.Dispatch("ADODB.Connection")
            conn.Open(self.str_conn)
            
            reschedule_date = datetime(calendar_key[2], calendar_key[1], calendar_key[0]).date()
            date_str_sql = reschedule_date.strftime('%Y-%m-%d')
            
            rs = win32com.client.Dispatch("ADODB.Recordset")
            rs.Open(f"SELECT ts.schedule_id FROM [IST-PGE].dbo.Time_Slots ts WHERE ts.slot_date = '{date_str_sql}' AND ts.id_sector = {self.id_sector}", conn)
            
            schedule_ids = []
            if not rs.EOF:
                rs.MoveFirst()
                while not rs.EOF:
                    schedule_ids.append(rs.Fields('schedule_id').Value)
                    rs.MoveNext()
            rs.Close()
            
            if not schedule_ids:
                conn.Close()
                return
            
            conn.Execute(f"DELETE FROM [IST-PGE].dbo.Time_Slots WHERE slot_date = '{date_str_sql}' AND id_sector = {self.id_sector}")
            
            id_list = ','.join(str(sid) for sid in schedule_ids)
            conn.Execute(f"UPDATE [IST-PGE].dbo.Service_Schedule SET status = 'PENDING', scheduled_start = NULL, scheduled_end = NULL, updated_at = GETDATE() WHERE id IN ({id_list})")
            conn.Close()
            
            self.show_loading_screen(f"Carregando {len(schedule_ids)} itens...")

            self.next_start, self.next_end = self.schedule_all_pending(from_date=new_datetime)
            
            self.load_calendar_data()
            self.refresh_calendar()
            self.update_queue_count()
            self.hide_loading_screen()
            
            new_date_br = new_date.strftime('%d/%m/%Y')
            messagebox.showinfo('Info agendamento', f"Serviço(s) reagendado(s) para {new_date_br} a partir das {new_time.strftime('%H:%M')}")
        except ValueError:
            messagebox.showwarning("Aviso", "Data ou horário inválido!")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao reagendar:\n{str(e)}")

    # ===== DAY ORDERS POPUP =====
    def show_day_orders(self, calendar_key):
        popup = tk.Toplevel(self.frame_certificados)
        popup.geometry("850x400")
        popup.minsize(800, 400)
        popup.configure(bg=self.cor_fundo)
        popup.calendar_key = calendar_key

        date_str = f"{calendar_key[0]:02d}/{calendar_key[1]:02d}/{calendar_key[2]}"
        popup.title(f"Serviços - {date_str} ({self.lab_name})")
        
        orders = self.calendar_data.get(calendar_key, [])
        popup.schedule_ids = [o.get('schedule_id') for o in orders]
        has_orders = len(orders) > 0
        
        if has_orders:
            tree_frame = ttk.Frame(popup)
            tree_frame.pack(pady=10, padx=10, expand=True, fill='both')
            
            tree_scroll = ttk.Scrollbar(tree_frame)
            tree_scroll.pack(side='right', fill='y')
            
            columns = ("Horário", "OS", "Item", "Serviço", "Descrição", "Duração", "Código")
            tree = ttk.Treeview(tree_frame, columns=columns, show="headings", yscrollcommand=tree_scroll.set, selectmode="extended")
            tree_scroll.config(command=tree.yview)
            
            for col, width in [("Horário", 100), ("OS", 80), ("Item", 70), ("Serviço", 150), ("Descrição", 150), ("Duração", 55), ("Código", 70)]:
                tree.heading(col, text=col)
                tree.column(col, width=width)
            
            tree.pack(expand=True, fill='both')
            tree.tag_configure('urgente', background='#FFD9D9')
            tree.tag_configure('normal', background='#FFFFFF')
            
            for order in orders:
                start_str = str(order['start_time'])[:5] if order['start_time'] else "--:--"
                end_str = str(order['end_time'])[:5] if order['end_time'] else "--:--"
                code = self.extract_code_from_notes(order.get('notes', ''))
                os_code = self.extract_os_from_notes(order.get('notes', ''))
                
                is_urgent = order.get('priority', 10) <= 5
                tag = 'urgente' if is_urgent else 'normal'
                
                tree.insert("", "end", values=(
                    f"{start_str} - {end_str}", os_code, code,
                    order['specification'], order['description'],
                    f"{order['execution_time_minutes']} min", order['code']
                ), tags=(tag,))
        else:
            tree = None
            msg_frame = tk.Frame(popup, bg=self.cor_fundo)
            msg_frame.pack(expand=True)
            tk.Label(msg_frame, text="Nenhum serviço agendado para este dia", font=('Segoe UI', 11), bg=self.cor_fundo, fg='#B7D5F5').pack()
        
        # Exceptions
        exc_date = datetime(calendar_key[2], calendar_key[1], calendar_key[0]).date()
        exceptions_raw = self.load_exceptions_for_date_raw(exc_date)
        
        if exceptions_raw:
            exc_frame = tk.Frame(popup, bg=self.cor_fundo)
            exc_frame.pack(fill='x', padx=10, pady=(0, 5))
            tk.Label(exc_frame, text="Bloqueios de Agenda:", font=('Segoe UI', 9, 'bold'), bg=self.cor_fundo, fg='#FFD700').pack(anchor='w')
            
            for i, exc in enumerate(exceptions_raw):
                exc_text = self._format_exception_text(exc)
                lbl = tk.Label(exc_frame, text=f"  ⛔ {exc_text}", font=('Segoe UI', 8), bg=self.cor_fundo, fg='#FF9999', cursor='hand2')
                lbl.pack(anchor='w')
                lbl.bind('<Button-3>', lambda e, idx=i: self._show_exception_context_menu(e, exc_frame, idx, exc_date, exceptions_raw, popup))
                lbl.bind('<Enter>', lambda e, l=lbl: l.configure(bg='#3A3A5C'))
                lbl.bind('<Leave>', lambda e, l=lbl: l.configure(bg=self.cor_fundo))
        
        btn_frame = tk.Frame(popup, bg=self.cor_fundo)
        btn_frame.pack(fill='x', padx=10, pady=(5, 10))
        ttk.Button(btn_frame, text=f"Iniciar chat do Teams com {self.lab_name}", command=self.open_teams_chat, cursor='hand2').pack(side='left', padx=5)
        
        if tree:
            tree.bind('<Button-3>', lambda e, t=tree: self.show_popup_context_menu(e, t))

    def _format_exception_text(self, exc):
        if exc['start_time'] and exc['end_time']:
            return f"{exc['start_time']}-{exc['end_time']}: {exc['notes']}"
        return f"Dia inteiro: {exc['notes']}"

    def _show_exception_context_menu(self, event, parent_frame, exc_index, exc_date, exceptions_raw, popup):
        exc = exceptions_raw[exc_index]
        if 'intervalinho' in str(exc.get('notes', '')).lower():
            return
        
        context_menu = tk.Menu(parent_frame, tearoff=0)
        exc_text = self._format_exception_text(exc)
        context_menu.add_command(label=f"Remover bloqueio '{exc_text[:50]}'", command=lambda: self._remove_single_exception(exc['id'], exc_date, popup))
        context_menu.post(event.x_root, event.y_root)

    def _remove_single_exception(self, exception_id, exc_date, popup):
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            conn.Open(self.str_conn)
            conn.Execute(f"DELETE FROM [IST-PGE].dbo.Calendar_Exceptions WHERE id = {exception_id}")
            conn.Close()
            
            self.load_calendar_data()
            self.render_calendar()
            
            if popup and popup.winfo_exists():
                calendar_key = getattr(popup, 'calendar_key', None)
                if calendar_key:
                    for widget in popup.winfo_children():
                        widget.destroy()
                    self._refresh_popup_content(popup, calendar_key)
            
            self.status_label.config(text="Bloqueio de agenda removido com sucesso!")
            messagebox.showinfo('Info agendamento', 'Bloqueio de agenda removido com sucesso!')
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao remover bloqueio:\n{str(e)}")

    def _refresh_popup_content(self, popup, calendar_key):
        popup.calendar_key = calendar_key
        date_str = f"{calendar_key[0]:02d}/{calendar_key[1]:02d}/{calendar_key[2]}"
        popup.title(f"Serviços - {date_str} ({self.lab_name})")
        
        orders = self.calendar_data.get(calendar_key, [])
        popup.schedule_ids = [o.get('schedule_id') for o in orders]
        has_orders = len(orders) > 0
        
        if has_orders:
            tree_frame = ttk.Frame(popup)
            tree_frame.pack(pady=10, padx=10, expand=True, fill='both')
            tree_scroll = ttk.Scrollbar(tree_frame)
            tree_scroll.pack(side='right', fill='y')
            
            columns = ("Horário", "OS", "Item", "Serviço", "Descrição", "Duração", "Código")
            tree = ttk.Treeview(tree_frame, columns=columns, show="headings", yscrollcommand=tree_scroll.set, selectmode="extended")
            tree_scroll.config(command=tree.yview)
            
            for col, width in [("Horário", 100), ("OS", 80), ("Item", 70), ("Serviço", 150), ("Descrição", 150), ("Duração", 55), ("Código", 70)]:
                tree.heading(col, text=col)
                tree.column(col, width=width)
            tree.pack(expand=True, fill='both')
            tree.tag_configure('urgente', background='#FFD9D9')
            tree.tag_configure('normal', background='#FFFFFF')
            
            for order in orders:
                start_str = str(order['start_time'])[:5] if order['start_time'] else "--:--"
                end_str = str(order['end_time'])[:5] if order['end_time'] else "--:--"
                code = self.extract_code_from_notes(order.get('notes', ''))
                os_code = self.extract_os_from_notes(order.get('notes', ''))
                tag = 'urgente' if order.get('priority', 10) <= 5 else 'normal'
                tree.insert("", "end", values=(f"{start_str} - {end_str}", os_code, code, order['specification'], order['description'], f"{order['execution_time_minutes']} min", order['code']), tags=(tag,))
        
        exc_date = datetime(calendar_key[2], calendar_key[1], calendar_key[0]).date()
        exceptions_raw = self.load_exceptions_for_date_raw(exc_date)
        
        if exceptions_raw:
            exc_frame = tk.Frame(popup, bg=self.cor_fundo)
            exc_frame.pack(fill='x', padx=10, pady=(0, 5))
            tk.Label(exc_frame, text="Bloqueios de Agenda:", font=('Segoe UI', 9, 'bold'), bg=self.cor_fundo, fg='#FFD700').pack(anchor='w')
            for i, exc in enumerate(exceptions_raw):
                exc_text = self._format_exception_text(exc)
                lbl = tk.Label(exc_frame, text=f"  ⛔ {exc_text}", font=('Segoe UI', 8), bg=self.cor_fundo, fg='#FF9999', cursor='hand2')
                lbl.pack(anchor='w')
                lbl.bind('<Button-3>', lambda e, idx=i: self._show_exception_context_menu(e, exc_frame, idx, exc_date, exceptions_raw, popup))
                lbl.bind('<Enter>', lambda e, l=lbl: l.configure(bg='#3A3A5C'))
                lbl.bind('<Leave>', lambda e, l=lbl: l.configure(bg=self.cor_fundo))
        
        btn_frame = tk.Frame(popup, bg=self.cor_fundo)
        btn_frame.pack(fill='x', padx=10, pady=(5, 10))
        ttk.Button(btn_frame, text=f"Iniciar chat do Teams com {self.lab_name}", command=self.open_teams_chat, cursor='hand2').pack(side='left', padx=5)
        
        if has_orders:
            tree.bind('<Button-3>', lambda e, t=tree: self.show_popup_context_menu(e, t))

    def show_popup_context_menu(self, event, tree):
        selected = tree.selection()
        if not selected:
            return
        
        popup = tree.winfo_toplevel()
        calendar_key = getattr(popup, 'calendar_key', None)
        total_count = len(self.calendar_data.get(calendar_key, []))
        selected_count = len(selected)
        
        context_menu = tk.Menu(tree, tearoff=0)
        values = tree.item(selected[0])['values']
        os_code = values[1] if len(values) > 1 else ""
        
        context_menu.add_command(label=f"Iniciar chat do Teams com {self.lab_name}", command=self.open_teams_chat)
        context_menu.add_separator()
        
        if os_code:
            sharepoint_url = self.get_sharepoint_url(os_code)
            if sharepoint_url:
                context_menu.add_command(label="Abrir link do SharePoint", command=lambda u=sharepoint_url: webbrowser.open(u))
                context_menu.add_command(label="Copiar link do SharePoint", command=lambda u=sharepoint_url: self.copy_to_clipboard(u) if hasattr(self, 'copy_to_clipboard') else None)
        
        context_menu.add_command(label="Abrir diretório da planilha de cálculo", command=lambda: self.open_teams_link_planilha(None))
        
        if calendar_key and total_count > 0:
            context_menu.add_separator()
            if selected_count == 1:
                context_menu.add_command(label="Reagendar serviço selecionado automaticamente", command=lambda t=tree, s=selected: self.reschedule_selected_services(t, s))
            else:
                context_menu.add_command(label=f"Reagendar {selected_count} serviços selecionados automaticamente", command=lambda t=tree, s=selected: self.reschedule_selected_services(t, s))
        
        context_menu.post(event.x_root, event.y_root)

    def reschedule_selected_services(self, tree, selected_items):
        popup = tree.winfo_toplevel()
        calendar_key = getattr(popup, 'calendar_key', None)
        if not calendar_key:
            return
        
        all_items = tree.get_children()
        selected_indices = [all_items.index(item) for item in selected_items]
        schedule_ids = getattr(popup, 'schedule_ids', [])
        selected_schedule_ids = [schedule_ids[i] for i in selected_indices if i < len(schedule_ids)]
        
        if not selected_schedule_ids:
            return
        
        count = len(selected_schedule_ids)
        date_str = f"{calendar_key[0]:02d}/{calendar_key[1]:02d}/{calendar_key[2]}"
        confirm = messagebox.askyesno("Info agendamento", f"Deseja reagendar {'o serviço selecionado' if count == 1 else f'os {count} serviços selecionados'} de {date_str}?")
        if not confirm:
            return
        
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            conn.Open(self.str_conn)
            
            id_list = ','.join(str(sid) for sid in selected_schedule_ids)
            conn.Execute(f"UPDATE [IST-PGE].dbo.Service_Schedule SET status = 'PENDING', scheduled_start = NULL, scheduled_end = NULL, updated_at = GETDATE() WHERE id IN ({id_list})")
            conn.Execute(f"DELETE FROM [IST-PGE].dbo.Time_Slots WHERE schedule_id IN ({id_list})")
            conn.Close()
            
            reschedule_date = datetime(calendar_key[2], calendar_key[1], calendar_key[0]).date()

            self.show_loading_screen(f"Carregando {len(selected_indices)} itens...")
            first_start, last_end = self.schedule_all_pending(from_date=reschedule_date + timedelta(days=1))
            
            self.load_calendar_data()

            self.render_calendar()
            popup.destroy()
            self.hide_loading_screen()
            
            #msg = f"Serviço reagendado para {self.next_start.strftime('%d/%m/%Y')}" if count == 1 else f"{count} serviços reagendados"
            msg = f"Serviço reagendado para {last_end.strftime('%d/%m/%Y')} às {last_end.strftime('%H:%M')}" if count == 1 else f"Serviços reagendados para a partir de {first_start.strftime('%d/%m/%Y')} às {first_start.strftime('%H:%M')}"
            self.status_label.config(text=msg)
            messagebox.showinfo('Info agendamento', msg)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao reagendar:\n{str(e)}")

    def remove_day_exceptions(self, calendar_key):
        date_str = f"{calendar_key[0]:02d}/{calendar_key[1]:02d}/{calendar_key[2]}"
        exc_date = datetime(calendar_key[2], calendar_key[1], calendar_key[0]).date()
        exceptions = self.load_exceptions_for_date(exc_date)
        count = len(exceptions)
        
        if count == 1:
            
            messagebox.showerror("Erro kkkkkkkkkkk", f"Falha ao remover bloqueios:\n\nalá tentou remover o bloqueio 'Intervalinho' kkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkk")
            return
        
        confirm = messagebox.askyesno("Info agendamento", f"Deseja remover {'o bloqueio' if count == 1 else f'os {count-1} bloqueios'} do dia {date_str}?")
        if not confirm:
            return
        
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            conn.Open(self.str_conn)
            date_str_sql = exc_date.strftime('%Y-%m-%d')
            conn.Execute(f"DELETE FROM [IST-PGE].dbo.Calendar_Exceptions WHERE exception_date = '{date_str_sql}' AND id_sector = {self.id_sector} AND (notes IS NULL OR notes NOT LIKE '%Intervalinho%')")
            conn.Close()
            
            self.load_calendar_data()
            self.render_calendar()
            self.status_label.config(text=f"Bloqueio(s) removido(s) de {date_str}")
            messagebox.showinfo("Sucesso", f"Bloqueio(s) removido(s) de {date_str}")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao remover bloqueios:\n{str(e)}")

    def get_sharepoint_url(self, os_code):
        if not os_code:
            return ""
        os_year = os_code[-4:]
        os_parts = os_code.split('/')
        os_code_num = os_parts[0].lstrip('0')
        os_code_num = format_os_code(os_code_num)
        return f"https://sesirs.sharepoint.com/:f:/r/sites/gdms-ISISistemasdeSensoriamento/Documentos%20Compartilhados/ISI%20SIM%20-%20Metrologia/Atendimento%20ao%20Cliente/E%20-%20Ordens%20de%20Servi%C3%A7o/{os_year}/{os_code_num}"

    def load_services(self):
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            rs = win32com.client.Dispatch("ADODB.Recordset")
            conn.Open(self.str_conn)
            
            rs.Open("SELECT id, code, specification, execution_time_minutes, description FROM [IST-PGE].dbo.Service_Modes_Local ORDER BY code ASC", conn)
            
            self.services_data = []
            if not rs.EOF:
                rs.MoveFirst()
                while not rs.EOF:
                    self.services_data.append({
                        'id': rs.Fields('id').Value,
                        'code': rs.Fields('code').Value,
                        'specification': rs.Fields('specification').Value,
                        'description': rs.Fields('description').Value,
                        'execution_time_minutes': rs.Fields('execution_time_minutes').Value
                    })
                    rs.MoveNext()
            rs.Close()
            conn.Close()
        except Exception as e:
            self.status_label.config(text=f"Erro ao carregar serviços: {str(e)}")

    def clear_all_scheduled(self):
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            conn.Open(self.str_conn)
            conn.Execute(f"DELETE FROM [IST-PGE].dbo.Time_Slots WHERE id_sector = {self.id_sector}")
            conn.Execute(f"DELETE FROM [IST-PGE].dbo.Service_Schedule WHERE id_sector = {self.id_sector}")
            conn.Close()
            
            self.load_calendar_data()
            self.render_calendar()
            self.update_queue_count()
            self.status_label.config(text="Os registros da agenda foram removidos")
            messagebox.showinfo("Sucesso", "Todos os agendamentos foram removidos!")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao limpar agenda:\n{str(e)}")

    def update_queue_count(self):
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            conn.Open(self.str_conn)
            rs = win32com.client.Dispatch("ADODB.Recordset")
            rs.Open(f"SELECT COUNT(*) as cnt FROM [IST-PGE].dbo.Service_Schedule WHERE status = 'PENDING' AND id_sector = {self.id_sector}", conn)
            count = rs.Fields('cnt').Value if not rs.EOF else 0
            rs.Close()
            conn.Close()
        except Exception:
            pass

    def change_month(self, offset):
        self.current_month_offset += offset
        self.load_calendar_data()
        self.render_calendar()

    def go_to_today(self):
        self.current_month_offset = 0
        self.load_calendar_data()
        self.render_calendar()

    def refresh_calendar(self):
        self.load_calendar_data()
        self.render_calendar()
        self.update_queue_count()


class ServiceSchedulerLinkedDirect:
    """ ABA: Serviços para Agendamento (Linked Server) - Multi-Laboratório """
    
    def __init__(self, parent_notebook, str_conn, str_conn_primary, cor_fundo, lab_code="", lab_config=None):
        self.str_conn = str_conn
        self.str_conn_primary = str_conn_primary
        self.cor_fundo = cor_fundo
        self.lab_code = lab_code
        self.lab_config = lab_config or {}
        
        lab_name = self.lab_config.get('name', lab_code)
        sector_id = self.lab_config.get('sector_id', 0)
        
        tab_name = f" Serviços {lab_name} "
        self.frame_certificados = ttk.Frame(parent_notebook)
        parent_notebook.add(self.frame_certificados, text=tab_name)

        parent_notebook.bind('<<NotebookTabChanged>>', self.on_tab_changed)
    
        self.selected_services = []
        self.all_rows = []
        
        self.setup_ui()
        self.load_services()
        self.cleanup_pending_services()
        
    def on_tab_changed(self, event):
        notebook = event.widget
        current_tab = notebook.select()
        current_tab_text = notebook.tab(current_tab, "text")
        
        if self.lab_config.get('name', '') in current_tab_text:
            self.load_services()

    def setup_ui(self):
        lab_name = self.lab_config.get('name', '')
        
        main_container = tk.Frame(self.frame_certificados, bg=self.cor_fundo)
        main_container.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Top frame
        top_frame = tk.Frame(main_container, bg=self.cor_fundo)
        top_frame.pack(fill='x', pady=(0, 10))
        
        self.greetings = ttk.Label(top_frame, text=f"{buenas}", font=("Segoe UI", 10), background='black')
        self.greetings.pack(side='left', padx=5)
        
        tk.Label(top_frame, text=f"Laboratório de {lab_name}",font=('Segoe UI', 12, 'bold'), bg=self.cor_fundo, fg='white').pack(side='left', padx=20)
        
        self.lbl_count = tk.Label(top_frame, text="0 selecionados", font=('Segoe UI', 9, 'bold'),bg=self.cor_fundo, fg='#FFD700')
        self.lbl_count.pack(side='right', padx=20)
        
        # Search
        search_frame = tk.Frame(main_container, bg=self.cor_fundo)
        search_frame.pack(fill='x', pady=(0, 5))
        
        tk.Label(search_frame, text="Buscar:", font=('Segoe UI', 9), bg=self.cor_fundo, fg='white').pack(side='left', padx=(0, 5))
        
        self.search_var = tk.StringVar()
        self.search_var.trace('w', lambda *args: self.filter_treeview())
        
        self.search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=40)
        self.search_entry.pack(side='left', padx=5)
        
        # Treeview
        tree_frame = ttk.Frame(main_container)
        tree_frame.pack(fill='both', expand=True)
        
        tree_scroll_y = ttk.Scrollbar(tree_frame)
        tree_scroll_y.pack(side='right', fill='y')
        
        tree_scroll_x = ttk.Scrollbar(tree_frame, orient="horizontal")
        tree_scroll_x.pack(side='bottom', fill='x')
        
        columns = ("OS", "Item", "Especificação", "Descrição", "Código", "Duração")
        
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings",
                                  yscrollcommand=tree_scroll_y.set,
                                  xscrollcommand=tree_scroll_x.set,
                                  selectmode="extended")
        
        tree_scroll_y.config(command=self.tree.yview)
        tree_scroll_x.config(command=self.tree.xview)
        
        column_widths = {"OS": 80, "Item": 120, "Especificação": 200, "Descrição": 250, "Código": 100, "Duração": 70}
        
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=column_widths.get(col, 80), minwidth=50)
        
        self.tree.pack(expand=True, fill='both')
        
        self.tree.bind('<Button-3>', self.show_context_menu)
        self.tree.bind('<<TreeviewSelect>>', self.on_tree_select)
        self.tree.bind('<Control-a>', self.select_all)
        self.tree.bind('<Control-A>', self.select_all)
        
        button_frame = tk.Frame(main_container, bg=self.cor_fundo)
        button_frame.pack(fill='x', pady=(10, 0))

        self.btn_select_all = ttk.Button(button_frame, text=" Selecionar Todos ", command=self.select_all, cursor='hand2')
        self.btn_select_all.pack(side="left", padx=5)

        self.btn_schedule = ttk.Button(button_frame, text=" Agendar Selecionados ", command=self.schedule_selected, cursor='hand2')
        self.btn_schedule.pack(side="left", padx=5)
        self.btn_schedule.config(state="disabled")

        btn_refresh = ttk.Button(button_frame, text=" Atualizar Lista ", command=self.load_services, cursor='hand2')
        btn_refresh.pack(side="left", padx=5)

        self.btn_clear = ttk.Button(button_frame, text=" Limpar Seleção ", command=self.clear_selection, cursor='hand2')
        self.btn_clear.pack(side="left", padx=5)
        
        self.btn_exception = ttk.Button(button_frame, text=" Adicionar Bloqueio de Agenda ", command=self.add_calendar_exception, cursor='hand2')
        self.btn_exception.pack(side="left", padx=5)

        self.status_label = ttk.Label(self.frame_certificados,text="Selecione os serviços e use o botão direito para agendar",font=("Segoe UI", 12))
        self.status_label.pack(pady=5)

    def cleanup_pending_services(self):
        """Remove PENDING services that have no schedule (orphaned)"""
        sector_id = self.lab_config.get('sector_id', 0)
        
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            conn.Open(self.str_conn_primary)
            
            # Get orphaned services
            rs = win32com.client.Dispatch("ADODB.Recordset")
            sql = f"""
                SELECT ss.id, ss.notes
                FROM [IST-PGE].dbo.Service_Schedule ss
                WHERE ss.status = 'PENDING'
                AND ss.scheduled_start IS NULL
                AND ss.scheduled_end IS NULL
                AND ss.id_sector = {sector_id}
            """
            rs.Open(sql, conn)
            
            orphaned = []
            if not rs.EOF:
                rs.MoveFirst()
                while not rs.EOF:
                    orphaned.append({
                        'id': rs.Fields('id').Value,
                        'notes': rs.Fields('notes').Value if rs.Fields('notes').Value else "Sem observações"
                    })
                    rs.MoveNext()
            rs.Close()
            
            if orphaned:
                # Delete them
                ids = ','.join(str(o['id']) for o in orphaned)
                conn.Execute(f"DELETE FROM [IST-PGE].dbo.Service_Schedule WHERE id IN ({ids})")
                
                # Build message
                count = len(orphaned)
                if count == 1:
                    msg = f"Foi removido 1 registro pendente sem agendamento:\n\n"
                else:
                    msg = f"Foram removidos {count} registros com erro de agendamento:\n\n"
                
                for o in orphaned[:5]:  # Show max 5
                    msg += f"• {o['notes'][:80]}\n"
                
                if count > 5:
                    msg += f"\n... e mais {count - 5} registro(s)."
                
                messagebox.showinfo("Limpeza de Registros Pendentes", msg)
            
            conn.Close()
            
        except Exception as e:
            print(f"Error cleaning up pending services: {e}")

    def show_loading_screen(self, message="Carregando..."):
        """Show a loading overlay with dynamic info"""
        self.loading_popup = tk.Toplevel(self.frame_certificados)
        self.loading_popup.geometry("350x130")
        self.loading_popup.configure(bg=self.cor_fundo)
        self.loading_popup.title("")
        self.loading_popup.overrideredirect(True)
        self.loading_popup.attributes('-topmost', True)
        
        # Center on parent
        self.loading_popup.update_idletasks()
        x = self.frame_certificados.winfo_rootx() + (self.frame_certificados.winfo_width() // 2) - 175
        y = self.frame_certificados.winfo_rooty() + (self.frame_certificados.winfo_height() // 2) - 65
        self.loading_popup.geometry(f"+{x}+{y}")
        
        frame = tk.Frame(self.loading_popup, bg=self.cor_fundo, highlightbackground="#FFFFFF", highlightthickness=2)
        frame.pack(fill='both', expand=True, padx=2, pady=2)

        tk.Label(frame, text=f"Oioi, {username}! O app está carregando:", font=('Segoe UI', 10), background='black', fg="#FFFFFF").pack(pady=(15, 5))

        #tk.Label(frame, text=f"Laboratório de {self.lab_}", font=('Segoe UI', 13, 'bold'), bg=self.cor_fundo, fg="#FFFFFF").pack(pady=(5, 5))
        
        # Message
        tk.Label(frame, text=message, font=('Segoe UI', 11),bg=self.cor_fundo, fg='white').pack(pady=(5, 5))
        
        self.loading_popup.grab_set()
        self.loading_popup.update()

    def hide_loading_screen(self):
        """Hide the loading overlay"""
        if hasattr(self, 'loading_popup') and self.loading_popup.winfo_exists():
            self.loading_popup.grab_release()
            self.loading_popup.destroy()

    def load_services(self):
        """Load services from linked server for this lab"""
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        self.all_rows.clear()
        
        lab_code = self.lab_code
        sector_id = self.lab_config.get('sector_id', 0)

        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            rs = win32com.client.Dispatch("ADODB.Recordset")
            conn.Open(self.str_conn)

            sql = f"""
                SELECT 
                    os.code AS 'OS',
                    i.code AS 'Item',
                    sm.specification AS 'Especificação',
                    sm.description AS 'Descrição',
                    sm.code AS 'Code',
                    sm.execution_time AS 'execution_time'

                FROM instruments_services iss

                LEFT JOIN orders_services AS os ON os.id = iss.id_service_order 
                LEFT JOIN service_modes AS sm ON sm.id = iss.id_service
                LEFT JOIN instruments AS i ON i.id = iss.id_instrument AND i.id_service_order = os.id
                LEFT JOIN budgets AS b ON b.id_order_service = os.id

                WHERE os.removed = 0
                AND iss.removed = 0
                AND sm.removed = 0
                AND i.removed = 0
                AND b.is_last_revision = 1
                AND i.id_current_sector = {sector_id}
                AND sm.code LIKE '%631{lab_code.lower()}%'
                ORDER BY os.receiving_date ASC
            """
            #sql = f"""
            #    SELECT 
            #        os.code AS 'OS',
            #        i.code AS 'Item',
            #        sm.specification AS 'Especificação',
            #        sm.description AS 'Descrição',
            #        sm.code AS 'Code',
            #        sm.execution_time AS 'execution_time'
#
            #    FROM instruments_services iss
#
            #    LEFT JOIN orders_services AS os ON os.id = iss.id_service_order 
            #    LEFT JOIN service_modes AS sm ON sm.id = iss.id_service
            #    LEFT JOIN instruments AS i ON i.id = iss.id_instrument AND i.id_service_order = os.id
            #    LEFT JOIN budgets AS b ON b.id_order_service = os.id
#
            #    WHERE os.removed = 0
            #    AND iss.removed = 0
            #    AND sm.removed = 0
            #    AND i.removed = 0
            #    AND b.is_last_revision = 1
            #    --AND (i.id_current_sector = {sector_id} OR i.id_current_sector = 1)
            #    AND (i.id_current_sector = {sector_id})
            #    AND sm.code LIKE '%631{lab_code.lower()}%'
            #"""
            #print('sql\n\n',sql)
            rs.Open(sql, conn)
            
            if not rs.EOF:
                rs.MoveFirst()
                while not rs.EOF:
                    execution_time_raw = rs.Fields('execution_time').Value
                    duration_minutes = self.convert_time_to_minutes(execution_time_raw) if execution_time_raw else 0
                    
                    os_code = rs.Fields('OS').Value if rs.Fields('OS').Value else ""
                    os_year = os_code[-4:]
                    os_parts = os_code.split('/')
                    os_code_num = f"{os_parts[0].lstrip('0')}/{os_year}"

                    row_data = {
                        'OS': os_code_num,
                        'Item': rs.Fields('Item').Value if rs.Fields('Item').Value else "",
                        'Especificação': rs.Fields('Especificação').Value if rs.Fields('Especificação').Value else "",
                        'Descrição': rs.Fields('Descrição').Value if rs.Fields('Descrição').Value else "",
                        'Code': rs.Fields('Code').Value if rs.Fields('Code').Value else "",
                        'execution_time_raw': str(execution_time_raw) if execution_time_raw else "00:00",
                        'duration_minutes': duration_minutes,
                    }
                    
                    self.all_rows.append(row_data)
                    
                    self.tree.insert("", "end", values=(
                        row_data['OS'], row_data['Item'], row_data['Especificação'],
                        row_data['Descrição'], row_data['Code'], f"{duration_minutes} min"
                    ))
                    rs.MoveNext()
            
            rs.Close()
            conn.Close()
            
            count = len(self.all_rows)
            lab_name = self.lab_config.get('name', '')
            if count == 0:
                self.status_label.config(text=f"Nenhum serviço encontrado para {lab_name}")
            elif count == 1:
                self.status_label.config(text=f"1 serviço carregado")
            else:
                self.status_label.config(text=f"{count} serviços carregados")
            
        except Exception as e:
            self.status_label.config(text=f"Erro: {str(e)}")
    
    def convert_time_to_minutes(self, time_str):
        if not time_str:
            return 0
        try:
            time_str = str(time_str).strip()
            if ':' in time_str:
                parts = time_str.split(':')
                return int(parts[0]) * 60 + int(parts[1])
            return 0
        except Exception:
            return 0
    
    def filter_treeview(self):
        search_term = self.search_var.get().lower()
        
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        for row in self.all_rows:
            if not search_term or \
               search_term in str(row['OS']).lower() or \
               search_term in str(row['Item']).lower() or \
               search_term in str(row['Especificação']).lower() or \
               search_term in str(row['Descrição']).lower() or \
               search_term in str(row['Code']).lower():
                
                self.tree.insert("", "end", values=(
                    row['OS'], row['Item'], row['Especificação'],
                    row['Descrição'], row['Code'], f"{row['duration_minutes']} min"
                ))
    
    def on_tree_select(self, event):
        selected = self.tree.selection()
        count = len(selected)
        self.lbl_count.config(text=f"{count} selecionado(s)")
        
        if count > 0:
            self.btn_schedule.config(state='normal')
            self.btn_clear.config(state='normal')
        else:
            self.btn_schedule.config(state='disabled')
            self.btn_clear.config(state='disabled')
    
    def show_context_menu(self, event):
        selected = self.tree.selection()
        if not selected:
            return
        
        context_menu = tk.Menu(self.tree, tearoff=0)
        count = len(selected)
        
        if count == 1:
            context_menu.add_command(label=f"Agendar serviço", command=self.schedule_selected)
        elif count > 1:
            context_menu.add_command(label=f"Agendar serviços", command=self.schedule_selected)
        
        context_menu.add_separator()
        context_menu.add_command(label="Selecionar Todos", command=self.select_all)
        context_menu.add_command(label="Limpar Seleção", command=self.clear_selection)
        
        context_menu.post(event.x_root, event.y_root)
    
    def select_all(self, event=None):
        for item in self.tree.get_children():
            self.tree.selection_add(item)
        self.on_tree_select(None)
    
    def clear_selection(self):
        for item in self.tree.selection():
            self.tree.selection_remove(item)
        self.on_tree_select(None)
    
    def schedule_selected(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Aviso", "Selecione pelo menos um serviço!")
            return
        
        selected_indices = []
        all_items = self.tree.get_children()
                
        for item in selected:
            idx = all_items.index(item)
            if idx < len(self.all_rows):
                selected_indices.append(idx)
        
        sector_id = self.lab_config.get('sector_id', 0)
        lab_name = self.lab_config.get('name', '')

        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            conn.Open(self.str_conn_primary)
            
            added_count = 0
            for idx in selected_indices:
                row = self.all_rows[idx]
                duration = row['duration_minutes']
                
                # Sync to Service_Modes_Local
                sync_sql = f"""
                    IF NOT EXISTS (SELECT 1 FROM [IST-PGE].dbo.Service_Modes_Local WHERE code = '{row['Code']}')
                    BEGIN
                        INSERT INTO [IST-PGE].dbo.Service_Modes_Local (code, specification, description, execution_time_minutes)
                        VALUES ('{row['Code']}', '{str(row['Especificação']).replace("'", "''")}', '{str(row['Descrição']).replace("'", "''")}', {duration});
                    END
                    ELSE
                    BEGIN
                        UPDATE [IST-PGE].dbo.Service_Modes_Local
                        SET execution_time_minutes = {duration},
                            specification = '{str(row['Especificação']).replace("'", "''")}'
                        WHERE code = '{row['Code']}';
                    END
                """
                print('sync_sql',sync_sql)
                conn.Execute(sync_sql)
                
                rs = win32com.client.Dispatch("ADODB.Recordset")
                print(f"SELECT id FROM [IST-PGE].dbo.Service_Modes_Local WHERE code = '{row['Code']}'")
                rs.Open(f"SELECT id FROM [IST-PGE].dbo.Service_Modes_Local WHERE code = '{row['Code']}'", conn)

                local_service_id = rs.Fields('id').Value if not rs.EOF else None
                rs.Close()
                
                if local_service_id:
                    notes = f"OS {row['OS']} - {row['Item']} - {row['Especificação']}"
                    schedule_sql = f"""
                        INSERT INTO [IST-PGE].dbo.Service_Schedule (service_id, notes, id_sector)
                        VALUES ({local_service_id}, '{notes.replace("'", "''")}', {sector_id})
                    """
                    print('schedule_sql\n\n',schedule_sql)
                    conn.Execute(schedule_sql)
                    added_count += 1
            
            conn.Close()
            
            if added_count == 1:
                self.status_label.config(text=f"{added_count} serviço adicionado à fila!")
            elif added_count == 0:
                self.status_label.config(text=f"Nenhum serviço adicionado à fila!")
            else:
                self.status_label.config(text=f"{added_count} serviços adicionados à fila!")
                
            self.clear_selection()
            self.process_fifo()
            self.hide_loading_screen()
            
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao agendar:\n{str(e)}")
    
    def process_fifo(self):
        
        sector_id = self.lab_config.get('sector_id', 0)
        lab_name = self.lab_config.get('name', '')
        
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            conn.Open(self.str_conn_primary)
            
            rs = win32com.client.Dispatch("ADODB.Recordset")
            
            sql = f"""
                SELECT ss.id as schedule_id, ss.service_id,
                       s.execution_time_minutes, ss.priority, ss.requested_at, ss.notes
                FROM [IST-PGE].dbo.Service_Schedule ss
                JOIN [IST-PGE].dbo.Service_Modes_Local s ON ss.service_id = s.id
                WHERE ss.status = 'PENDING' AND ss.id_sector = {sector_id}
                ORDER BY ss.priority ASC, ss.requested_at ASC
            """
            print('sql\n\n',sql)
            rs.Open(sql, conn)
            
            pending_services = []
            if not rs.EOF:
                rs.MoveFirst()
                while not rs.EOF:
                    pending_services.append({
                        'schedule_id': rs.Fields('schedule_id').Value,
                        'service_id': rs.Fields('service_id').Value,
                        'execution_time_minutes': rs.Fields('execution_time_minutes').Value,
                        'priority': rs.Fields('priority').Value,
                        'requested_at': rs.Fields('requested_at').Value,
                        'notes': rs.Fields('notes').Value
                    })
                    rs.MoveNext()
            rs.Close()
            
            pending_count = len(pending_services)
            
            #self.show_loading_screen(f"Carregando agenda para {pending_count} {"itens" if pending_count > 1 else "item"}...")

            if pending_count == 0:
                messagebox.showinfo("Aviso", "Nenhum serviço pendente na fila!")
                conn.Close()
                return
            
            self.show_loading_screen(f"Agendando {pending_count} serviço(s) em {lab_name}...")
            self.frame_certificados.update()

            current_time = datetime.now()
            scheduled_count = 0
            for service in pending_services:
                print("next_start, next_end = self.find_next_available_slot(current_time, service['execution_time_minutes'])")
                next_start, next_end = self.find_next_available_slot(current_time, service['execution_time_minutes'])
                
                print("if next_start is None:")

                if next_start is None:
                    continue
                conn.Execute(f"""
                    UPDATE [IST-PGE].dbo.Service_Schedule
                    SET scheduled_start = '{next_start.strftime('%Y-%m-%d %H:%M:%S')}',
                        scheduled_end = '{next_end.strftime('%Y-%m-%d %H:%M:%S')}',
                        status = 'SCHEDULED', updated_at = GETDATE()
                    WHERE id = {service['schedule_id']}
                """)
                print('statement 1\n\n',f"""
                    UPDATE [IST-PGE].dbo.Service_Schedule
                    SET scheduled_start = '{next_start.strftime('%Y-%m-%d %H:%M:%S')}',
                        scheduled_end = '{next_end.strftime('%Y-%m-%d %H:%M:%S')}',
                        status = 'SCHEDULED', updated_at = GETDATE()
                    WHERE id = {service['schedule_id']}
                """)
                
                conn.Execute(f"""
                    INSERT INTO [IST-PGE].dbo.Time_Slots (slot_date, start_time, end_time, schedule_id, service_id, id_sector)
                    VALUES ('{next_start.strftime('%Y-%m-%d')}', '{next_start.strftime('%H:%M:%S')}',
                            '{next_end.strftime('%H:%M:%S')}', {service['schedule_id']},
                            {service['service_id']}, {sector_id})
                """)
                
                print('statement 2\n\n',f"""
                    INSERT INTO [IST-PGE].dbo.Time_Slots (slot_date, start_time, end_time, schedule_id, service_id, id_sector)
                    VALUES ('{next_start.strftime('%Y-%m-%d')}', '{next_start.strftime('%H:%M:%S')}',
                            '{next_end.strftime('%H:%M:%S')}', {service['schedule_id']},
                            {service['service_id']}, {sector_id})
                """)
                
                current_time = next_end
                scheduled_count += 1
            
            self.hide_loading_screen()
            conn.Close()
            
            if scheduled_count == 1:
                messagebox.showinfo("Info agendamento", f"{scheduled_count} serviço agendado em {lab_name}!")
            else:
                messagebox.showinfo("Info agendamento", f"{scheduled_count} serviços agendados em {lab_name}!")
            
        except Exception as e:
            self.hide_loading_screen()
            messagebox.showerror("Erro", f"Falha ao processar fila:\n{str(e)}")
    
    # ===== FIFO ENGINE (same as ServiceScheduler) =====
    def find_next_available_slot(self, from_time, duration_minutes):
        
        
        check_date = from_time.date()
        max_days = 365
        
        availability = self._get_availability()
        exceptions = self._get_exceptions()
        
        for _ in range(max_days):
            
            sql_day_of_week = (check_date.weekday() + 1) % 7
            day_windows = [w for w in availability if w['day_of_week'] == sql_day_of_week]
            
            if not day_windows:
                check_date += timedelta(days=1)
                from_time = datetime.combine(check_date, time(8, 0))
                continue
            
            full_day_blocked = any(ex['exception_date'] == check_date and ex['is_available'] == False and ex['start_time'] is None for ex in exceptions)
            
            if full_day_blocked:
                check_date += timedelta(days=1)
                from_time = datetime.combine(check_date, time(8, 0))
                continue
            
            loop_count = 0
            for window in sorted(day_windows, key=lambda w: w['start_time']):
                
                window_start_dt = datetime.combine(check_date, window['start_time'])
                window_end_dt = datetime.combine(check_date, window['end_time'])
                proposed_start = max(from_time, window_start_dt)

                print("while proposed_start + timedelta(minutes=duration_minutes) <= window_end_dt:")
                while proposed_start + timedelta(minutes=duration_minutes) <= window_end_dt:
                    loop_count += 1
                    print('loop_count: ', loop_count)
                    proposed_end = proposed_start + timedelta(minutes=duration_minutes)
                    
                    blocked_by_exception = False
                    for ex in exceptions:
                        if (ex['exception_date'] == check_date and ex['is_available'] == False
                            and ex['start_time'] is not None and ex['end_time'] is not None):
                            ex_start_dt = datetime.combine(check_date, ex['start_time'])
                            ex_end_dt = datetime.combine(check_date, ex['end_time'])
                            if proposed_start < ex_end_dt and proposed_end > ex_start_dt:
                                blocked_by_exception = True
                                proposed_start = ex_end_dt
                                
                                self.hide_loading_screen()
                                break
                    
                    if blocked_by_exception:
                        continue
                    
                    if self._has_time_slot_conflict(check_date, proposed_start.time(), proposed_end.time()):
                        
                        next_free = self._get_next_free_time(check_date, proposed_start.time())
                        if next_free is None:
                            
                            self.hide_loading_screen()
                            break
                        proposed_start = datetime.combine(check_date, next_free)
                        continue
                    #print('GET OUT')
                    return proposed_start, proposed_end
                
            check_date += timedelta(days=1)
            if day_windows:
                earliest = min(w['start_time'] for w in day_windows)
                from_time = datetime.combine(check_date, earliest)
            else:
                from_time = datetime.combine(check_date, time(8, 0))
                
        return None, None

    def _get_availability(self):
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            rs = win32com.client.Dispatch("ADODB.Recordset")
            conn.Open(self.str_conn_primary)
            
            rs.Open("SELECT day_of_week, start_time, end_time FROM [IST-PGE].dbo.Calendar_Availability WHERE is_active = 1 ORDER BY day_of_week, start_time", conn)
            
            result = []
            if not rs.EOF:
                rs.MoveFirst()
                while not rs.EOF:
                    start_t = rs.Fields('start_time').Value
                    end_t = rs.Fields('end_time').Value
                    if isinstance(start_t, str):
                        parts = start_t.split(':')
                        start_t = time(int(parts[0]), int(parts[1]))
                    if isinstance(end_t, str):
                        parts = end_t.split(':')
                        end_t = time(int(parts[0]), int(parts[1]))
                    
                    result.append({'day_of_week': rs.Fields('day_of_week').Value, 'start_time': start_t, 'end_time': end_t})
                    rs.MoveNext()
            rs.Close()
            conn.Close()
            return result
        except Exception as e:
            print(f"Error loading availability: {e}")
            return []

    def _get_exceptions(self):
        sector_id = self.lab_config.get('sector_id', 0)
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            rs = win32com.client.Dispatch("ADODB.Recordset")
            conn.Open(self.str_conn_primary)
            
            rs.Open(f"SELECT exception_date, exception_type, start_time, end_time, is_available FROM [IST-PGE].dbo.Calendar_Exceptions WHERE id_sector = {sector_id}", conn)
            
            result = []
            if not rs.EOF:
                rs.MoveFirst()
                while not rs.EOF:
                    exc_date = rs.Fields('exception_date').Value
                    if isinstance(exc_date, str):
                        exc_date = datetime.strptime(exc_date, '%Y-%m-%d').date()
                    elif isinstance(exc_date, datetime):
                        exc_date = exc_date.date()
                    
                    start_t = rs.Fields('start_time').Value
                    end_t = rs.Fields('end_time').Value
                    if start_t and isinstance(start_t, str):
                        parts = start_t.split(':')
                        start_t = time(int(parts[0]), int(parts[1]))
                    if end_t and isinstance(end_t, str):
                        parts = end_t.split(':')
                        end_t = time(int(parts[0]), int(parts[1]))
                    
                    result.append({'exception_date': exc_date, 'exception_type': rs.Fields('exception_type').Value,
                                   'start_time': start_t, 'end_time': end_t, 'is_available': rs.Fields('is_available').Value})
                    rs.MoveNext()
            rs.Close()
            conn.Close()
            return result
        except Exception as e:
            print(f"Error loading exceptions: {e}")
            return []

    def _has_time_slot_conflict(self, check_date, start_time, end_time):
        sector_id = self.lab_config.get('sector_id', 0)
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            rs = win32com.client.Dispatch("ADODB.Recordset")
            conn.Open(self.str_conn_primary)
            
            start_str = start_time.strftime('%H:%M:%S') if isinstance(start_time, time) else str(start_time)
            end_str = end_time.strftime('%H:%M:%S') if isinstance(end_time, time) else str(end_time)
            date_str = check_date.strftime('%Y-%m-%d')
            
            rs.Open(f"""
                SELECT COUNT(*) as cnt FROM [IST-PGE].dbo.Time_Slots
                WHERE slot_date = '{date_str}' AND start_time < '{end_str}' AND end_time > '{start_str}' AND id_sector = {sector_id}
            """, conn)
            count = rs.Fields('cnt').Value if not rs.EOF else 0
            rs.Close()
            conn.Close()
            return count > 0
        except Exception as e:
            print(f"Error checking conflicts: {e}")
            return True

    def _get_next_free_time(self, check_date, after_time):
        sector_id = self.lab_config.get('sector_id', 0)
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            rs = win32com.client.Dispatch("ADODB.Recordset")
            conn.Open(self.str_conn_primary)
            
            date_str = check_date.strftime('%Y-%m-%d')
            time_str = after_time.strftime('%H:%M:%S') if isinstance(after_time, time) else str(after_time)
            
            rs.Open(f"""
                SELECT TOP 1 end_time FROM [IST-PGE].dbo.Time_Slots
                WHERE slot_date = '{date_str}' AND start_time <= '{time_str}' AND end_time > '{time_str}' AND id_sector = {sector_id}
                ORDER BY end_time ASC
            """, conn)
            
            if not rs.EOF:
                end_time = rs.Fields('end_time').Value
                result = self._parse_time_value(end_time)
                rs.Close()
                conn.Close()
                return result
            rs.Close()
            
            rs.Open(f"""
                SELECT TOP 1 end_time FROM [IST-PGE].dbo.Time_Slots
                WHERE slot_date = '{date_str}' AND start_time > '{time_str}' AND id_sector = {sector_id}
                ORDER BY start_time ASC
            """, conn)
            
            if not rs.EOF:
                end_time = rs.Fields('end_time').Value
                result = self._parse_time_value(end_time)
                rs.Close()
                conn.Close()
                return result
            
            rs.Close()
            conn.Close()
            return None
        except Exception as e:
            print(f"Error getting next free time: {e}")
            return None

    def _parse_time_value(self, value):
        if value is None: return None
        if isinstance(value, time): return value
        if isinstance(value, str):
            value = value.split('.')[0]
            if ':' in value:
                parts = value.split(':')
                return time(int(parts[0]), int(parts[1]), 0)
            return time(int(value), 0, 0)
        if isinstance(value, float):
            total_hours = value * 24
            return time(int(total_hours), int((total_hours - int(total_hours)) * 60), 0)
        if isinstance(value, datetime): return time(value.hour, value.minute, 0)
        return None

    # ===== CALENDAR EXCEPTION (ADD) =====
    def add_calendar_exception(self):
        popup = tk.Toplevel(self.frame_certificados)
        popup.geometry("400x300")
        popup.minsize(400, 350)
        popup.configure(bg='black')
        popup.title("Adicionar Bloqueio ao Calendário")
        
        tk.Label(popup, text="Data:", font=('Segoe UI', 9), bg='black', fg='white').pack(anchor='w', padx=20, pady=(15, 0))
        
        date_frame = tk.Frame(popup, bg='black')
        date_frame.pack(fill='x', padx=20, pady=(2, 5))
        
        self.exc_day_var = tk.StringVar(value=datetime.now().strftime('%d'))
        self.exc_month_var = tk.StringVar(value=datetime.now().strftime('%m'))
        self.exc_year_var = tk.StringVar(value=datetime.now().strftime('%Y'))
        
        ttk.Entry(date_frame, textvariable=self.exc_day_var, width=3).pack(side='left')
        tk.Label(date_frame, text="/", bg='black', fg='white').pack(side='left')
        ttk.Entry(date_frame, textvariable=self.exc_month_var, width=3).pack(side='left')
        tk.Label(date_frame, text="/", bg='black', fg='white').pack(side='left')
        ttk.Entry(date_frame, textvariable=self.exc_year_var, width=5).pack(side='left')
        
        tk.Label(popup, text="Horário de Início (HH:MM):", font=('Segoe UI', 9), bg='black', fg='white').pack(anchor='w', padx=20, pady=(10, 0))
        self.exc_start_var = tk.StringVar(value="")
        exc_start_entry = ttk.Entry(popup, textvariable=self.exc_start_var, width=10)
        exc_start_entry.pack(anchor='w', padx=20, pady=(2, 5))

        tk.Label(popup, text="Horário de Encerramento (HH:MM):", font=('Segoe UI', 9), bg='black', fg='white').pack(anchor='w', padx=20, pady=(5, 0))
        self.exc_end_var = tk.StringVar(value="")
        exc_end_entry = ttk.Entry(popup, textvariable=self.exc_end_var, width=10)
        exc_end_entry.pack(anchor='w', padx=20, pady=(2, 5))

        tk.Label(popup, text="Razão:", font=('Segoe UI', 9), bg='black', fg='white').pack(anchor='w', padx=20, pady=(5, 0))
        self.exc_notes_var = tk.StringVar(value="")
        ttk.Entry(popup, textvariable=self.exc_notes_var, width=40).pack(anchor='w', padx=20, pady=(2, 5))

        def toggle_time_fields():
            if self.exc_full_day.get():
                exc_start_entry.config(state='disabled')
                exc_end_entry.config(state='disabled')
            else:
                exc_start_entry.config(state='normal')
                exc_end_entry.config(state='normal')

        self.exc_full_day = tk.BooleanVar(value=False)
        ttk.Checkbutton(popup, text="Dia inteiro", variable=self.exc_full_day, command=toggle_time_fields).pack(anchor='w', padx=20, pady=(10, 5))

        btn_frame = tk.Frame(popup, bg='black')
        btn_frame.pack(pady=15)
        ttk.Button(btn_frame, text=" Salvar ", command=lambda: self.save_exception(popup)).pack(side='left', padx=5)
        ttk.Button(btn_frame, text=" Cancelar ", command=popup.destroy).pack(side='left', padx=5)

    def save_exception(self, popup):
        sector_id = self.lab_config.get('sector_id', 0)
        try:
            day = self.exc_day_var.get().strip().zfill(2)
            month = self.exc_month_var.get().strip().zfill(2)
            year = self.exc_year_var.get().strip()
            exception_date_sql = f"{year}-{month}-{day}"
            exception_date_br = f"{day}/{month}/{year}"
            datetime.strptime(exception_date_sql, '%Y-%m-%d')
            
            is_full_day = self.exc_full_day.get()
            start_time = None if is_full_day else self.exc_start_var.get().strip()
            end_time = None if is_full_day else self.exc_end_var.get().strip()
            notes = self.exc_notes_var.get().strip() or "Indisponibilidade"
            
            if not is_full_day:
                if not start_time or not end_time:
                    messagebox.showwarning("Aviso", "Informe horário de início e fim!")
                    return
                for t in [start_time, end_time]:
                    parts = t.split(':')
                    if len(parts) != 2 or not (0 <= int(parts[0]) <= 23 and 0 <= int(parts[1]) <= 59):
                        messagebox.showwarning("Aviso", "Formato de horário inválido! Use HH:MM")
                        return
                start_time_str = f"{start_time}:00"
                end_time_str = f"{end_time}:00"
            else:
                start_time_str = None
                end_time_str = None
            
            conn = win32com.client.Dispatch("ADODB.Connection")
            conn.Open(self.str_conn_primary)
            
            sql = f"""
                INSERT INTO [IST-PGE].dbo.Calendar_Exceptions (exception_date, exception_type, start_time, end_time, is_available, notes, id_sector)
                VALUES ('{exception_date_sql}', 'MODIFIED_HOURS', {f"'{start_time_str}'" if start_time_str else 'NULL'},
                        {f"'{end_time_str}'" if end_time_str else 'NULL'}, 0, '{notes.replace("'", "''")}', {sector_id})
            """
            conn.Execute(sql)
            conn.Close()
            
            popup.destroy()
            
            if is_full_day:
                msg = f"Bloqueio de agenda adicionado:\n\nData: {exception_date_br}\nDia inteiro\nRazão: {notes}"
            else:
                msg = f"Bloqueio de agenda adicionado:\n\nData: {exception_date_br}\nHorário: {start_time} às {end_time}\nRazão: {notes}"
            
            messagebox.showinfo("Info agendamento", msg)
            self.status_label.config(text=f"Bloqueio de calendário adicionado em: {exception_date_br}")
        except ValueError:
            messagebox.showwarning("Aviso", "Data inválida!")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao salvar Bloqueio de Calendário:\n{str(e)}")


# ==================== MAIN ====================
status_servidor, cor_status = verificar_disponibilidade3()
status_servidor3, cor_status3 = verificar_disponibilidade3()

root = tk.Tk()
root.title("Assistente de Secretaria Técnica")
root.geometry("900x500")
root.minsize(1150, 675)
cor_fundo = _from_rgb((27, 75, 159))
root.configure(bg=cor_fundo)

style = ttk.Style()
style.theme_use('vista')
style.configure("TNotebook", background=cor_fundo, padding=5)
style.configure("TNotebook.Tab", font=("Segoe UI", 9, "bold"), padding=[10, 5])
style.configure("TFrame", background=cor_fundo)
style.configure("TLabel", background=cor_fundo, foreground="white", font=("Segoe UI", 10))

notebook = ttk.Notebook(root)
notebook.pack(expand=True, fill="both")

for lab_code, lab_config in LABS.items():
    
    ServiceScheduler(notebook, STR_CONN, cor_fundo,id_sector=lab_config['sector_id'],lab_name=lab_config['name'],lab_config=lab_config)
    ServiceSchedulerLinkedDirect(notebook, STR_CONN_LINKED, STR_CONN, cor_fundo,lab_code=lab_code,lab_config=lab_config)
    
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
