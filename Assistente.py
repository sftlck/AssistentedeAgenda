import win32com.client
import tkinter as tk
import win32timezone
from tkinter import ttk, messagebox
from datetime import datetime, timedelta,date
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

db = "Castro_Services"

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


#STR_CONN_ZEBRA = (
#    f"Provider=SQLOLEDB;"
#    f"Data Source={ip_zebra};"
#    f"Initial Catalog={db_zebra};"
#    f"User ID={user_zebra};"
#    f"Password={password_zebra};"
#)

username00= getuser()
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

def verificar_disponibilidade():
    try:
        conn = win32com.client.Dispatch("ADODB.Connection")
        conn.ConnectionTimeout = 1
        conn.Open(STR_CONN)
        conn.Close()
        return "ONLINE", "lime"
    except Exception:
        return "OFFLINE", "orange"


def verificar_disponibilidade3():
    try:
        conn = win32com.client.Dispatch("ADODB.Connection")
        conn.ConnectionTimeout = 1
        conn.Open(STR_CONN_LINKED)
        conn.Close()
        return "ONLINE", "lime"
    except Exception:
        return "OFFLINE", "orange"

def format_os_code(os_code):
    
    if len(os_code) == 1:
        os_code = f"000{os_code}"
        return os_code
    
    elif len(os_code) == 2:
        os_code = f"00{os_code}"
        return os_code

    elif len(os_code) == 3:
        os_code = f"0{os_code}"
        return os_code


class ServiceScheduler:
    """ ABA: Agendamento de Serviços (Calibração FIFO) """
    
    def __init__(self, parent_notebook, str_conn, cor_fundo):
        self.str_conn = str_conn
        self.cor_fundo = cor_fundo
        
        #self.str_conn_primary = str_conn_primary
        self.frame_certificados = ttk.Frame(parent_notebook)
        parent_notebook.add(self.frame_certificados, text=" Calendário de Pressão [Em desenvolvimento] ")
        
        parent_notebook.bind('<<NotebookTabChanged>>', self.on_tab_changed)
        
        self.current_month_offset = 0
        self.calendar_data = {}
        self.selected_date = None
        
        self.setup_ui()
        self.load_services()
        self.load_calendar_data()
        self.render_calendar()
        self.update_queue_count()

        
        self.start_auto_refresh()
        #self.load_pending_list()

    def start_auto_refresh(self):
        """Start periodic calendar refresh every 20 seconds"""
        self._auto_refresh()

    def _auto_refresh(self):
        """Internal auto-refresh loop"""
        self.refresh_calendar()
        # Schedule next refresh in 5 seconds
        self.frame_certificados.after(8000, self._auto_refresh)

    def on_tab_changed(self, event):
        """Refresh calendar only when this tab is selected"""
        notebook = event.widget
        current_tab = notebook.select()
        current_tab_text = notebook.tab(current_tab, "text")
        
        if "Agendamento de Serviços" in current_tab_text:
            self.refresh_calendar()

    def setup_ui(self):
        main_container = tk.Frame(self.frame_certificados, bg=self.cor_fundo)
        main_container.pack(fill='both', expand=True, padx=10, pady=10)
        
        
        
        # ===== RIGHT PANEL (Calendar) =====
        self.calendar_frame = tk.Frame(main_container, bg=self.cor_fundo)
        self.calendar_frame.pack(side='left', fill='both', expand=True)
        
        # Status bar
        self.status_label = ttk.Label(self.frame_certificados, 
                                       text="Pronto para agendamento", font=("Segoe UI", 12))
        self.status_label.pack(pady=5)

    def extract_os_from_notes(self, notes):
        """Extract OS code from notes string"""
        if not notes:
            return ""
        # Format: "OS 12345/2026 - ITEM-001 - Specification"
        parts = notes.split(' - ')
        if len(parts) >= 1:
            os_part = parts[0].replace("OS ", "").strip()
            return os_part
        return ""
    
    def load_exceptions_for_date(self, check_date):
        """Load calendar exceptions for a specific date"""
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            rs = win32com.client.Dispatch("ADODB.Recordset")
            
            conn.Open(self.str_conn)
            
            date_str = check_date.strftime('%Y-%m-%d') if isinstance(check_date, date) else str(check_date)
            
            sql = f"""
                SELECT exception_type, start_time, end_time, notes
                FROM [castro_services].dbo.Calendar_Exceptions
                WHERE exception_date = '{date_str}'
                AND is_available = 0
            """
            
            rs.Open(sql, conn)
            
            exceptions = []
            if not rs.EOF:
                rs.MoveFirst()
                while not rs.EOF:
                    exc_type = rs.Fields('exception_type').Value
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
                        exceptions.append(f"Dia inteiro: {notes}")
                    
                    rs.MoveNext()
            
            rs.Close()
            conn.Close()
            
            return exceptions
            
        except Exception as e:
            print(f"Error loading exceptions: {e}")
            return []

    def _get_exceptions(self):
        """Load Calendar_Exceptions into memory"""
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            rs = win32com.client.Dispatch("ADODB.Recordset")
            
            conn.Open(self.str_conn)
            
            sql = """
                SELECT exception_date, exception_type, start_time, end_time, is_available
                FROM [castro_services].dbo.Calendar_Exceptions
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
        """Load all scheduled services grouped by date"""
        self.calendar_data.clear()
        
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            conn.Open(self.str_conn)
            
            sql = """
                SELECT 
                    ts.slot_date,
                    ts.start_time,
                    ts.end_time,
                    s.code,
                    s.specification,
                    s.description,
                    s.execution_time_minutes,
                    ss.notes,
                    ss.status,
                    ss.priority,
                    ss.id as schedule_id
                FROM [castro_services].dbo.Time_Slots ts
                JOIN [castro_services].dbo.Service_Schedule ss ON ts.schedule_id = ss.id
                JOIN [castro_services].dbo.Service_Modes_Local s ON ts.service_id = s.id
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
                    'schedule_id': rs.Fields('schedule_id').Value  # ADD THIS LINE
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
        
        """Render the monthly calendar view"""
        
        for widget in self.calendar_frame.winfo_children():
            widget.destroy()
        
        today = datetime.now()
        target_date = today + relativedelta(months=self.current_month_offset)
        target_date = target_date.replace(day=1)
        year, month = target_date.year, target_date.month
        
        meses_pt = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho','Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
        
        # ===== Navigation Header =====
        nav_frame = tk.Frame(self.calendar_frame, bg=self.cor_fundo)
        nav_frame.pack(fill='x', pady=(0, 5))
        
        # ===== Navigation Header =====
        nav_frame = tk.Frame(self.calendar_frame, bg=self.cor_fundo)
        nav_frame.pack(fill='x', pady=(0, 5))

        self.greetings = ttk.Label(nav_frame, text=f"{buenas}", font=("Segoe UI", 10),background='black')
        self.greetings.pack(side="left", padx=5, pady=(0, 0))
        
        # Row 1: Title
        title_frame = tk.Frame(nav_frame, bg=self.cor_fundo)
        title_frame.pack(fill='x')

        #tk.Label(title_frame, text="Serviços de Pressão",font=('Segoe UI', 12, 'bold'), bg=self.cor_fundo, fg='white').pack(side='left', padx=20)

        # Row 2: Navigation
        nav_row = tk.Frame(nav_frame, bg=self.cor_fundo)
        nav_row.pack(fill='x')

        btn_prev = tk.Label(nav_row, text="◀", anchor='center', justify='center',
                            font=('Segoe UI', 12, 'bold'), bg=self.cor_fundo, fg='white',
                            cursor='hand2', padx=0)
        btn_prev.pack(side='left', padx=5)
        btn_prev.bind('<Button-1>', lambda e: self.change_month(-1))

        month_label = tk.Label(nav_row, text=f"{meses_pt[month-1]} {year}",
                            font=('Segoe UI', 12, 'bold'), bg=self.cor_fundo, fg='white')
        month_label.pack(side='left', padx=5)

        btn_next = tk.Label(nav_row, text="▶", anchor='center', justify='center',
                            font=('Segoe UI', 12, 'bold'), bg=self.cor_fundo, fg='white',
                            cursor='hand2', padx=0)
        btn_next.pack(side='left', padx=5)
        btn_next.bind('<Button-1>', lambda e: self.change_month(1))

        header_frame = tk.Frame(self.calendar_frame, bg=self.cor_fundo)
        header_frame.pack(fill='x')
        
        dias_semana = ['Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb', 'Dom']
        for dia in dias_semana:
            lbl = tk.Label(header_frame, text=dia, font=('Segoe UI', 9, 'bold'),bg='#1B4B9F', fg='white', width=16, pady=4)
            lbl.pack(side='left', padx=1, pady=1)
        
        # ===== Calendar Grid =====
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
                    
                    if is_today:
                        bg_color = '#B7D5F5'
                    elif is_weekend:
                        bg_color = "#908F8F"
                    elif has_orders:
                        bg_color = '#FFFFFF'
                    else:
                        bg_color = '#908F8F'
                    
                    if has_orders and not is_weekend:
                        day_frame = tk.Frame(week_frame, bg=bg_color, width=120, height=80,bd=1, cursor='hand2')
                        day_frame.pack(side='left', padx=1, pady=1)
                        day_frame.pack_propagate(False)
                    else:
                        day_frame = tk.Frame(week_frame, bg=bg_color, width=120, height=80, bd=1)
                        day_frame.pack(side='left', padx=1, pady=1)
                        day_frame.pack_propagate(False)
                    
                    day_num_frame = tk.Frame(day_frame, bg=bg_color)
                    day_num_frame.pack(fill='x', padx=2, pady=1)
                    
                    fg_day = 'white' if is_today else 'black'
                    day_label = tk.Label(day_num_frame, text=str(day),font=('Segoe UI', 9, 'bold'),bg=bg_color, fg=fg_day, anchor='w')
                    day_label.pack(side='left')
                    
                    if has_orders and not is_weekend:
                        count = len(self.calendar_data[calendar_key])
                        badge = tk.Label(day_num_frame, text=str(count),font=('Segoe UI', 8, 'bold'),bg='#2462D2', fg='white', width=3, height=1)
                        badge.pack(side='right')
                    
                    if has_orders and not is_weekend:
                        unique_items = []
                        seen = set()
                        for order in self.calendar_data[calendar_key]:
                            result = self.extract_item_from_notes(order.get('notes', ''))

                            if isinstance(result, tuple) and len(result) >= 2:
                                item_code = result[0]  # The Item code
                            else:
                                item_code = str(result)[:15] if result else ""
                            
                            if item_code and item_code not in seen:
                                unique_items.append(item_code)
                                seen.add(item_code)
                        
                        preview_text = "\n".join(unique_items[:4])
                        
                        preview = tk.Label(day_frame, text=preview_text,font=('Segoe UI', 7), bg=bg_color,fg='#333333', anchor='w', justify='left',cursor='hand2')
                        preview.pack(fill='x', padx=3)

                        if len(unique_items) > 4:
                            more_label = tk.Label(day_frame,text=f"+{len(unique_items) - 4} mais...",font=('Segoe UI', 6), bg=bg_color,fg="#4A4949", anchor='w')
                            more_label.pack(fill='x', padx=3)
                    
                    # Bind left-click to show orders
                    if has_orders and not is_weekend:
                        day_frame.bind('<Button-1>', lambda e, key=calendar_key: self.show_day_orders(key))
                        # Bind right-click for context menu
                        day_frame.bind('<Button-3>', lambda e, key=calendar_key: self.show_day_context_menu(e, key))
                        for child in day_frame.winfo_children():
                            child.bind('<Button-1>', lambda e, key=calendar_key: self.show_day_orders(key))
                            child.bind('<Button-3>', lambda e, key=calendar_key: self.show_day_context_menu(e, key))
        
        btn_today = ttk.Button(grid_frame, text="Hoje", command=self.go_to_today, cursor='hand2')
        btn_today.pack(side="left", padx=5)
        
        #self.buscar = ttk.Button(nav_frame, text="Limpar agenda", command=self.clear_all_scheduled, cursor='hand2')
        #self.buscar.pack(side="left", padx=5)
        
        btn_refresh = ttk.Button(grid_frame, text="Atualizar", command=self.refresh_calendar, cursor='hand2')
        btn_refresh.pack(side="left", padx=5)
        
        self.buscar = ttk.Button(grid_frame, text=" Limpar agenda ", command=self.clear_all_scheduled, cursor='hand2')
        self.buscar.pack(side="left", padx=5)
        

    def extract_item_from_notes(self, notes):
        """Extract item code from notes string"""
        if not notes:
            return notes
        parts = notes.split(' - ')
        if len(parts) >= 2:
            
            return parts[0], parts [1], parts [2]
        return notes[:15]  

    def show_day_context_menu(self, event, calendar_key):
        context_menu = tk.Menu(self.calendar_frame, tearoff=0)
        
        date_str = f"{calendar_key[0]:02d}/{calendar_key[1]:02d}/{calendar_key[2]}"
        count = len(self.calendar_data.get(calendar_key, []))
        
        if count == 1:
            context_menu.add_command(label=f"Ver serviço", command=lambda k=calendar_key: self.show_day_orders(k))
        elif count > 1:
            context_menu.add_command(label=f"Ver serviços", command=lambda k=calendar_key: self.show_day_orders(k))
        
        context_menu.add_separator()
        
        if count == 1:
            context_menu.add_command(label=f"Reagendar o serviço deste dia",command=lambda k=calendar_key: self.reschedule_day_services(k))
        else:
            context_menu.add_command(label=f"Reagendar todos os {count} serviços deste dia",command=lambda k=calendar_key: self.reschedule_day_services(k))
        
        context_menu.add_separator()
        context_menu.add_command(label="Iniciar chat do Teams com Laboratório de Pressão", command=self.open_teams_chat)
        
        context_menu.post(event.x_root, event.y_root)

    def reschedule_day_services(self, calendar_key):
        """Reschedule all services from a specific day to future dates only"""
        if calendar_key not in self.calendar_data:
            messagebox.showinfo("Aviso", "Nenhum serviço neste dia!")
            return
        
        count = len(self.calendar_data[calendar_key])
        date_str = f"{calendar_key[0]:02d}/{calendar_key[1]:02d}/{calendar_key[2]}"
        
        confirm = messagebox.askyesno("Info agendamento",f"Deseja reagendar {count} serviço(s) do dia {date_str}?\n\n")
        
        if not confirm:
            return
        
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            conn.Open(self.str_conn)
            
            reschedule_date = datetime(calendar_key[2], calendar_key[1], calendar_key[0]).date()
            date_str_sql = reschedule_date.strftime('%Y-%m-%d')
            
            # Get schedule_ids for this date
            rs = win32com.client.Dispatch("ADODB.Recordset")
            sql = f"""
                SELECT ts.schedule_id
                FROM [castro_services].dbo.Time_Slots ts
                WHERE ts.slot_date = '{date_str_sql}'
            """
            rs.Open(sql, conn)
            
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
            
            # Delete time slots for this date
            conn.Execute(f"DELETE FROM [castro_services].dbo.Time_Slots WHERE slot_date = '{date_str_sql}'")
            
            # Reset services to PENDING
            id_list = ','.join(str(sid) for sid in schedule_ids)
            conn.Execute(f"""
                UPDATE [castro_services].dbo.Service_Schedule
                SET status = 'PENDING',
                    scheduled_start = NULL,
                    scheduled_end = NULL,
                    updated_at = GETDATE()
                WHERE id IN ({id_list})
            """)
            
            conn.Close()
            
            # Re-schedule starting from NEXT DAY (skip the cleared day)
            self.next_start, self.next_end = self.schedule_all_pending(from_date=reschedule_date + timedelta(days=1))
            
            # Refresh
            self.load_calendar_data()
            self.refresh_calendar()
            self.update_queue_count()

            if count == 1 :
                self.status_label.config(text=f"{count} serviço reagendado para {self.next_start.strftime('%d/%m/%Y')} das {self.next_start.strftime('%H:%M')} às {self.next_end.strftime('%H:%M')}")    
                messagebox.showinfo('Info agendamento',f'{count} serviço reagendado para {self.next_start.strftime('%d/%m/%Y')} das {self.next_start.strftime('%H:%M')} às {self.next_end.strftime('%H:%M')}')

            if count > 1:
                self.status_label.config(text=f"Serviço reagendado para a partir de {self.next_start.strftime('%d/%m/%Y')}")
                messagebox.showinfo('Info agendamento',f'Serviço reagendado para a partir de {self.next_start.strftime('%d/%m/%Y')}')

        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao reagendar:\n{str(e)}")

    def schedule_all_pending(self, from_date=None):
        """Schedule all pending services using Python FIFO"""
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            conn.Open(self.str_conn)
            
            rs = win32com.client.Dispatch("ADODB.Recordset")
            sql = """
                SELECT 
                    ss.id as schedule_id,
                    ss.service_id,
                    s.execution_time_minutes,
                    ss.priority
                FROM [castro_services].dbo.Service_Schedule ss
                JOIN [castro_services].dbo.Service_Modes_Local s ON ss.service_id = s.id
                WHERE ss.status = 'PENDING'
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
                return
            
            # Start from specified date or now
            if from_date:
                current_time = datetime.combine(from_date, time(8, 0))
            else:
                current_time = datetime.now()
            
            for service in pending_services:
                self.next_start, self.next_end = self.find_next_available_slot(
                    current_time,
                    service['execution_time_minutes']
                )
                
                if self.next_start is None:
                    continue
                
                conn.Execute(f"""
                    UPDATE [castro_services].dbo.Service_Schedule
                    SET scheduled_start = '{self.next_start.strftime('%Y-%m-%d %H:%M:%S')}',
                        scheduled_end = '{self.next_end.strftime('%Y-%m-%d %H:%M:%S')}',
                        status = 'SCHEDULED',
                        updated_at = GETDATE()
                    WHERE id = {service['schedule_id']}
                """)
                
                conn.Execute(f"""
                    INSERT INTO [castro_services].dbo.Time_Slots
                        (slot_date, start_time, end_time, schedule_id, service_id)
                    VALUES (
                        '{self.next_start.strftime('%Y-%m-%d')}',
                        '{self.next_start.strftime('%H:%M:%S')}',
                        '{self.next_end.strftime('%H:%M:%S')}',
                        {service['schedule_id']},
                        {service['service_id']}
                    )
                """)
                
                current_time = self.next_end
                
            conn.Close()
            
            return self.next_start, self.next_end
            
        except Exception as e:
            print(f"Error in schedule_all_pending: {e}")

    def find_next_available_slot(self, from_time, duration_minutes):
        """
        Find next available time slot starting from `from_time`
        Replicates sp_FindNextAvailableSlot logic in Python
        
        Returns: (start_datetime, end_datetime) or (None, None) if no slot found
        """
        check_date = from_time.date()
        max_days = 365
        
        # Load all data once
        availability = self._get_availability()
        exceptions = self._get_exceptions()
        
        for day_offset in range(max_days):
            sql_day_of_week = (check_date.weekday() + 1) % 7
            
            # Check if this day has any availability windows
            day_windows = [w for w in availability if w['day_of_week'] == sql_day_of_week]
            
            if not day_windows:
                check_date += timedelta(days=1)
                from_time = datetime.combine(check_date, time(8, 0))
                continue
            
            # Check for full-day exception (holiday)
            full_day_blocked = False
            for ex in exceptions:
                if ex['exception_date'] == check_date and ex['is_available'] == False and ex['start_time'] is None:
                    full_day_blocked = True
                    break
            
            if full_day_blocked:
                check_date += timedelta(days=1)
                from_time = datetime.combine(check_date, time(8, 0))
                continue
            
            # Try each availability window
            for window in sorted(day_windows, key=lambda w: w['start_time']):
                window_start_dt = datetime.combine(check_date, window['start_time'])
                window_end_dt = datetime.combine(check_date, window['end_time'])
                
                # Adjust start if from_time is later
                proposed_start = max(from_time, window_start_dt)
                
                # TRY MULTIPLE POSITIONS WITHIN THIS WINDOW
                while proposed_start + timedelta(minutes=duration_minutes) <= window_end_dt:
                    proposed_end = proposed_start + timedelta(minutes=duration_minutes)
                    
                    # Check for time-specific exceptions (lunch break)
                    blocked_by_exception = False
                    for ex in exceptions:
                        if (ex['exception_date'] == check_date and 
                            ex['is_available'] == False and 
                            ex['start_time'] is not None and 
                            ex['end_time'] is not None):
                            
                            ex_start_dt = datetime.combine(check_date, ex['start_time'])
                            ex_end_dt = datetime.combine(check_date, ex['end_time'])
                            
                            if proposed_start < ex_end_dt and proposed_end > ex_start_dt:
                                blocked_by_exception = True
                                # Jump past the exception
                                proposed_start = ex_end_dt
                                break
                    
                    if blocked_by_exception:
                        continue  # Re-test with new proposed_start
                    
                    # Check for conflicts with existing Time_Slots
                    if self._has_time_slot_conflict(check_date, proposed_start.time(), proposed_end.time()):
                        # Jump to the end of the conflicting slot
                        next_free = self._get_next_free_time(check_date, proposed_start.time())
                        if next_free is None:
                            break  # No more free time in this window
                        proposed_start = datetime.combine(check_date, next_free)
                        continue  # Re-test with new proposed_start
                    
                    # Found a valid slot!
                    return proposed_start, proposed_end
                
                # No slot in this window, try next window
                # Reset proposed_start to beginning of next window or from_time
                pass
            
            # No slot found today, move to next day
            check_date += timedelta(days=1)
            if day_windows:
                earliest = min(w['start_time'] for w in day_windows)
                from_time = datetime.combine(check_date, earliest)
            else:
                from_time = datetime.combine(check_date, time(8, 0))
        
        return None, None

    def _get_next_free_time(self, check_date, after_time):
        """
        Get the end time of the conflicting slot that starts at or contains `after_time`.
        Returns the end time as a time object, or None if at end of day.
        """
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            rs = win32com.client.Dispatch("ADODB.Recordset")
            
            conn.Open(self.str_conn)
            
            date_str = check_date.strftime('%Y-%m-%d')
            time_str = after_time.strftime('%H:%M:%S') if isinstance(after_time, time) else str(after_time)
            
            sql = f"""
                SELECT TOP 1 end_time
                FROM [castro_services].dbo.Time_Slots
                WHERE slot_date = '{date_str}'
                AND start_time <= '{time_str}'
                AND end_time > '{time_str}'
                ORDER BY end_time ASC
            """
            
            rs.Open(sql, conn)
            
            if not rs.EOF:
                end_time = rs.Fields('end_time').Value
                if isinstance(end_time, str):
                    parts = end_time.split(':')
                    result = time(int(parts[0]), int(parts[1]), int(parts[2][:2]) if len(parts) > 2 else 0)
                else:
                    result = end_time
                rs.Close()
                conn.Close()
                return result
            
            # If no slot contains this exact time, find the next slot that starts after
            sql2 = f"""
                SELECT TOP 1 start_time, end_time
                FROM [castro_services].dbo.Time_Slots
                WHERE slot_date = '{date_str}'
                AND start_time > '{time_str}'
                ORDER BY start_time ASC
            """
            
            rs.Open(sql2, conn)
            
            if not rs.EOF:
                end_time = rs.Fields('end_time').Value
                if isinstance(end_time, str):
                    parts = end_time.split(':')
                    result = time(int(parts[0]), int(parts[1]), int(parts[2]) if len(parts) > 2 else 0)
                else:
                    result = end_time
                rs.Close()
                conn.Close()
                return result
            
            rs.Close()
            conn.Close()
            return None
            
        except Exception as e:
            messagebox.showinfo("Erro:", f"Error getting next free time: {e}")
            print(f"Error getting next free time: {e}")
            return None

    def _has_time_slot_conflict(self, check_date, start_time, end_time):
        """Check if proposed time conflicts with existing Time_Slots"""
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            rs = win32com.client.Dispatch("ADODB.Recordset")
            
            conn.Open(self.str_conn)
            
            # Convert times to string format for SQL comparison
            start_str = start_time.strftime('%H:%M:%S') if isinstance(start_time, time) else str(start_time)
            end_str = end_time.strftime('%H:%M:%S') if isinstance(end_time, time) else str(end_time)
            date_str = check_date.strftime('%Y-%m-%d') if isinstance(check_date, date) else str(check_date)
            
            sql = f"""
                SELECT COUNT(*) as cnt
                FROM [castro_services].dbo.Time_Slots
                WHERE slot_date = '{date_str}'
                AND start_time < '{end_str}'
                AND end_time > '{start_str}'
            """
            
            rs.Open(sql, conn)
            
            count = rs.Fields('cnt').Value if not rs.EOF else 0
            rs.Close()
            conn.Close()
            
            return count > 0
            
        except Exception as e:
            print(f"Error checking conflicts: {e}")
            return True  # Assume conflict on error (safe side)


    def _get_availability(self):
        """Load Calendar_Availability into memory"""
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            rs = win32com.client.Dispatch("ADODB.Recordset")
            
            conn.Open(self.str_conn)
            
            sql = """
                SELECT day_of_week, start_time, end_time
                FROM [castro_services].dbo.Calendar_Availability
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
                    
                    # Convert to time objects if string
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

    def open_teams_chat(self):
        try:
            
            link = f"https://teams.cloud.microsoft/l/chat/19:fc9176a569904180bbfa3eeb1bd52651@thread.v2/conversations?context=%7B%22contextType%22%3A%22chat%22%7D"
            webbrowser.open(link)
                        
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao abrir chat:\n{str(e)}")

    def context_function_1(self):
        """Blank placeholder function 1"""
        pass
    
    def context_function_2(self):
        """Blank placeholder function 2"""
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
        #self.load_pending_list()
    def show_day_orders(self, calendar_key):
        
        popup = tk.Toplevel(self.frame_certificados)
        popup.geometry("850x400")
        popup.minsize(800, 400)
        popup.configure(bg=self.cor_fundo)
        popup.calendar_key = calendar_key

        if calendar_key not in self.calendar_data:
            return
        
        orders = self.calendar_data[calendar_key]
        date_str = f"{calendar_key[0]:02d}/{calendar_key[1]:02d}/{calendar_key[2]}"
        
        popup.title(f"Serviços - {date_str} (Laboratório de Pressão)")
        
        popup.schedule_ids = [o.get('schedule_id') for o in orders]
        # ===== TOP FRAME =====
        top_frame = tk.Frame(popup, bg=self.cor_fundo)
        top_frame.pack(fill='x', padx=10, pady=(10, 0))
        
        #tk.Label(top_frame, text=f"Serviços agendados para {date_str}",font=('Segoe UI', 10, 'bold'), bg=self.cor_fundo, fg='white').pack(side='left')
        
        # Treeview
        tree_frame = ttk.Frame(popup)
        tree_frame.pack(pady=10, padx=10, expand=True, fill='both')
        
        tree_scroll = ttk.Scrollbar(tree_frame)
        tree_scroll.pack(side='right', fill='y')
        
        columns = ("Horário", "OS", "Código", "Serviço", "Descrição", "Duração", "Status")
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings",
                            yscrollcommand=tree_scroll.set, selectmode="extended")
        tree_scroll.config(command=tree.yview)
        
        tree.heading("Horário", text="Horário")
        tree.heading("OS", text="OS")
        tree.heading("Código", text="Código")
        tree.heading("Serviço", text="Serviço")
        tree.heading("Descrição", text="Descrição")
        tree.heading("Duração", text="Duração")
        tree.heading("Status", text="Status")
        
        tree.column("Horário", width=100)
        tree.column("OS", width=80)
        tree.column("Código", width=70)
        tree.column("Serviço", width=150)
        tree.column("Descrição", width=150)
        tree.column("Duração", width=55)
        tree.column("Status", width=70)
        
        tree.pack(expand=True, fill='both')
        
        tree.tag_configure('urgente', background='#FFD9D9')
        tree.tag_configure('normal', background='#FFFFFF')
        
        for order in orders:
            start_str = str(order['start_time'])[:5] if order['start_time'] else "--:--"
            end_str = str(order['end_time'])[:5] if order['end_time'] else "--:--"
            os_code = self.extract_os_from_notes(order.get('notes', ''))
            
            is_urgent = order.get('priority', 10) <= 5
            tag = 'urgente' if is_urgent else 'normal'
            
            tree.insert("", "end", values=(
                f"{start_str} - {end_str}",
                os_code,
                order['code'],
                order['specification'],
                order['description'],
                f"{order['execution_time_minutes']} min",
                order['status']
            ), tags=(tag,))
        # Load and display exceptions for this date

        exc_date = datetime(calendar_key[2], calendar_key[1], calendar_key[0]).date()
        exceptions = self.load_exceptions_for_date(exc_date)

        if exceptions:
            exc_frame = tk.Frame(popup, bg=self.cor_fundo)
            exc_frame.pack(fill='x', padx=10, pady=(0, 5))
            
            tk.Label(exc_frame, text="Bloqueios de Agenda:", font=('Segoe UI', 9, 'bold'),bg=self.cor_fundo, fg='#FFD700').pack(anchor='w')
            
            for exc in exceptions:
                tk.Label(exc_frame, text=f"  ⛔ {exc}", font=('Segoe UI', 8),bg=self.cor_fundo, fg='#FF9999').pack(anchor='w')
                
        btn_frame = tk.Frame(popup, bg=self.cor_fundo)
        btn_frame.pack(fill='x', padx=10, pady=(5, 10))
        
        ttk.Button(btn_frame, text="Inciar chat do Teams com Laboratório de Pressão", command=self.open_teams_chat, cursor='hand2').pack(side='left', padx=5)
        
        # Context menu
        tree.bind('<Button-3>', lambda e, t=tree: self.show_popup_context_menu(e, t))

    def get_sharepoint_url(self, os_code):
        if not os_code:
            return ""
        os_year = os_code[-4:]
        os_parts = os_code.split('/')
        os_code_num = os_parts[0].lstrip('0')
        os_code_num = format_os_code(os_code_num)
        return f"https://sesirs.sharepoint.com/:f:/r/sites/gdms-ISISistemasdeSensoriamento/Documentos%20Compartilhados/ISI%20SIM%20-%20Metrologia/Atendimento%20ao%20Cliente/E%20-%20Ordens%20de%20Servi%C3%A7o/{os_year}/{os_code_num}"

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
        sharepoint_url = self.get_sharepoint_url(os_code)
        
        context_menu.add_command(label="Iniciar chat do Teams com Laboratório de Pressão", command=self.open_teams_chat)
        context_menu.add_separator()
        
        if sharepoint_url:
            context_menu.add_command(label="Abrir link do SharePoint no navegador", command=lambda u=sharepoint_url: webbrowser.open(u))
            context_menu.add_command(label="Copiar link SharePoint", command=lambda u=sharepoint_url: self.copy_to_clipboard(u) if hasattr(self, 'copy_to_clipboard') else None)
        
        context_menu.add_command(label="Abrir diretório da planilha de cálculo no Sharepoint", command=lambda: self.open_teams_link_planilha(None))
        
        if calendar_key and total_count > 0:
            context_menu.add_separator()
            
            # Reschedule SELECTED services
            if selected_count == 1:
                label = f"Reagendar o serviço selecionado"
            else:
                label = f"Reagendar os {selected_count} serviços selecionados"
            
            context_menu.add_command(label=label,command=lambda t=tree, s=selected: self.reschedule_selected_services(t, s))
            
            # Reschedule ALL services (only if selecting all)
            #if selected_count == total_count and total_count != 1:
            #    context_menu.add_command(label=f"Reagendar todos os {total_count} serviços deste dia",command=lambda k=calendar_key: self.reschedule_day_services(k))
        
        context_menu.post(event.x_root, event.y_root)

    def reschedule_selected_services(self, tree, selected_items):
        """Reschedule only selected services from the popup"""
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

        if count == 1:
            confirm = messagebox.askyesno("Info agendamento",f"Deseja reagendar o serviço selecionado de {date_str}?")
        if count > 1:
            confirm = messagebox.askyesno("Info agendamento",f"Deseja reagendar os {count} serviço selecionado de {date_str}?")
        
        if not confirm:
            return
        
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            conn.Open(self.str_conn)
            
            id_list = ','.join(str(sid) for sid in selected_schedule_ids)
            
            conn.Execute(f"""
                UPDATE [castro_services].dbo.Service_Schedule
                SET status = 'PENDING', scheduled_start = NULL, scheduled_end = NULL, updated_at = GETDATE()
                WHERE id IN ({id_list})
            """)
            
            conn.Execute(f"DELETE FROM [castro_services].dbo.Time_Slots WHERE schedule_id IN ({id_list})")
            conn.Close()
            
            reschedule_date = datetime(calendar_key[2], calendar_key[1], calendar_key[0]).date()
            self.schedule_all_pending(from_date=reschedule_date + timedelta(days=1))
            
            self.load_calendar_data()
            self.render_calendar()
            popup.destroy()
            
            if count == 1:
                self.status_label.config(text=f"Serviço reagendado para {self.next_start.strftime('%d/%m/%Y')} das {self.next_start.strftime('%H:%M:%S')} às {self.next_end.strftime('%H:%M:%S')}")
                messagebox.showinfo('Info agendamento',f'Serviço reagendado para {self.next_start.strftime('%d/%m/%Y')} das {self.next_start.strftime('%H:%M:%S')} às {self.next_end.strftime('%H:%M:%S')}')
            if count > 1:
                self.status_label.config(text=f"{count} serviços reagendados para a partir de {self.next_start.strftime('%d/%m/%Y')}")
                messagebox.showinfo('Info agendamento',f'Serviço reagendado para {self.next_start.strftime('%d/%m/%Y')} das {self.next_start.strftime('%H:%M:%S')} às {self.next_end.strftime('%H:%M:%S')}')

            #self.status_label.config(text=f"✅ {count} serviço(s) reagendado(s)!")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao reagendar:\n{str(e)}")

    def open_teams_link_planilha(self, item):
        try:
            link = 'https://sesirs.sharepoint.com/:f:/r/sites/gdms-ISISistemasdeSensoriamento/Documentos%20Compartilhados/ISI%20SIM%20-%20Metrologia/Relat%C3%B3rios%20e%20Registros/Laborat%C3%B3rio%20de%20Press%C3%A3o'
            
            webbrowser.open(link)
                
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao abrir link de diretório:\n\n{str(e)}")


    def load_services(self):
        """Load available services from database"""
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            rs = win32com.client.Dispatch("ADODB.Recordset")
            
            conn.Open(self.str_conn)
            
            sql = """
                SELECT id, code, specification, execution_time_minutes, description
                FROM [castro_services].dbo.Service_Modes_Local
                --WHERE is_active = 1
                ORDER BY code asc
            """
            
            rs.Open(sql, conn)
            
            self.services_data = []
            service_list = []
            
            if not rs.EOF:
                rs.MoveFirst()
                while not rs.EOF:
                    service_id = rs.Fields('id').Value
                    service_code = rs.Fields('code').Value
                    spec = rs.Fields('specification').Value
                    desc = rs.Fields('description').Value
                    #duration = rs.Fields('execution_time_minutes').Value
                    duration = rs.Fields('execution_time_minutes').Value
                    
                    self.services_data.append({
                        'id': service_id,
                        'code': service_code,
                        'specification': spec,
                        'description': desc,
                        'execution_time_minutes': duration
                    })
                    
                    service_list.append(f"{service_code} - {spec} ({duration} min)")
                    rs.MoveNext()
            
            rs.Close()
            conn.Close()
            
        except Exception as e:
            self.status_label.config(text=f"Erro ao carregar serviços: {str(e)}")
    
    def clear_all_scheduled(self):
        """Clear all scheduled services completely"""
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            conn.Open(self.str_conn)
            
            conn.Execute("DELETE FROM [castro_services].dbo.Time_Slots")
            conn.Execute("DELETE FROM [castro_services].dbo.Service_Schedule")
            conn.Execute("DBCC CHECKIDENT ('[castro_services].dbo.Service_Schedule', RESEED, 0)")
            
            conn.Close()
            
            self.load_calendar_data()
            self.render_calendar()
            self.update_queue_count()
            if hasattr(self, 'load_pending_list'):
                self.load_pending_list()
            
            self.status_label.config(text="Os registros da agenda foram removidos")
            messagebox.showinfo("Sucesso", "Todos os agendamentos foram removidos!")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao limpar agenda:\n{str(e)}")

    def schedule_all(self):
        """Process all pending services using FIFO"""
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            conn.Open(self.str_conn)
            
            rs = win32com.client.Dispatch("ADODB.Recordset")
            rs.Open("SELECT COUNT(*) as cnt FROM [castro_services].dbo.Service_Schedule WHERE status = 'PENDING'", conn)
            
            pending_count = rs.Fields('cnt').Value if not rs.EOF else 0
            rs.Close()
            
            if pending_count == 0:
                messagebox.showinfo("Aviso", "Nenhum serviço pendente na fila!")
                conn.Close()
                return
            
            conn.Execute("EXEC [castro_services].dbo.sp_ScheduleNextServices @batch_size = 1000")
            conn.Close()
            
            if pending_count == 1:
                self.status_label.config(text=f"{pending_count} serviço agendado")
            else:
                self.status_label.config(text=f"{pending_count} serviços agendados")
                
            self.load_calendar_data()
            self.render_calendar()
            self.update_queue_count()
            self.load_pending_list()
            
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao agendar:\n\n{str(e)}")
    
    def update_queue_count(self):
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            conn.Open(self.str_conn)
            
            rs = win32com.client.Dispatch("ADODB.Recordset")
            rs.Open("SELECT COUNT(*) as cnt FROM [castro_services].dbo.Service_Schedule WHERE status = 'PENDING'", conn)
            
            count = rs.Fields('cnt').Value if not rs.EOF else 0
            rs.Close()
            conn.Close()
            
            #self.lbl_queue.config(text=f"Fila: {count} pendente(s)")
            
        except Exception:
            pass

class ServiceSchedulerLinkedDirect:
    """ ABA: Serviços para Agendamento (Linked Server - 631CP) """
    
    def __init__(self, parent_notebook, str_conn, str_conn_primary,cor_fundo):
        self.str_conn = str_conn
        self.str_conn_primary = str_conn_primary
        self.cor_fundo = cor_fundo
        self.service_scheduler = ServiceScheduler(notebook, STR_CONN, self.cor_fundo)   ### ZEBRA
        self.frame_certificados = ttk.Frame(parent_notebook)
        parent_notebook.add(self.frame_certificados, text=" Serviços a Agendar ")

        parent_notebook.bind('<<NotebookTabChanged>>', self.on_tab_changed)
    
        self.selected_services = []
        
        self.setup_ui()
        self.load_services()
        
    def add_calendar_exception(self):
        
        popup = tk.Toplevel(self.frame_certificados)
        popup.geometry("400x300")
        popup.minsize(400, 350)
        popup.configure(bg='black')
        popup.title("Adicionar Bloqueio ao Calendário")
        
        # Exception Date
        tk.Label(popup, text="Data:", font=('Segoe UI', 9), bg='black', fg='white').pack(anchor='w', padx=20, pady=(15, 0))
        
        date_frame = tk.Frame(popup, bg='black')
        date_frame.pack(fill='x', padx=20, pady=(2, 5))
        
        self.exc_day_var =      tk.StringVar(value=datetime.now().strftime('%d'))
        self.exc_month_var =    tk.StringVar(value=datetime.now().strftime('%m'))
        self.exc_year_var =     tk.StringVar(value=datetime.now().strftime('%Y'))
        
        ttk.Entry(date_frame, textvariable=self.exc_day_var, width=3,background='black').pack(side='left')
        tk.Label(date_frame, text="/", bg=self.cor_fundo, fg='white',background='black').pack(side='left')
        ttk.Entry(date_frame, textvariable=self.exc_month_var, width=3,background='black').pack(side='left')
        tk.Label(date_frame, text="/", bg=self.cor_fundo, fg='white',background='black').pack(side='left')
        ttk.Entry(date_frame, textvariable=self.exc_year_var, width=5,background='black').pack(side='left')
        
        # Start Time
        tk.Label(popup, text="Horário de Início (HH:MM):", font=('Segoe UI', 9), bg='black', fg='white').pack(anchor='w', padx=20, pady=(10, 0))

        self.exc_start_var = tk.StringVar(value="")
        exc_start_entry = ttk.Entry(popup, textvariable=self.exc_start_var, width=10)
        exc_start_entry.pack(anchor='w', padx=20, pady=(2, 5))

        # End Time
        tk.Label(popup, text="Horário de Encerramento (HH:MM):", font=('Segoe UI', 9),bg='black', fg='white').pack(anchor='w', padx=20, pady=(5, 0))

        self.exc_end_var = tk.StringVar(value="")
        exc_end_entry = ttk.Entry(popup, textvariable=self.exc_end_var, width=10)
        exc_end_entry.pack(anchor='w', padx=20, pady=(2, 5))

        # Notes
        tk.Label(popup, text="Razão:", font=('Segoe UI', 9),bg='black', fg='white').pack(anchor='w', padx=20, pady=(5, 0))

        self.exc_notes_var = tk.StringVar(value="")
        ttk.Entry(popup, textvariable=self.exc_notes_var, width=40).pack(anchor='w', padx=20, pady=(2, 5))

        # Checkbox: Full day
        def toggle_time_fields():
            if self.exc_full_day.get():
                exc_start_entry.config(state='disabled')
                exc_end_entry.config(state='disabled')
            else:
                exc_start_entry.config(state='normal')
                exc_end_entry.config(state='normal')

        self.exc_full_day = tk.BooleanVar(value=False)
        ttk.Checkbutton(popup, text="Dia inteiro", variable=self.exc_full_day, command=toggle_time_fields).pack(anchor='w', padx=20, pady=(10, 5))

        # Buttons
        btn_frame = tk.Frame(popup, bg='black')
        btn_frame.pack(pady=15)
        
        ttk.Button(btn_frame, text=" Salvar ", command=lambda: self.save_exception(popup)).pack(side='left', padx=5)
        ttk.Button(btn_frame, text=" Cancelar ", command=popup.destroy).pack(side='left', padx=5)

    def save_exception(self, popup):
        """Save the new calendar exception to database"""
        try:
            day = self.exc_day_var.get().strip().zfill(2)
            month = self.exc_month_var.get().strip().zfill(2)
            year = self.exc_year_var.get().strip()
            
            exception_date_sql = f"{year}-{month}-{day}"
            exception_date_br = f"{day}/{month}/{year}"
            
            # Validate date
            datetime.strptime(exception_date_sql, '%Y-%m-%d')
            
            is_full_day = self.exc_full_day.get()
            start_time = None if is_full_day else self.exc_start_var.get().strip()
            end_time = None if is_full_day else self.exc_end_var.get().strip()
            notes = self.exc_notes_var.get().strip()
            
            if not notes:
                notes = "Indisponibilidade"
            
            # Validate times if not full day
            if not is_full_day:
                if not start_time or not end_time:
                    messagebox.showwarning("Aviso", "Informe horário de início e fim!")
                    return
                
                # Validate format HH:MM
                for t in [start_time, end_time]:
                    parts = t.split(':')
                    if len(parts) != 2:
                        messagebox.showwarning("Aviso", "Formato de horário inválido! Use HH:MM")
                        return
                    if not (0 <= int(parts[0]) <= 23 and 0 <= int(parts[1]) <= 59):
                        messagebox.showwarning("Aviso", "Horário inválido!")
                        return
                
                start_time_str = f"{start_time}:00"
                end_time_str = f"{end_time}:00"
            else:
                start_time_str = None
                end_time_str = None
            
            conn = win32com.client.Dispatch("ADODB.Connection")
            conn.Open(self.str_conn_primary)
            
            sql = f"""
                INSERT INTO [castro_services].dbo.Calendar_Exceptions
                    (exception_date, exception_type, start_time, end_time, is_available, notes)
                VALUES (
                    '{exception_date_sql}',
                    'MODIFIED_HOURS',
                    {f"'{start_time_str}'" if start_time_str else 'NULL'},
                    {f"'{end_time_str}'" if end_time_str else 'NULL'},
                    0,
                    '{notes.replace("'", "''")}'
                )
            """
            
            conn.Execute(sql)
            conn.Close()
            
            popup.destroy()
            
            # Build message text
            if is_full_day:
                msg = f"Bloqueio de agenda adicionado:\n\nData: {exception_date_br}\nDia inteiro\nRazão: {notes}"
            else:
                msg = f"Bloqueio de agenda adicionado:\n\nData: {exception_date_br}\nHorário: {start_time} às {end_time}\nRazão: {notes}"
            
            messagebox.showinfo("Info agendamento", msg)
            
            self.status_label.config(text=f"Bloqueio de agenda adicionado em: {exception_date_br} das {start_time} às {end_time}")
            #self.load_calendar_data()
            self.render_calendar()
            
        except ValueError:
            messagebox.showwarning("Aviso", "Data inválida!")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao salvar exceção:\n{str(e)}")

    #def load_calendar_data(self):
    #    """Load all scheduled services grouped by date"""
    #    self.calendar_data.clear()
    #    
    #    try:
    #        conn = win32com.client.Dispatch("ADODB.Connection")
    #        conn.Open(self.str_conn)
    #        
    #        sql = """
    #            SELECT 
    #                ts.slot_date,
    #                ts.start_time,
    #                ts.end_time,
    #                s.code,
    #                s.specification,
    #                s.description,
    #                s.execution_time_minutes,
    #                ss.notes,
    #                ss.status,
    #                ss.priority,
    #                ss.id as schedule_id
    #            FROM [castro_services].dbo.Time_Slots ts
    #            JOIN [castro_services].dbo.Service_Schedule ss ON ts.schedule_id = ss.id
    #            JOIN [castro_services].dbo.Service_Modes_Local s ON ts.service_id = s.id
    #            ORDER BY ts.slot_date, ts.start_time
    #        """
    #        
    #        rs = win32com.client.Dispatch("ADODB.Recordset")
    #        rs.Open(sql, conn)
    #        
    #        if not rs.EOF:
    #            rs.MoveFirst()
    #            while not rs.EOF:
    #                slot_date = rs.Fields('slot_date').Value
    #                
    #                if isinstance(slot_date, str):
    #                    slot_date = datetime.strptime(slot_date, '%Y-%m-%d').date()
    #                elif isinstance(slot_date, datetime):
    #                    slot_date = slot_date.date()
    #                
    #                calendar_key = (slot_date.day, slot_date.month, slot_date.year)
#
    #                order_data = {
    #                'start_time': rs.Fields('start_time').Value,
    #                'end_time': rs.Fields('end_time').Value,
    #                'code': rs.Fields('code').Value,
    #                'specification': rs.Fields('specification').Value,
    #                'description': rs.Fields('description').Value,
    #                'execution_time_minutes': rs.Fields('execution_time_minutes').Value,
    #                'notes': rs.Fields('notes').Value,
    #                'status': rs.Fields('status').Value,
    #                'priority': rs.Fields('priority').Value,
    #                'schedule_id': rs.Fields('schedule_id').Value  # ADD THIS LINE
    #            }
    #                
    #                if calendar_key not in self.calendar_data:
    #                    self.calendar_data[calendar_key] = []
    #                self.calendar_data[calendar_key].append(order_data)
    #                
    #                rs.MoveNext()
    #        
    #        rs.Close()
    #        conn.Close()
    #        
    #    except Exception as e:
    #        print(f"Error loading calendar data: {e}")
    
    def find_next_available_slot(self, from_time, duration_minutes):
        """
        Find next available time slot starting from `from_time`
        Replicates sp_FindNextAvailableSlot logic in Python
        
        Returns: (start_datetime, end_datetime) or (None, None) if no slot found
        """
        check_date = from_time.date()
        max_days = 365
        
        # Load all data once
        availability = self._get_availability()
        exceptions = self._get_exceptions()
        
        for day_offset in range(max_days):
            sql_day_of_week = (check_date.weekday() + 1) % 7
            
            # Check if this day has any availability windows
            day_windows = [w for w in availability if w['day_of_week'] == sql_day_of_week]
            
            if not day_windows:
                check_date += timedelta(days=1)
                from_time = datetime.combine(check_date, time(8, 0))
                continue
            
            # Check for full-day exception (holiday)
            full_day_blocked = False
            for ex in exceptions:
                if ex['exception_date'] == check_date and ex['is_available'] == False and ex['start_time'] is None:
                    full_day_blocked = True
                    break
            
            if full_day_blocked:
                check_date += timedelta(days=1)
                from_time = datetime.combine(check_date, time(8, 0))
                continue
            
            # Try each availability window
            for window in sorted(day_windows, key=lambda w: w['start_time']):
                window_start_dt = datetime.combine(check_date, window['start_time'])
                window_end_dt = datetime.combine(check_date, window['end_time'])
                
                # Adjust start if from_time is later
                proposed_start = max(from_time, window_start_dt)
                
                # TRY MULTIPLE POSITIONS WITHIN THIS WINDOW
                while proposed_start + timedelta(minutes=duration_minutes) <= window_end_dt:
                    proposed_end = proposed_start + timedelta(minutes=duration_minutes)
                    
                    # Check for time-specific exceptions (lunch break)
                    blocked_by_exception = False
                    for ex in exceptions:
                        if (ex['exception_date'] == check_date and 
                            ex['is_available'] == False and 
                            ex['start_time'] is not None and 
                            ex['end_time'] is not None):
                            
                            ex_start_dt = datetime.combine(check_date, ex['start_time'])
                            ex_end_dt = datetime.combine(check_date, ex['end_time'])
                            
                            if proposed_start < ex_end_dt and proposed_end > ex_start_dt:
                                blocked_by_exception = True
                                # Jump past the exception
                                proposed_start = ex_end_dt
                                break
                    
                    if blocked_by_exception:
                        continue  # Re-test with new proposed_start
                    
                    # Check for conflicts with existing Time_Slots
                    if self._has_time_slot_conflict(check_date, proposed_start.time(), proposed_end.time()):
                        # Jump to the end of the conflicting slot
                        next_free = self._get_next_free_time(check_date, proposed_start.time())
                        if next_free is None:
                            break  # No more free time in this window
                        proposed_start = datetime.combine(check_date, next_free)
                        continue  # Re-test with new proposed_start
                    
                    # Found a valid slot!
                    return proposed_start, proposed_end
                
                # No slot in this window, try next window
                # Reset proposed_start to beginning of next window or from_time
                pass
            
            # No slot found today, move to next day
            check_date += timedelta(days=1)
            if day_windows:
                earliest = min(w['start_time'] for w in day_windows)
                from_time = datetime.combine(check_date, earliest)
            else:
                from_time = datetime.combine(check_date, time(8, 0))
        
        return None, None


    def _get_next_free_time(self, check_date, after_time):
        """
        Get the end time of the conflicting slot that starts at or contains `after_time`.
        Returns the end time as a time object, or None if at end of day.
        """
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            rs = win32com.client.Dispatch("ADODB.Recordset")
            
            conn.Open(self.str_conn_primary)
            
            date_str = check_date.strftime('%Y-%m-%d')
            time_str = after_time.strftime('%H:%M:%S') if isinstance(after_time, time) else str(after_time)
            
            sql = f"""
                SELECT TOP 1 end_time
                FROM [castro_services].dbo.Time_Slots
                WHERE slot_date = '{date_str}'
                AND start_time <= '{time_str}'
                AND end_time > '{time_str}'
                ORDER BY end_time ASC
            """
            
            rs.Open(sql, conn)
            
            if not rs.EOF:
                end_time = rs.Fields('end_time').Value
                if isinstance(end_time, str):
                    parts = end_time.split(':')
                    result = time(int(parts[0]), int(parts[1]), int(parts[2][:2]) if len(parts) > 2 else 0)
                else:
                    result = end_time
                rs.Close()
                conn.Close()
                return result
            
            # If no slot contains this exact time, find the next slot that starts after
            sql2 = f"""
                SELECT TOP 1 start_time, end_time
                FROM [castro_services].dbo.Time_Slots
                WHERE slot_date = '{date_str}'
                AND start_time > '{time_str}'
                ORDER BY start_time ASC
            """
            
            rs.Open(sql2, conn)
            
            if not rs.EOF:
                end_time = rs.Fields('end_time').Value
                if isinstance(end_time, str):
                    parts = end_time.split(':')
                    result = time(int(parts[0]), int(parts[1]), int(parts[2]) if len(parts) > 2 else 0)
                else:
                    result = end_time
                rs.Close()
                conn.Close()
                return result
            
            rs.Close()
            conn.Close()
            return None
            
        except Exception as e:
            messagebox.showinfo("Erro:", f"Error getting next free time: {e}")
            print(f"Error getting next free time: {e}")
            return None

    def _get_availability(self):
        """Load Calendar_Availability into memory"""
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            rs = win32com.client.Dispatch("ADODB.Recordset")
            
            conn.Open(self.str_conn_primary)
            
            sql = """
                SELECT day_of_week, start_time, end_time
                FROM [castro_services].dbo.Calendar_Availability
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
                    
                    # Convert to time objects if string
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

    def _get_exceptions(self):
        """Load Calendar_Exceptions into memory"""
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            rs = win32com.client.Dispatch("ADODB.Recordset")
            
            conn.Open(self.str_conn_primary)
            
            sql = """
                SELECT exception_date, exception_type, start_time, end_time, is_available
                FROM [castro_services].dbo.Calendar_Exceptions
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


    def _has_time_slot_conflict(self, check_date, start_time, end_time):
        """Check if proposed time conflicts with existing Time_Slots"""
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            rs = win32com.client.Dispatch("ADODB.Recordset")
            
            conn.Open(self.str_conn_primary)
            
            # Convert times to string format for SQL comparison
            start_str = start_time.strftime('%H:%M:%S') if isinstance(start_time, time) else str(start_time)
            end_str = end_time.strftime('%H:%M:%S') if isinstance(end_time, time) else str(end_time)
            date_str = check_date.strftime('%Y-%m-%d') if isinstance(check_date, date) else str(check_date)
            
            sql = f"""
                SELECT COUNT(*) as cnt
                FROM [castro_services].dbo.Time_Slots
                WHERE slot_date = '{date_str}'
                AND start_time < '{end_str}'
                AND end_time > '{start_str}'
            """
            
            rs.Open(sql, conn)
            
            count = rs.Fields('cnt').Value if not rs.EOF else 0
            rs.Close()
            conn.Close()
            
            return count > 0
            
        except Exception as e:
            print(f"Error checking conflicts: {e}")
            return True  # Assume conflict on error (safe side)

    def process_fifo(self):
        
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            conn.Open(self.str_conn_primary)
            
            # Get pending services in FIFO order
            rs = win32com.client.Dispatch("ADODB.Recordset")
            sql = """
                SELECT 
                    ss.id as schedule_id,
                    ss.service_id,
                    s.execution_time_minutes,
                    ss.priority,
                    ss.requested_at,
                    ss.notes
                FROM [castro_services].dbo.Service_Schedule ss
                JOIN [castro_services].dbo.Service_Modes_Local s ON ss.service_id = s.id
                WHERE ss.status = 'PENDING'
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
                        'priority': rs.Fields('priority').Value,
                        'requested_at': rs.Fields('requested_at').Value,
                        'notes': rs.Fields('notes').Value
                    })
                    rs.MoveNext()
            
            rs.Close()
            
            pending_count = len(pending_services)
            
            if pending_count == 0:
                messagebox.showinfo("Aviso", "Nenhum serviço pendente na fila!")
                conn.Close()
                return
            
            # Start from now
            current_time = datetime.now()
            scheduled_count = 0
            
            for service in pending_services:
                # Find next available slot
                next_start, next_end = self.find_next_available_slot(
                    current_time, 
                    service['execution_time_minutes']
                )
                
                if next_start is None:
                    print(f"WARNING: Could not schedule service {service['schedule_id']}")
                    continue
                
                # Update Service_Schedule
                update_sql = f"""
                    UPDATE [castro_services].dbo.Service_Schedule
                    SET scheduled_start = '{next_start.strftime('%Y-%m-%d %H:%M:%S')}',
                        scheduled_end = '{next_end.strftime('%Y-%m-%d %H:%M:%S')}',
                        status = 'SCHEDULED',
                        updated_at = GETDATE()
                    WHERE id = {service['schedule_id']}
                """
                conn.Execute(update_sql)
                
                # Insert into Time_Slots
                insert_sql = f"""
                    INSERT INTO [castro_services].dbo.Time_Slots 
                        (slot_date, start_time, end_time, schedule_id, service_id)
                    VALUES (
                        '{next_start.strftime('%Y-%m-%d')}',
                        '{next_start.strftime('%H:%M:%S')}',
                        '{next_end.strftime('%H:%M:%S')}',
                        {service['schedule_id']},
                        {service['service_id']}
                    )
                """
                conn.Execute(insert_sql)
                
                # Advance pointer
                current_time = next_end
                scheduled_count += 1
            
            conn.Close()
            
            if scheduled_count == 1:
                messagebox.showinfo("Info agendamento", 
                    f"{scheduled_count} serviço agendado em Laboratório de Pressão!")
            else:
                messagebox.showinfo("Info agendamento", 
                    f"{scheduled_count} serviços agendados em Laboratório de Pressão!")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao processar fila:\n{str(e)}")

    def on_tab_changed(self, event):
        notebook = event.widget
        current_tab = notebook.select()
        current_tab_text = notebook.tab(current_tab, "text")

        #self.service_scheduler.refresh_calendar()

        if "Serviços Vinculados" in current_tab_text or "Agendar Serviços" in current_tab_text:
            self.load_services()

    def setup_ui(self):
        main_container = tk.Frame(self.frame_certificados, bg=self.cor_fundo)
        main_container.pack(fill='both', expand=True, padx=10, pady=10)
        
        # ===== TOP FRAME - Info & Controls =====
        top_frame = tk.Frame(main_container, bg=self.cor_fundo)
        top_frame.pack(fill='x', pady=(0, 10))
        
        # Greeting
        self.greetings = ttk.Label(top_frame, text=f"{buenas}", font=("Segoe UI", 10),background='black')
        self.greetings.pack(side='left', padx=5)
        
        # Title
        tk.Label(top_frame, text="Serviços de Pressão no setor Qualidade",font=('Segoe UI', 12, 'bold'), bg=self.cor_fundo, fg='white').pack(side='left', padx=20)
        
        # Refresh button
        #btn_refresh = tk.Button(top_frame, text="🔄 Atualizar Lista",font=('Segoe UI', 9), bg='#17A2B8', fg='white',padx=10, pady=2, cursor='hand2',command=self.load_services)
        #btn_refresh.pack(side='right', padx=5)
        
        #btn_refresh.config(state="disabled")
        
        # Selection count
        self.lbl_count = tk.Label(top_frame, text="0 selecionados", font=('Segoe UI', 9, 'bold'), bg=self.cor_fundo, fg='#FFD700')
        self.lbl_count.pack(side='right', padx=20)
        
        search_frame = tk.Frame(main_container, bg=self.cor_fundo)
        search_frame.pack(fill='x', pady=(0, 5))
        
        tk.Label(search_frame, text="Buscar:", font=('Segoe UI', 9), bg=self.cor_fundo, fg='white').pack(side='left', padx=(0, 5))
        
        self.search_var = tk.StringVar()
        self.search_var.trace('w', lambda *args: self.filter_treeview())
        
        self.search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=40)
        self.search_entry.pack(side='left', padx=5)
        
        # ===== TREEVIEW FRAME =====
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
        
        column_widths = {
            "OS": 80,
            "Item": 120,
            "Especificação": 200,
            "Descrição": 250,
            "Código": 100,
            "Duração": 70
        }
        
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=column_widths.get(col, 80), minwidth=50)
        
        self.tree.pack(expand=True, fill='both')
        
        # Bind events
        self.tree.bind('<Button-3>', self.show_context_menu)
        self.tree.bind('<<TreeviewSelect>>', self.on_tree_select)
        self.tree.bind('<Control-a>', self.select_all)
        self.tree.bind('<Control-A>', self.select_all)
        
        # ===== BUTTON FRAME =====
        button_frame = tk.Frame(main_container, bg=self.cor_fundo)
        button_frame.pack(fill='x', pady=(10, 0))

        # Left side buttons
        
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

        # Right side - Process FIFO button
        #self.btn_process_fifo = ttk.Button(button_frame, text=" ⚙️ Processar Fila (FIFO) ", command=self.process_fifo, cursor='hand2')
        #self.btn_process_fifo.pack(side="right", padx=5)

        # Process all pending button
        #self.btn_process = tk.Button(button_frame, text="⚙️ Processar Fila FIFO",font=('Segoe UI', 10, 'bold'),bg='#1B4B9F', fg='white', padx=15, pady=6,cursor='hand2', command=self.process_fifo)
        #self.btn_process.pack(side='right', padx=5)
        
        #self.btn_process = ttk.Button(button_frame, text="Processar Fila (FIFO)", command=self.process_fifo, cursor='hand2').pack(side='right',padx=5)

        # ===== STATUS BAR =====
        self.status_label = ttk.Label(self.frame_certificados,text="Selecione os serviços e use o botão direito para agendar",font=("Segoe UI", 12))
        self.status_label.pack(pady=5)
        
        # Store all rows for filtering
        self.all_rows = []
    
    def load_services(self):
        """Load services from linked server"""
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        self.all_rows.clear()
        
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            rs = win32com.client.Dispatch("ADODB.Recordset")
            
            conn.Open(self.str_conn)

            ###VISUALIZA SERVIÇOS NO SETOR, HERDA DADOS DO MYLOGICAL

            sql = """

                select 

                os.code as 'OS',
                i.code as 'Item',
                sm.specification as 'Especificação',
                sm.description as 'Descrição',
                sm.code as 'Code',
                sm.execution_time as 'execution_time'

                from 

                instruments_services iss

                left join orders_services as os on os.id = iss.id_service_order 
                left join service_modes as sm on sm.id = iss.id_service
                left join instruments as i on i.id = iss.id_instrument and i.id_service_order = os.id
                left join budgets as b on b.id_order_service = os.id

                where os.removed = 0
                and iss.removed = 0
                and sm.removed = 0
                and i.removed = 0
                and i.id_current_sector = 17
                and sm.code like '%631cp%'
                and b.is_last_revision = 1

                order by os.receiving_date asc
            """
            
            rs.Open(sql, conn)
            
            if not rs.EOF:
                rs.MoveFirst()
                while not rs.EOF:
                    execution_time_raw = rs.Fields('execution_time').Value
                    duration_minutes = self.convert_time_to_minutes(execution_time_raw) if execution_time_raw else 0
                    
                    os_code = rs.Fields('OS').Value if rs.Fields('OS').Value else ""
                    os_year = os_code[-4:]
                    os_parts = os_code.split('/')
                    os_code_num =f"{os_parts[0].lstrip('0')}/{os_year}"
                    #os_code_num = format_os_code(os_code_num)}

                    row_data = {
                        #'OS': rs.Fields('OS').Value if rs.Fields('OS').Value else "",
                        'OS': os_code_num,
                        'Item': rs.Fields('Item').Value if rs.Fields('Item').Value else "",
                        'Especificação': rs.Fields('Especificação').Value if rs.Fields('Especificação').Value else "",
                        'Descrição': rs.Fields('Descrição').Value if rs.Fields('Descrição').Value else "",
                        'Code': rs.Fields('Code').Value if rs.Fields('Code').Value else "",
                        'execution_time_raw': str(execution_time_raw) if execution_time_raw else "00:00",
                        'duration_minutes': duration_minutes,
                        #'instrument_service_id': rs.Fields('instrument_service_id').Value
                    }
                    
                    self.all_rows.append(row_data)
                    
                    self.tree.insert("", "end", values=(
                        row_data['OS'],
                        row_data['Item'],
                        row_data['Especificação'],
                        row_data['Descrição'],
                        row_data['Code'],
                        #f"{row_data['execution_time_raw']} ({duration_minutes} min)"
                        f"{duration_minutes} min"
                    ))
                    
                    rs.MoveNext()
            
            rs.Close()
            conn.Close()
            
            count = len(self.all_rows)
            if count == 0:
                self.status_label.config(text="Nenhum serviço 631CP encontrado no setor 1")
            elif count == 1:
                self.status_label.config(text=f"1 serviço carregado")
            else:
                self.status_label.config(text=f"{count} serviços carregados")
            
        except Exception as e:
            self.status_label.config(text=f"Erro: {str(e)}")
    
    def convert_time_to_minutes(self, time_str):
        """Convert HH:MM string to minutes"""
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
        """Filter treeview based on search entry"""
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
                    row['Descrição'], row['Code'],
                    f"{row['duration_minutes']} min"
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
            context_menu.add_command(label=f"Agendar serviço",command=self.schedule_selected)
            
        if count > 1:
            context_menu.add_command(label=f"Agendar serviços",command=self.schedule_selected)
        
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
        """Schedule all selected services"""
        selected = self.tree.selection()
        
        if not selected:
            messagebox.showwarning("Aviso", "Selecione pelo menos um serviço!")
            return
        
        # Get selected rows data
        selected_indices = []
        all_items = self.tree.get_children()
        
        for item in selected:
            idx = all_items.index(item)
            if idx < len(self.all_rows):
                selected_indices.append(idx)
                
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            conn.Open(self.str_conn_primary)
            
            added_count = 0
            for idx in selected_indices:
                row = self.all_rows[idx]
                #service_mode_id = row['service_mode_id']
                duration = row['duration_minutes']
                
                # Sync to local Services table if not exists
                #print('self.str_conn_primary\n\n',self.str_conn_primary)
                sync_sql = f"""
                    IF NOT EXISTS (SELECT 1 FROM [castro_services].dbo.Service_Modes_Local WHERE code = '{row['Code']}')
                    BEGIN
                        INSERT INTO [castro_services].dbo.Service_Modes_Local 
                            (code, specification, description, execution_time_minutes)
                        VALUES (
                            '{row['Code']}',
                            '{str(row['Especificação']).replace("'", "''")}',
                            '{str(row['Descrição']).replace("'", "''")}',
                            {duration}
                        );
                    END
                    ELSE
                    BEGIN
                        UPDATE [castro_services].dbo.Service_Modes_Local
                        SET execution_time_minutes = {duration},
                            specification = '{str(row['Especificação']).replace("'", "''")}'
                            --updated_at = GETDATE()
                        WHERE code = '{row['Code']}';
                    END
                """
                conn.Execute(sync_sql)
                
                # Get the local service ID
                rs = win32com.client.Dispatch("ADODB.Recordset")
                rs.Open(f"SELECT id FROM [castro_services].dbo.Service_Modes_Local WHERE code = '{row['Code']}'", conn)
                
                local_service_id = rs.Fields('id').Value if not rs.EOF else None
                rs.Close()
                
                if local_service_id:
                    # Add to schedule queue
                    notes = f"OS {row['OS']} - {row['Item']} - {row['Especificação']}"
                    schedule_sql = f"""
                        INSERT INTO [castro_services].dbo.Service_Schedule 
                            (service_id, notes)
                        VALUES ({local_service_id}, '{notes.replace("'", "''")}')
                    """
                    conn.Execute(schedule_sql)
                    added_count += 1
            
            conn.Close()
            
            if added_count == 1:
                self.status_label.config(text=f"{added_count} serviço adicionado à fila!")
            if added_count == 0:
                self.status_label.config(text=f"Nenhum serviço adicionado à fila!")
            if added_count > 1:
                self.status_label.config(text=f"{added_count} serviços adicionados à fila!")

            self.clear_selection()
            self.process_fifo()
            
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao agendar:\n{str(e)}")
 
class DatabaseViewer6: ### VISTA GERAL

    def __init__(self, parent_notebook, str_conn, cor_fundo):   
        #print('str_conn',str_conn) 
        self.str_conn = str_conn    
        self.cor_fundo = cor_fundo  
        self.selected_items = []    
        self.frame_certificados = ttk.Frame(parent_notebook)    
        parent_notebook.add(self.frame_certificados, text=" Vista Geral de Ordens de Serviço ") 
        
        self.teams_chat_ids = {                 
                    'Laboratório de Dimensional':                   '461bebb602cc4c2f997670af64f1bc50',
                    'Laboratório de Massa':                         '474353252d224d7caf749a4b5301c4d8',
                    'Laboratório de Pressão':                       'fc9176a569904180bbfa3eeb1bd52651',
                    'Laboratório de Força, Torque e Dureza':        'fc9176a569904180bbfa3eeb1bd52651',
                    'Laboratório de Volume e Massa Específica':     '34b79fe77c3e41b7a3030c4f010de73e',
                    'Laboratório de Metrologia por Coordenadas':    '716ff922fa584a2582ecb48e509edc2e',
                    'Laboratório de Temperatura e Umidade':         '680cba2db76043939995da4b21cfd11c',
                    'Laboratório de Vazão':                         'e04abb13e3114d95b399290fe371a391',
                    'Laboratório de Eletricidade':                  '680cba2db76043939995da4b21cfd11c',
                    'Laboratório de Tempo e Frequência':            '680cba2db76043939995da4b21cfd11c',
                    'Laboratório de Físico-Química':                '34b79fe77c3e41b7a3030c4f010de73e',
                    'Revisão de Relatórios':                        '7e40040f299041f799e60f152711f017',
                    'Relatórios Aguardando Assinaturas':            '7e40040f299041f799e60f152711f017',
                    'Expedição':                                    'missing',
                }
        
        # SILVIA NUNES 04/05/2026
        #'https://teams.microsoft.com/l/chat/19:461bebb602cc4c2f997670af64f1bc50@thread.v2/conversations?context=%7B%22contextType%22%3A%22chat%22%7D',
        #'https://teams.microsoft.com/l/chat/19:474353252d224d7caf749a4b5301c4d8@thread.v2/conversations?context=%7B%22contextType%22%3A%22chat%22%7D',
        #'https://teams.microsoft.com/l/chat/19:fc9176a569904180bbfa3eeb1bd52651@thread.v2/conversations?context=%7B%22contextType%22%3A%22chat%22%7D',
        #'https://teams.microsoft.com/l/chat/19:fc9176a569904180bbfa3eeb1bd52651@thread.v2/conversations?context=%7B%22contextType%22%3A%22chat%22%7D',
        #'https://teams.microsoft.com/l/chat/19:34b79fe77c3e41b7a3030c4f010de73e@thread.v2/conversations?context=%7B%22contextType%22%3A%22chat%22%7D',
        #'716ff922fa584a2582ecb48e509edc2e',
        #'https://teams.microsoft.com/l/chat/19:680cba2db76043939995da4b21cfd11c@thread.v2/conversations?context=%7B%22contextType%22%3A%22chat%22%7D',
        #'https://teams.microsoft.com/l/chat/19:e04abb13e3114d95b399290fe371a391@thread.v2/conversations?context=%7B%22contextType%22%3A%22chat%22%7D',
        #'',
        #'https://teams.microsoft.com/l/chat/19:680cba2db76043939995da4b21cfd11c@thread.v2/conversations?context=%7B%22contextType%22%3A%22chat%22%7D',
        #'https://teams.microsoft.com/l/chat/19:680cba2db76043939995da4b21cfd11c@thread.v2/conversations?context=%7B%22contextType%22%3A%22chat%22%7D',
        #'https://teams.microsoft.com/l/chat/19:34b79fe77c3e41b7a3030c4f010de73e@thread.v2/conversations?context=%7B%22contextType%22%3A%22chat%22%7D',

        self.setup_ui_agenda_os()
    
    def on_entry_focus_in(self, event):
        if self.search_entry.get() == self.placeholder_text:
            self.search_entry.delete(0, 'end')
            self.search_entry.configure(style='TEntry')

    def on_entry_focus_out(self, event):
        if not self.search_entry.get():
            self.search_entry.insert(0, self.placeholder_text)
            style = ttk.Style()
            self.search_entry.configure(style='Placeholder.TEntry')

    def setup_ui_agenda_os(self):
        search_frame = ttk.Frame(self.frame_certificados)
        search_frame.pack(pady=10, padx=10, fill="x")

        #ttk.Label(search_frame, text="Buscar:").pack(side="left", padx=(0, 15))
        #self.search_entry = ttk.Entry(search_frame, width=30)
        #self.search_entry.pack(side="left", padx=5)
        #self.search_entry.bind('<Return>', lambda e: self.perform_search())

        #ttk.Label(search_frame, text="Buscar:").pack(side="left", padx=(0, 15))
        
        #greeting_frame = ttk.Frame(search_frame)
        #greeting_frame.pack(fill='x', pady=(0, 5))
        
        self.greetings = ttk.Label(search_frame, text=f"{buenas}", font=("Segoe UI", 10),background='black')
        self.greetings.pack(side="left", padx=5, pady=(0, 0))
        
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
        
        self.buscar = ttk.Button(search_frame, text="Buscar", command=self.perform_search).pack(side="left", padx=5)
        self.mostrar_todos = ttk.Button(search_frame, text="Mostrar Todos", command=self.load_all_records).pack(side="left", padx=5)
                
        tree_frame = ttk.Frame(self.frame_certificados)
        tree_frame.pack(pady=10, padx=10, expand=True, fill="both")
        
        tree_scroll_y = ttk.Scrollbar(tree_frame)
        tree_scroll_y.pack(side="right", fill="y")
        
        tree_scroll_x = ttk.Scrollbar(tree_frame, orient="horizontal")
        tree_scroll_x.pack(side="bottom", fill="x")
        
        columns = ("OS",
                    "Recebimento", 
                    "Entrega", 
                    "Status",
                    "Item",
                    "Setor"
                )
        
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set, selectmode="extended")
        
        tree_scroll_y.config(command=self.tree.yview)
        tree_scroll_x.config(command=self.tree.xview)
        
        column_widths3 = {
            "OS":                   20,
            "Recebimento":          20,
            "Entrega":              40,
            "Item":                 40,
            "Status":               20,
            "Setor":                40
        }
        
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=column_widths3.get(col, 50), minwidth=20)
        
        self.tree.pack(expand=True, fill="both")
        
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
        
        self.btn_limpar = ttk.Button(button_frame, text="Ir para Hoje", command=self.scroll_to_today)
        self.btn_limpar.pack(side="left", padx=5)

        self.status_label = ttk.Label(self.frame_certificados, text="Pronto para consulta", font=("Segoe UI", 9))
        self.status_label.pack(pady=5)
        
        self.tree.tag_configure('atrasado',     background="#FFD9D9")    
        self.tree.tag_configure('hoje',         background="#B7D5F5")        
        self.tree.tag_configure('proximo',      background="#E2FFD3")  
        
        self.load_all_records()
        
    def show_context_menu(self, event):
        item = self.tree.identify_row(event.y)
        column = self.tree.identify_column(event.x)
        
        if item:
            self.tree.selection_set(item)
            
            context_menu = tk.Menu(self.tree, tearoff=0)
            
            values = self.tree.item(item)['values']
            sharepoint_url = values[0] if values and len(values) > 0 else ""
            setor_teams = values[5]
            if sharepoint_url:
                context_menu.add_command(label="Abrir link do Sharepoint no navegador", command=lambda u=sharepoint_url: self.open_sharepoint_link(u))
                context_menu.add_command(label="Copiar link", command=lambda u=sharepoint_url: self.copy_to_clipboard(u))
                context_menu.add_command(label="Copiar linha", command=lambda i=item: self.copy_row_data(i))
                context_menu.add_separator()
                context_menu.add_command(label=f"Iniciar chat do Teams com {setor_teams}",command=lambda i=item: self.open_teams_chat(i))
                #context_menu.add_separator()
                #context_menu.add_command(label="Enviar para impressora Zebra [Em desenvolvimento]", command=lambda: self.send_to_printer(self.tree.item(item)['values']))
            
            context_menu.post(event.x_root, event.y_root)               
            context_menu.add_separator()            
            context_menu.post(event.x_root, event.y_root)

    def scroll_to_today(self, offset=0):
        target_item = None
        target_index = 0
        all_items = self.tree.get_children()
        
        for i, item in enumerate(all_items):
            values = self.tree.item(item)['values']
            status = values[3] if len(values) > 3 else ""
            if status == "Hoje":
                target_item = item 
                target_index = i
                break
        
        if not target_item:
            for i, item in enumerate(all_items):
                values = self.tree.item(item)['values']
                status = values[3] if len(values) > 3 else ""
                if status == "Atraso":
                    target_item = item
                    target_index = i
                    break
        
        if target_item:
            new_index = max(0, min(target_index + offset, len(all_items) - 1))
            target_item = all_items[new_index]
            
            # Select the item first
            self.tree.selection_set(target_item)
            self.tree.focus(target_item)
            
            # Then scroll it to the top
            self.tree.yview_moveto(new_index / len(all_items))

    def find_chat_id(self, chat_name):
        
        if not chat_name:
            return ""
        
        # First try exact match
        if chat_name in self.teams_chat_ids:
            return self.teams_chat_ids[chat_name]
        
        # Split the search term into words
        search_words = chat_name.lower().split()
        
        # Try to find a match by checking if all search words appear in any key
        for key, chat_id in self.teams_chat_ids.items():
            key_lower = key.lower()
            
            # Check if ALL search words are in the key (strong match)
            if all(word in key_lower for word in search_words):
                return chat_id
        
        # If still no match, try partial matching (at least one significant word)
        significant_words = [w for w in search_words if len(w) > 3]  # Skip short words like "de", "da", "por"
        
        if significant_words:
            for key, chat_id in self.teams_chat_ids.items():
                key_lower = key.lower()
                
                # Check if at least half of significant words match
                matches = sum(1 for word in significant_words if word in key_lower)
                if matches >= len(significant_words) / 2:
                    return chat_id
        
        return ""
    
    def open_teams_chat(self, item):
        try:
            values = self.tree.item(item)['values']
            
            if values:
                chat_name = values[5] if len(values) > 5 else ""
                
                # Use the search method to find the chat_id
                chat_id = self.find_chat_id(chat_name)
                
                temp_data = {
                    'row_text': "\n".join([f"{val}" for val in values]),
                    'values': values,
                    'chat_name': chat_name,
                    'chat_id': chat_id,
                    'item_id': item
                }
                
                if temp_data['chat_id']:
                    #link = f"https://teams.cloud.microsoft/l/chat/19:{temp_data['chat_id']}@thread.v2/conversations?context=%7B%22contextType%22%3A%22chat%22%7D"
                    link = f"https://teams.cloud.microsoft/l/chat/19:fc9176a569904180bbfa3eeb1bd52651@thread.v2/conversations?context=%7B%22contextType%22%3A%22chat%22%7D"
                    webbrowser.open(link)
                else:
                    
                    #print(f"Available chat names: {list(self.teams_chat_ids.keys())}")
                    messagebox.showwarning("Aviso", f"Chat ID não encontrado para: {chat_name}")
                        
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao abrir chat:\n{str(e)}")

    def send_to_printer(self,values):
        print('send_to_printer')

    def open_sharepoint_link(self, url):
        url = self.get_sharepoint_url(url)
        try:
            webbrowser.open(url)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao abrir o link:\n{str(e)}")
            self.status_label.config(text="Falha ao abrir link")
            
    def get_sharepoint_url(self, os_code):
        if not os_code:
            return ""
        os_year = os_code[-4:]
        os_parts = os_code.split('/')
        os_code_num = os_parts[0].lstrip('0')
        os_code_num = format_os_code(os_code_num) if os_code_num else ""
        return f"https://sesirs.sharepoint.com/:f:/r/sites/gdms-ISISistemasdeSensoriamento/Documentos%20Compartilhados/ISI%20SIM%20-%20Metrologia/Atendimento%20ao%20Cliente/E%20-%20Ordens%20de%20Servi%C3%A7o/{os_year}/{os_code_num}"
    

    def copy_to_clipboard(self, text):
        
        try:
            self.frame_certificados.clipboard_clear()
            self.frame_certificados.clipboard_append(self.get_sharepoint_url(text))
            self.status_label.config(text="Link copiado para a área de transferência!")
            messagebox.showinfo("Sucesso", "Link copiado para a área de transferência!")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao copiar:\n{str(e)}")

    def copy_row_data(self, item):
        
        try:
            values = self.tree.item(item)['values']

            if values:
                

                row_text = "\n".join([f"{val}" for val in values])
                
                self.frame_certificados.clipboard_clear()
                self.frame_certificados.clipboard_append(row_text)
                self.status_label.config(text="Linha copiada para a área de transferência!")
                messagebox.showinfo("Sucesso", "Dados da linha copiados para a área de transferência!")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao copiar linha:\n{str(e)}")

    def select_all_rows(self):
        
        all_items = self.tree.get_children()
        for item in all_items:
            self.tree.selection_add(item)
        self.status_label.config(text=f"{len(all_items)} linhas selecionadas")
        
        ##self.btn_gerar.config(state="enabled")
        self.btn_visualizar.config(state="enabled")
        self.btn_limpar.config(state="enabled")

    def execute_query(self, where_clause=None,groupby_clause=None):
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            rs = win32com.client.Dispatch("ADODB.Recordset")
            
            conn.Open(self.str_conn)
            
            sql = """
              
            SELECT 
                os.code 'OS',
                CONVERT(VARCHAR(10), os.receiving_date, 103) as 'Recebimento',
                CONVERT(VARCHAR(10), os.date_finished, 103) as 'Entrega',
                CASE 
                    WHEN datediff(d, os.date_finished, GETDATE()) = 0 THEN 'Hoje'
                    WHEN datediff(d, os.date_finished, GETDATE()) > 0 THEN 'Atraso'
                    WHEN os.date_finished >= DATEADD(wk, DATEDIFF(wk, 0, GETDATE()), 0) AND os.date_finished < DATEADD(wk, DATEDIFF(wk, 0, GETDATE()) + 1, 0) THEN 'Esta semana'
                    WHEN os.date_finished >= DATEADD(wk, DATEDIFF(wk, 0, GETDATE()) + 1, 0) AND os.date_finished < DATEADD(wk, DATEDIFF(wk, 0, GETDATE()) + 2, 0) THEN 'Próxima semana'
                    ELSE 'Futuro'
                END as 'Status',
                i.code as 'Item',
                s.name as 'Setor'
            FROM sectors s
            LEFT JOIN instruments i ON i.id_current_sector = s.id
            LEFT JOIN orders_services os ON os.id = i.id_service_order

            """

    #############
            if where_clause:
                sql += f" WHERE {where_clause}"
            
            if groupby_clause:
                sql += f" GROUP BY {groupby_clause}"
            
            sql += " ORDER BY os.date_finished ASC"
            
            rs.Open(sql, conn)
            
            results = []
            if not rs.EOF:
                rs.MoveFirst()
                while not rs.EOF:
                    row = []
                    for i in range(rs.Fields.Count):
                        value = rs.Fields(i).Value
                        row.append(value if value is not None else "")
                        
                    os_parts = row[0].split('/')
                    os_code = os_parts[0].lstrip('0')
                    os_code = f"{os_code}/{os_parts[1]}"
                    row[0] = os_code

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
        self.status_label.config(text="Carregando Ordens de Serviço...", font=("Segoe UI", 12))
        
        #self.status_label.config(text=f"Os dados são inválidos", font=("Segoe UI", 12))
        self.frame_certificados.update()
        
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        where_clause = " i.removed = 0 AND os.removed = 0 AND s.id IN (2,5,6,7,8,9,11,25,1026,2030,2033,15,16,2029) AND os.receiving_date > '20250101'"
        groupby_clause = "s.name, s.id, os.code, os.receiving_date, os.date_finished, i.code"
        
        results = self.execute_query(where_clause,groupby_clause)
        
        #for row in results:
        #    self.tree.insert("", "end", values=row)
        
        self.status_label.config(text=f"")
        #self.status_label.config(text=f"Os dados são inválidos")
        
        #self.tree.tag_configure('atrasado',     background="#FFD9D9")    
        #self.tree.tag_configure('hoje',         background="#B7D5F5")        
        #self.tree.tag_configure('proximo',      background="#E2FFD3")  
        
        for row in results:
            
            if row[3] == "Atraso":
                item_id = self.tree.insert("", "end", values=row)
                self.tree.item(item_id, tags=('Atraso',))
                self.tree.tag_configure('Atraso', background="#FFD9D9") 
                
            elif row[3] == "Hoje":
                item_id = self.tree.insert("", "end", values=row)
                self.tree.item(item_id, tags=('Hoje',))
                self.tree.tag_configure('Hoje', background="#98C0EB") 

            elif row[3] == "Esta semana":
                item_id = self.tree.insert("", "end", values=row)
                self.tree.item(item_id, tags=('Esta semana',))
                self.tree.tag_configure('Esta semana', background="#B7D5F5") 
            
            elif row[3] == "Próxima semana":
                item_id = self.tree.insert("", "end", values=row)
                self.tree.item(item_id, tags=('Próximo',))
                self.tree.tag_configure('Próximo', background="#D4E4F5")

            elif row[3] == "Futuro":
                item_id = self.tree.insert("", "end", values=row)
                self.tree.item(item_id, tags=('Futuro',))
                self.tree.tag_configure('Futuro', background="#E4EEF8")
            
            else:
                self.tree.insert("", "end", values=row)
    
    def perform_search(self):
        search_term = self.search_entry.get().strip()
        
        if not search_term:
            self.load_all_records()
            return
        
        #self.status_label.config(text=f"Buscando por '{search_term}'...")
        self.status_label.config(text=f"A busca de dados é inválida")
        self.frame_certificados.update()
        
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        where_clause = f"(i.code LIKE '%{search_term}%' or s.name LIKE '%{search_term}%' or os.code LIKE '%{search_term}%') AND s.removed = 0 AND i.removed = 0 AND os.removed = 0 AND s.id IN (2,5,6,7,8,9,11,25,1026,2030,2033,15,16,2029) AND os.receiving_date > '20250101'"

        groupby_clause = " s.name, s.id, os.code, os.receiving_date, os.date_finished, i.code"
        #where_clause = f""
        
        results = self.execute_query(where_clause,groupby_clause)
                
        for row in results:
            
            if row[3] == "Atraso":
                item_id = self.tree.insert("", "end", values=row)
                self.tree.item(item_id, tags=('Atraso',))
                self.tree.tag_configure('Atraso', background="#FFD9D9") 
                
            elif row[3] == "Hoje":
                item_id = self.tree.insert("", "end", values=row)
                self.tree.item(item_id, tags=('Hoje',))
                self.tree.tag_configure('Hoje', background="#98C0EB") 

            elif row[3] == "Esta semana":
                item_id = self.tree.insert("", "end", values=row)
                self.tree.item(item_id, tags=('Esta semana',))
                self.tree.tag_configure('Esta semana', background="#B7D5F5") 
            
            elif row[3] == "Próxima semana":
                item_id = self.tree.insert("", "end", values=row)
                self.tree.item(item_id, tags=('Próximo',))
                self.tree.tag_configure('Próximo', background="#D4E4F5")

            elif row[3] == "Futuro":
                item_id = self.tree.insert("", "end", values=row)
                self.tree.item(item_id, tags=('Futuro',))
                self.tree.tag_configure('Futuro', background="#E4EEF8")
            
            else:
                self.tree.insert("", "end", values=row)
    
        self.status_label.config(text=f"Encontrados {len(results)} itens para '{search_term}'")
        #self.status_label.config(text=f"Os dados são inválidos")
    
    
    def on_select(self, event):
        selected = self.tree.selection()
        if len(selected) > 1:
            self.status_label.config(text=f"{len(selected)} itens selecionados")
            self.btn_visualizar.config(state="enabled")
            self.btn_limpar.config(state="enabled")
        elif len(selected) == 1:
            self.status_label.config(text=f"{len(selected)} item selecionado")
            self.btn_visualizar.config(state="enabled")
            self.btn_limpar.config(state="enabled")
        else:
            self.status_label.config(text=f"Nenhum item selecionado")
            self.btn_visualizar.config(state="disabled")
            self.btn_limpar.config(state="disabled")
    
    def get_selected_data(self):
        selected_items = self.tree.selection()
        selected_data = []
        
        for item in selected_items:
            values = self.tree.item(item)['values']
            selected_data.append({
                'sharepoint': f"{values[0]}/Cliente",
                'ordem_servico': values[1].lstrip('0'),
                'certificado': values[2],
                'item': values[3],
                'data_calibracao': values[4],
                'cliente': values[5]
            })
        
        return selected_data
    
    def print_selected(self):
        selected_data = self.get_selected_data()
        
        if not selected_data:
            messagebox.showinfo("Aviso", "Nenhum item selecionado para visualizar.")
            return
        
        print_window = tk.Toplevel(self.frame_certificados)
        print_window.title("Registros Selecionados")
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
            output_text += f"Certificado n° {item['certificado']}:\n\n"
            
            output_text += f"{item['sharepoint']}\n"
            output_text += f"{item['ordem_servico']}\n"
            output_text += f"{item['certificado']}\n"
            output_text += f"{item['item']}\n"
            output_text += f"{item['data_calibracao']}\n"
            output_text += f"{item['cliente']}\n"
            
            output_text += f"{'-'*40}\n\n"

        text_widget.insert("1.0", output_text)
        text_widget.config(state="disabled")
        
    def clear_selection(self):
        for item in self.tree.selection():
            self.tree.selection_remove(item)
        self.status_label.config(text="Seleção limpa")

status_servidor, cor_status =   verificar_disponibilidade()
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
#db_viewer0 = ServiceScheduler(notebook, STR_CONN, cor_fundo)   ### ZEBRA
db_viewer1 = ServiceSchedulerLinkedDirect(notebook, STR_CONN_LINKED,STR_CONN, cor_fundo)   ### ZEBRA
db_viewer6 = DatabaseViewer6(notebook, STR_CONN_LINKED, cor_fundo)         ### VISTA GERAL DE ORDENS DE SERVIÇO

#db_viewer = DatabaseViewer(notebook, STR_CONN, cor_fundo)
#db_viewer4 = DatabaseViewer4(notebook, STR_CONN, cor_fundo)         ### AGENDA DE ORDENS DE SERVIÇO
#db_viewer6 = DatabaseViewer6(notebook, STR_CONN, cor_fundo)         ### VISTA GERAL DE ORDENS DE SERVIÇO
##db_viewer5 = DatabaseViewer5(notebook, STR_CONN, cor_fundo)         ###
#db_viewer2 = DatabaseViewer2(notebook, STR_CONN, cor_fundo)         ### PADRÕES

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

ttk.Label(content_frame, text=F"Status do Servidor {db} {str_alternativa}").pack(pady=(20, 0))
label_status = tk.Label(content_frame, text=status_servidor, fg=cor_status, bg=cor_fundo, font=("Segoe UI", 10, "bold"))
label_status.pack(pady=5)

ttk.Label(content_frame, text=F"Status do Servidor {db_linked}").pack(pady=(0, 0))
label_status3 = tk.Label(content_frame, text=status_servidor, fg=cor_status, bg=cor_fundo, font=("Segoe UI", 10, "bold"))
label_status3.pack(pady=5)

info_texto = f"Driver: SQLOLEDB\nVersão: {INFO_VERSAO}"
label_info = ttk.Label(content_frame, text=info_texto, justify="center", font=("Consolas", 9))
label_info.pack(pady=15)

#ttk.Label(content_frame, text="Histórico de Atualizações", justify="left", font=("Consolas", 9, "bold")).pack(padx=10, pady=(15, 5))

#text_frame = ttk.Frame(content_frame)
#text_frame.pack(padx=10, pady=5, fill="both", expand=True)

#text_widget = tk.Text(text_frame, wrap="word", font=("Consolas", 10), height=21)
#text_scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=text_widget.yview)
#text_widget.configure(yscrollcommand=text_scrollbar.set)

#text_scrollbar.pack(side="right", fill="y")
#text_widget.pack(side="left", expand=True, fill="both")

updates_text = f"""
Versão {INFO_VERSAO}

> (Problema) Link do Sharepoint: Antes, caso os algarismos significativos do código da OS antes do caractere "/" possuíssem 4 algarismos (i.e. "1510","1250"), o link gerado para o Sharepoint estava incorreto (i.e. "01510","01250")
>>> (Ação corretiva) O algoritmo agora preenche este número com zeros à esquerda até completar 4 algarismos. Caso já possua esta quantidade, nenhum novo caractere é adicionado ao código da OS
> (Remoção) A coluna "Sharepoint" não está mais visível na aba "Certificados"
> (Adição) Link para suporte: Agora, o nome do desenvolvedor presente no rodapé desta página leva a um novo chat do Microsoft Teams com ele
> (Adição) Aba "Padrões", aba "Zebra [Em desenvolvimento]", aba "Agenda de Ordens de Serviço", aba "" e Histórico de atualizações
> (Remoção) Remoção de separador de lista adicional ao final do menu de contexto na aba "Certificados"
> (Problema) Tamanho da janela do Assistente de Dados: O limite inferior de dimensões da janela não estava definido, a janela poderia ser dimensionada para 0 x 0 pixels
>>> (Ação corretiva) As dimensões da janela agora possuem o limite inferior de 1150 x 675 pixels
"""
updates_text = ""

#text_widget.insert("1.0", updates_text)
#text_widget.config(state="disabled")

footer_frame = ttk.Frame(content_frame)
footer_frame.pack(side="bottom", pady=10, fill="x")

credits = ttk.Label(footer_frame, text="Castro", font=("Segoe UI", 8, "italic"), cursor="hand2")
credits.pack()
ttk.Label(footer_frame, text="2026 IST PGE - Laboratório de Metrologia", font=("Segoe UI", 8, "italic")).pack()

def open_link(event):
    webbrowser.open("https://teams.microsoft.com/l/chat/0/0?users=guilherme.castro@senairs.org.br&message=E%20aí%20meu%20chapa")

credits.bind("<Button-1>", open_link)

#db_viewer3 = DatabaseViewer3(notebook, STR_CONN_ZEBRA, cor_fundo)   ### ZEBRA

root.mainloop()
