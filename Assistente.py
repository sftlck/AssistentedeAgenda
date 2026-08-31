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


ip =            '10.165.212.74'
db =            'Castro_Services'
user =          'sa'
password =      'Wheelp0p2'
INFO_VERSAO =   "05/05/2026"

#ip_zebra =            '10.165.212.74'
#db_zebra =            'Zebra'
#user_zebra =          'standard_zebra'
#password_zebra =      'Wheelp0p1'


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

STR_CONN = (
    f"Provider=SQLOLEDB;"
    f"Data Source={ip};"
    f"Initial Catalog={db};"
    f"User ID={user};"
    f"Password={password};"
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
        self.frame_certificados = ttk.Frame(parent_notebook)
        parent_notebook.add(self.frame_certificados, text=" Agendamento de Serviços ")
        
        parent_notebook.bind('<<NotebookTabChanged>>', self.on_tab_changed)
    
        
        self.current_month_offset = 0
        self.calendar_data = {}
        self.selected_date = None
        
        self.setup_ui()
        self.load_services()
        self.load_calendar_data()
        self.render_calendar()
        self.refresh_calendar()

    def on_tab_changed(self, event):
        """Refresh calendar when this tab is selected"""
        notebook = event.widget
        current_tab = notebook.select()
        current_tab_text = notebook.tab(current_tab, "text")
        self.load_calendar_data()
        self.refresh_calendar()
        if "Agendamento de Serviços" in current_tab_text:
            self.refresh_calendar()

    def setup_ui(self):
        # Main container - split into left panel (new service) and right panel (calendar)
        main_container = tk.Frame(self.frame_certificados, bg=self.cor_fundo)
        main_container.pack(fill='both', expand=True, padx=10, pady=10)
        
        # ===== LEFT PANEL =====
        left_panel = tk.Frame(main_container, bg=self.cor_fundo, width=350)
        left_panel.pack(side='left', fill='y', padx=(0, 10))
        left_panel.pack_propagate(False)
        
        # Greeting
        self.greetings = ttk.Label(left_panel, text=f"{buenas}", font=("Segoe UI", 10),
                                   background='black')
        self.greetings.pack(pady=(0, 10))
        
        # Title
        #tk.Label(left_panel, text="Novo Serviço", font=('Segoe UI', 12, 'bold'), bg=self.cor_fundo, fg='white').pack(pady=(0, 10))
        
        # Service Selection
        #tk.Label(left_panel, text="Serviço:", font=('Segoe UI', 9), 
        #        bg=self.cor_fundo, fg='white').pack(anchor='w')
        #
        #self.service_var = tk.StringVar()
        #self.service_combo = ttk.Combobox(left_panel, textvariable=self.service_var, state='readonly', width=40)
        #self.service_combo.pack(fill='x', pady=(2, 10))
        #self.service_combo.bind('<<ComboboxSelected>>', self.on_service_selected)
        #
        ## Duration display
        #self.lbl_duration = tk.Label(left_panel, text="Duração: -- min", font=('Segoe UI', 9), bg=self.cor_fundo, fg='#B7D5F5')
        #self.lbl_duration.pack(anchor='w', pady=(0, 10))
        #
        ## Customer Notes
        #tk.Label(left_panel, text="Observações:", font=('Segoe UI', 9), bg=self.cor_fundo, fg='white').pack(anchor='w')
        #
        #self.notes_entry = tk.Text(left_panel, height=3, width=40, font=('Segoe UI', 9))
        #self.notes_entry.pack(fill='x', pady=(2, 5))
        
        # Priority
        #tk.Label(left_panel, text="Prioridade:", font=('Segoe UI', 9), bg=self.cor_fundo, fg='white').pack(anchor='w')
        #
        #priority_frame = tk.Frame(left_panel, bg=self.cor_fundo)
        #priority_frame.pack(fill='x', pady=(2, 10))
        #
        #self.priority_var = tk.StringVar(value="Normal")
#
        #style = ttk.Style()
        #style.configure("TRadiobutton", background='#1B4B9F', foreground='white', font=('Arial', 10))
        #ttk.Radiobutton(priority_frame, text="Normal", variable=self.priority_var, value="Normal",style='TRadiobutton').pack(side='left', padx=5)
        #ttk.Radiobutton(priority_frame, text="Urgente", variable=self.priority_var, value="Urgente",style='TRadiobutton').pack(side='left', padx=5)
#
        ##self.buscar = ttk.Button(search_frame, text="Buscar", command=self.perform_search).pack(side="left", padx=5)
        #self.buscar = ttk.Button(left_panel, text="Agendar Serviço", command=self.add_service, cursor='hand2').pack( pady=5)
#
        #self.buscar = ttk.Button(left_panel, text=" Adiantar atualização de serviços com Metroex ", command=self.sync_services, cursor='hand2').pack(pady=5)
        
        #self.btn_add = tk.Button(left_panel, text="➕ Agendar Serviço", font=('Segoe UI', 11, 'bold'), bg='#28A745', fg='white', padx=20, pady=8,cursor='hand2', command=self.add_service)
        #self.btn_add.pack(pady=10)

        #self.btn_schedule = tk.Button(left_panel, text="📅 Processar Fila (FIFO)", font=('Segoe UI', 11, 'bold'), bg='#1B4B9F', fg='white', padx=20, pady=8,cursor='hand2', command=self.schedule_all)
        
        # Queue count
        self.lbl_queue = tk.Label(left_panel, text="Fila: 0 pendentes", font=('Segoe UI', 9), bg=self.cor_fundo, fg='#FFD700')
        self.lbl_queue.pack(pady=(10, 0))
        
        # Pending list
        pending_frame = tk.Frame(left_panel, bg=self.cor_fundo)
        pending_frame.pack(fill='both', expand=True)
        
        tk.Label(pending_frame, text="Serviços Pendentes:", font=('Segoe UI', 9, 'bold'), bg=self.cor_fundo, fg='white').pack
        
        self.pending_listbox = tk.Listbox(pending_frame, font=('Segoe UI', 8), height=6, bg='white', fg='black')
        self.pending_listbox.pack(fill='both', expand=True, pady=(2, 0))
        
        # ===== RIGHT PANEL (Calendar) =====
        self.calendar_frame = tk.Frame(main_container, bg=self.cor_fundo)
        self.calendar_frame.pack(side='left', fill='both', expand=True)
        
        # Status bar
        self.status_label = ttk.Label(self.frame_certificados, text="Pronto para agendamento", font=("Segoe UI", 12))
        self.status_label.pack(pady=5)
    
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
                    s.service_code,
                    s.specification,
                    s.time_execution,
                    ss.notes,
                    ss.status,
                    ss.priority,
                    ss.id as schedule_id
                FROM [castro_services].dbo.Time_Slots ts
                JOIN [castro_services].dbo.Service_Schedule ss ON ts.schedule_id = ss.id
                JOIN [castro_services].dbo.Services s ON ts.service_id = s.id
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
                        'service_code': rs.Fields('service_code').Value,
                        'specification': rs.Fields('specification').Value,
                        'time_execution': rs.Fields('time_execution').Value,
                        'notes': rs.Fields('notes').Value,
                        'status': rs.Fields('status').Value,
                        'priority': rs.Fields('priority').Value
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
        """Render the monthly calendar view - based on DatabaseViewer4.render_calendar"""
        
        # Clear previous calendar
        for widget in self.calendar_frame.winfo_children():
            widget.destroy()
        
        # Get target month/year
        today = datetime.now()
        target_date = today + relativedelta(months=self.current_month_offset)
        target_date = target_date.replace(day=1)
        year, month = target_date.year, target_date.month
        
        # Month names in Portuguese
        meses_pt = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho','Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
        
        # ===== Navigation Header =====
        nav_frame = tk.Frame(self.calendar_frame, bg=self.cor_fundo)
        nav_frame.pack(fill='x', pady=(0, 5))
        
        btn_prev = tk.Label(nav_frame, text="◀", anchor='center', justify='center',font=('Segoe UI', 12, 'bold'), bg=self.cor_fundo, fg='white',cursor='hand2', padx=0)
        btn_prev.pack(side='left', padx=5)
        btn_prev.bind('<Button-1>', lambda e: self.change_month(-1))
        
        month_label = tk.Label(nav_frame, text=f"{meses_pt[month-1]} {year}",font=('Segoe UI', 12, 'bold'), bg=self.cor_fundo, fg='white')
        month_label.pack(side='left', padx=5)
        
        btn_next = tk.Label(nav_frame, text="▶", anchor='center', justify='center',font=('Segoe UI', 12, 'bold'), bg=self.cor_fundo, fg='white',cursor='hand2', padx=0)
        btn_next.pack(side='left', padx=5)
        btn_next.bind('<Button-1>', lambda e: self.change_month(1))
        
        #btn_today = tk.Label(nav_frame, text="Hoje", font=('Segoe UI', 9, 'bold'),bg="#FFFFFF", fg='black', padx=10, pady=2,cursor='hand2', borderwidth='1')
        
        btn_today = ttk.Button(nav_frame, text=" Hoje ", command=self.go_to_today, cursor='hand2')
        btn_today.pack(side="right", padx=5)

        #btn_today = tk.Label(nav_frame, text="Hoje", font=('Segoe UI', 9, 'bold'), padx=10, pady=2,cursor='hand2')
        #btn_today.pack(side='right', padx=5)
        #btn_today.bind('<Button-1>', lambda e: self.go_to_today())
        
        #btn_refresh = tk.Label(nav_frame, text="🔄", font=('Segoe UI', 9, 'bold'),bg="#FFFFFF", fg='black', padx=10, pady=2,cursor='hand2', borderwidth='1')
        #btn_refresh.pack(side='right', padx=5)

        self.buscar = ttk.Button(nav_frame, text=" Limpar agenda ", command=self.clear_all_scheduled, cursor='hand2')
        self.buscar.pack(side="right", padx=5)
        btn_refresh = ttk.Button(nav_frame, text=" Atualizar ", command=self.refresh_calendar, cursor='hand2')
        btn_refresh.pack(side="right", padx=5)

        
        self.btn_schedule = ttk.Button(nav_frame, text="Agendar", command=self.schedule_all, cursor='hand2')

        self.btn_schedule.pack(side="right", padx=5)


        #btn_refresh.bind('<Button-1>', lambda e: self.refresh_calendar())
        
        # ===== Weekday Headers =====
        header_frame = tk.Frame(self.calendar_frame, bg=self.cor_fundo)
        header_frame.pack(fill='x')
        
        dias_semana = ['Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb', 'Dom']
        for dia in dias_semana:
            lbl = tk.Label(header_frame, text=dia, font=('Segoe UI', 9, 'bold'),
                          bg='#1B4B9F', fg='white', width=16, pady=4)
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
                    # Empty cell
                    day_frame = tk.Frame(week_frame, bg='#908F8F', width=120, height=80)
                    day_frame.pack(side='left', padx=1, pady=1)
                    day_frame.pack_propagate(False)
                else:
                    current_date = datetime(year, month, day).date()
                    calendar_key = (day, month, year)
                    
                    is_today = (current_date == today_date)
                    is_weekend = current_date.weekday() >= 5
                    has_orders = calendar_key in self.calendar_data
                    
                    # Background color logic
                    if is_today:
                        bg_color = '#B7D5F5'
                    elif is_weekend:
                        bg_color = "#908F8F"
                    elif has_orders:
                        bg_color = '#FFFFFF'
                    else:
                        bg_color = '#908F8F'
                    
                    if has_orders and not is_weekend:
                        day_frame = tk.Frame(week_frame, bg=bg_color, width=120, height=80,
                                            bd=1, cursor='hand2')
                        day_frame.pack(side='left', padx=1, pady=1)
                        day_frame.pack_propagate(False)
                    else:
                        day_frame = tk.Frame(week_frame, bg=bg_color, width=120, height=80, bd=1)
                        day_frame.pack(side='left', padx=1, pady=1)
                        day_frame.pack_propagate(False)
                    
                    # Day number
                    day_num_frame = tk.Frame(day_frame, bg=bg_color)
                    day_num_frame.pack(fill='x', padx=2, pady=1)
                    
                    fg_day = 'white' if is_today else 'black'
                    day_label = tk.Label(day_num_frame, text=str(day),
                                        font=('Segoe UI', 9, 'bold'),
                                        bg=bg_color, fg=fg_day, anchor='w')
                    day_label.pack(side='left')
                    
                    # Order count badge
                    if has_orders and not is_weekend:
                        count = len(self.calendar_data[calendar_key])
                        badge = tk.Label(day_num_frame, text=str(count),
                                        font=('Segoe UI', 8, 'bold'),
                                        bg='#2462D2', fg='white', width=3, height=1)
                        badge.pack(side='right')
                    
                    # Order preview (up to 3 unique service codes)
                    if has_orders and not is_weekend:
                        unique_codes = []
                        seen = set()
                        for order in self.calendar_data[calendar_key]:
                            service_code = order['service_code']
                            if service_code not in seen:
                                unique_codes.append(service_code)
                                seen.add(service_code)
                        
                        preview_text = ""
                        for service_code in unique_codes[:3]:
                            preview_text += f"{service_code}\n"
                        
                        preview = tk.Label(day_frame, text=preview_text.strip(),
                                          font=('Segoe UI', 7), bg=bg_color,
                                          fg='#333333', anchor='w', justify='left',
                                          cursor='hand2')
                        preview.pack(fill='x', padx=3)
                        
                        if len(unique_codes) > 3:
                            more_label = tk.Label(day_frame,
                                                 text=f"+{len(unique_codes) - 3} mais...",
                                                 font=('Segoe UI', 6), bg=bg_color,
                                                 fg='#908F8F', anchor='w')
                            more_label.pack(fill='x', padx=3)
                    
                    # Bind click event
                    if has_orders and not is_weekend:
                        day_frame.bind('<Button-1>', lambda e, key=calendar_key: self.show_day_orders(key))
                        for child in day_frame.winfo_children():
                            child.bind('<Button-1>', lambda e, key=calendar_key: self.show_day_orders(key))
            
    
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
        self.load_pending_list()
        #self.status_label.config(text="Calendário atualizado")
    
    def show_day_orders(self, calendar_key):
        """Show orders for a specific day in a popup window"""
        if calendar_key not in self.calendar_data:
            return
        
        orders = self.calendar_data[calendar_key]
        date_str = f"{calendar_key[0]:02d}/{calendar_key[1]:02d}/{calendar_key[2]}"
        
        # Create popup window
        popup = tk.Toplevel(self.frame_certificados)
        popup.geometry("700x400")
        popup.minsize(550, 300)
        popup.configure(bg=self.cor_fundo)
        popup.title(f"Serviços Agendados - {date_str}")
        
        # Treeview
        tree_frame = ttk.Frame(popup)
        tree_frame.pack(pady=10, padx=10, expand=True, fill='both')
        
        tree_scroll = ttk.Scrollbar(tree_frame)
        tree_scroll.pack(side='right', fill='y')
        
        columns = ("Horário", "Código", "Serviço", "Duração", "Status", "Observações")
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings", yscrollcommand=tree_scroll.set, selectmode="extended")
        tree_scroll.config(command=tree.yview)
        
        tree.heading("Horário", text="Horário")
        tree.heading("Código", text="Código")
        tree.heading("Serviço", text="Serviço")
        tree.heading("Duração", text="Duração")
        tree.heading("Status", text="Status")
        tree.heading("Observações", text="Observações")
        
        tree.column("Horário", width=120)
        tree.column("Código", width=80)
        tree.column("Serviço", width=200)
        tree.column("Duração", width=60)
        tree.column("Status", width=80)
        tree.column("Observações", width=200)
        
        tree.pack(expand=True, fill='both')
        
        # Color tags
        tree.tag_configure('urgente', background='#FFD9D9')
        tree.tag_configure('normal', background='#FFFFFF')
        
        for order in orders:
            start_str = str(order['start_time'])[:5] if order['start_time'] else "--:--"
            end_str = str(order['end_time'])[:5] if order['end_time'] else "--:--"
            
            is_urgent = order.get('priority', 10) <= 5
            tag = 'urgente' if is_urgent else 'normal'
            
            tree.insert("", "end", values=(
                f"{start_str} - {end_str}",
                order['service_code'],
                order['specification'],
                f"{order['time_execution']} min",
                order['status'],
                order['notes'][:50]
            ), tags=(tag,))
        
        # Close button
        #btn_frame = tk.Frame(popup, bg=self.cor_fundo)
        #btn_frame.pack(pady=10)
        #tk.Button(btn_frame, text="Fechar", command=popup.destroy,
        #         font=('Segoe UI', 10), bg='#1B4B9F', fg='white',
        #         padx=20, pady=5).pack()
    
    def load_services(self):
        """Load available services from database"""
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            rs = win32com.client.Dispatch("ADODB.Recordset")
            
            conn.Open(self.str_conn)
            
            sql = """
                SELECT id, service_code, specification, time_execution, description
                FROM [castro_services].dbo.Services
                WHERE is_active = 1
                ORDER BY service_code
            """
            
            rs.Open(sql, conn)
            
            self.services_data = []
            service_list = []
            
            if not rs.EOF:
                rs.MoveFirst()
                while not rs.EOF:
                    service_id = rs.Fields('id').Value
                    service_code = rs.Fields('service_code').Value
                    spec = rs.Fields('specification').Value
                    duration = rs.Fields('time_execution').Value
                    
                    self.services_data.append({
                        'id': service_id,
                        'service_code': service_code,
                        'specification': spec,
                        'time_execution': duration
                    })
                    
                    service_list.append(f"{service_code} - {spec} ({duration} min)")
                    rs.MoveNext()
            
            rs.Close()
            conn.Close()
            
           # self.service_combo['values'] = service_list
            
        except Exception as e:
            self.status_label.config(text=f"Erro ao carregar serviços: {str(e)}")
    
    def on_service_selected(self, event=None):
        idx = self.service_combo.current()
        if idx >= 0 and idx < len(self.services_data):
            duration = self.services_data[idx]['time_execution']
            self.lbl_duration.config(text=f"Duração: {duration} min")
    
    def add_service(self):
        """Add a new service request to the queue"""
        idx = self.service_combo.current()
        
        if idx < 0:
            messagebox.showwarning("Aviso", "Selecione um serviço primeiro!")
            return
        
        service = self.services_data[idx]
        notes = self.notes_entry.get("1.0", "end-1c").strip()
        priority = 5 if self.priority_var.get() == "Urgente" else 10
        
        if not notes:
            notes = f"Cliente: Agendamento - {service['specification']}"
        
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            conn.Open(self.str_conn)
            
            sql = f"""

                INSERT INTO [castro_services].dbo.Service_Schedule (service_id, notes, priority)

                VALUES ({service['id']}, '{notes.replace("'", "''")}', {priority})

            """
            
            conn.Execute(sql)
            conn.Close()
            
            # Clear form
            self.notes_entry.delete("1.0", "end")
            self.service_var.set("")
            self.lbl_duration.config(text="Duração: -- min")
            
            self.status_label.config(text=f"✅ Serviço '{service['specification']}' adicionado à fila!")
            self.update_queue_count()
            self.load_pending_list()
            
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao adicionar serviço:\n{str(e)}")
    
    def sync_services(self):
        """Sync services from linked server"""
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            conn.Open(self.str_conn)
            conn.Execute("EXEC [castro_services].dbo.sp_Sync_Service_Modes")
            conn.Close()
            
            self.load_services()
            self.status_label.config(text="✅ Serviços sincronizados com o servidor vinculado!")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao sincronizar:\n{str(e)}")

    def clear_all_scheduled(self):
        """Clear all scheduled services completely - remove from agenda and queue"""
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            conn.Open(self.str_conn)
            
            # Delete all time slots (the agenda)
            conn.Execute("DELETE FROM [castro_services].dbo.Time_Slots")
            
            # Delete all service schedule records completely
            conn.Execute("DELETE FROM [castro_services].dbo.Service_Schedule")
            
            # Reset identity counter (optional - starts IDs from 1 again)
            conn.Execute("DBCC CHECKIDENT ('[castro_services].dbo.Service_Schedule', RESEED, 0)")
            
            conn.Close()
            
            # Refresh all UI elements
            self.load_calendar_data()
            self.render_calendar()
            self.update_queue_count()
            if hasattr(self, 'load_pending_list'):
                self.load_pending_list()
            
            self.status_label.config(text="Os registros da agendas foram removidos")
            messagebox.showinfo("Sucesso", "Todos os agendamentos foram removidos!")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao limpar agenda:\n{str(e)}")

    def schedule_all(self):
        """Process all pending services using FIFO"""
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            conn.Open(self.str_conn)
            
            # Count pending
            rs = win32com.client.Dispatch("ADODB.Recordset")
            rs.Open("SELECT COUNT(*) as cnt FROM [castro_services].dbo.Service_Schedule WHERE status = 'PENDING'", conn)
            
            pending_count = rs.Fields('cnt').Value if not rs.EOF else 0
            rs.Close()
            
            if pending_count == 0:
                messagebox.showinfo("Aviso", "Nenhum serviço pendente na fila!")
                conn.Close()
                return
            
            # Execute scheduling
            conn.Execute("EXEC [castro_services].dbo.sp_ScheduleNextServices @batch_size = 1000")
            
            conn.Close()

            if pending_count == 1:
                self.status_label.config(text=f"{pending_count} serviço agendado")
                
            if pending_count > 1:
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
            
            self.lbl_queue.config(text=f"Fila: {count} pendente(s)")
            
        except Exception:
            pass
    
    def load_pending_list(self):
        """Load pending services into the listbox"""
        self.pending_listbox.delete(0, 'end')
        
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            conn.Open(self.str_conn)
            
            sql = """
                SELECT 
                    s.service_code,
                    s.specification,
                    ss.notes,
                    ss.priority,
                    ss.requested_at
                FROM [castro_services].dbo.Service_Schedule ss
                JOIN [castro_services].dbo.Services s ON ss.service_id = s.id
                WHERE ss.status = 'PENDING'
                ORDER BY ss.priority ASC, ss.requested_at ASC
            """
            
            rs = win32com.client.Dispatch("ADODB.Recordset")
            rs.Open(sql, conn)
            
            if not rs.EOF:
                rs.MoveFirst()
                while not rs.EOF:
                    service_code = rs.Fields('service_code').Value
                    spec = rs.Fields('specification').Value
                    priority = rs.Fields('priority').Value
                    urgency = "⚡" if priority <= 5 else "  "
                    
                    self.pending_listbox.insert('end', f"{urgency} {service_code} - {spec[:30]}")
                    rs.MoveNext()
            
            rs.Close()
            conn.Close()
            
        except Exception:
            pass


class ServiceSchedulerLinkedDirect:
    """ ABA: Serviços para Agendamento (Linked Server - 631CP) """
    
    def __init__(self, parent_notebook, str_conn, str_conn_primary,cor_fundo):
        self.str_conn = str_conn
        self.str_conn_primary = str_conn_primary
        self.cor_fundo = cor_fundo
        self.frame_certificados = ttk.Frame(parent_notebook)
        parent_notebook.add(self.frame_certificados, text=" Serviços para Agendamento ")
        
        parent_notebook.bind('<<NotebookTabChanged>>', self.on_tab_changed)
    
        self.selected_services = []
        
        self.setup_ui()
        self.load_services()

    def on_tab_changed(self, event):
        print('tab changed')
        notebook = event.widget
        current_tab = notebook.select()
        current_tab_text = notebook.tab(current_tab, "text")

        if "Serviços Vinculados" in current_tab_text or "Agendar Serviços" in current_tab_text:
            self.load_services()

    def setup_ui(self):
        main_container = tk.Frame(self.frame_certificados, bg=self.cor_fundo)
        main_container.pack(fill='both', expand=True, padx=10, pady=10)
        
        # ===== TOP FRAME - Info & Controls =====
        top_frame = tk.Frame(main_container, bg=self.cor_fundo)
        top_frame.pack(fill='x', pady=(0, 10))
        
        # Greeting
        self.greetings = ttk.Label(top_frame, text=f"{buenas}", font=("Segoe UI", 10),
                                   background='black')
        self.greetings.pack(side='left', padx=5)
        
        # Title
        tk.Label(top_frame, text="Serviços [Qualidade]",font=('Segoe UI', 12, 'bold'), bg=self.cor_fundo, fg='white').pack(side='left', padx=20)
        
        # Refresh button
        #btn_refresh = tk.Button(top_frame, text="🔄 Atualizar Lista",font=('Segoe UI', 9), bg='#17A2B8', fg='white',padx=10, pady=2, cursor='hand2',command=self.load_services)
        #btn_refresh.pack(side='right', padx=5)
        
        btn_refresh = ttk.Button(top_frame, text=" Atualizar Lista ", command=self.load_services, cursor='hand2')
        btn_refresh.pack(side="right", padx=5)
        #btn_refresh.config(state="disabled")
        
        # Selection count
        self.lbl_count = tk.Label(top_frame, text="0 selecionados",
                                   font=('Segoe UI', 9, 'bold'),
                                   bg=self.cor_fundo, fg='#FFD700')
        self.lbl_count.pack(side='right', padx=20)
        
        # ===== SEARCH FRAME =====
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
        
        columns = ("OS", "Item", "Especificação", "Descrição", "Code", "Duração")
        
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
            "Code": 100,
            "Duração": 70
        }
        
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=column_widths.get(col, 80), minwidth=50)
        
        self.tree.pack(expand=True, fill='both')
        
        # Bind events
        self.tree.bind('<<TreeviewSelect>>', self.on_tree_select)
        self.tree.bind('<Button-3>', self.show_context_menu)
        self.tree.bind('<Control-a>', self.select_all)
        self.tree.bind('<Control-A>', self.select_all)
        
        # ===== BUTTON FRAME =====
        button_frame = tk.Frame(main_container, bg=self.cor_fundo)
        button_frame.pack(fill='x', pady=(10, 0))
        
        #self.btn_schedule = tk.Button(button_frame, text="📅 Agendar Selecionados",font=('Segoe UI', 10, 'bold'),bg='#28A745', fg='white', padx=15, pady=6,cursor='hand2', command=self.schedule_selected)

        self.btn_schedule = ttk.Button(button_frame, text=" Agendar Selecionados ", command=self.schedule_selected, cursor='hand2')
        self.btn_schedule.pack(side="left", padx=5)
        self.btn_schedule.config(state="disabled")
        
        #self.btn_select_all = tk.Button(button_frame, text="Selecionar Todos",font=('Segoe UI', 9),bg='#6C757D', fg='white', padx=10, pady=4,cursor='hand2', command=self.select_all)
        #self.btn_select_all.pack(side='left', padx=5)
        
        self.btn_select_all = ttk.Button(button_frame, text=" Selecionar Todos ", command=self.select_all, cursor='hand2')
        self.btn_select_all.pack(side="left", padx=5)
        #self.btn_select_all.config(state="disabled")
        
        #self.btn_clear = tk.Button(button_frame, text="Limpar Seleção",font=('Segoe UI', 9),bg='#6C757D', fg='white', padx=10, pady=4,cursor='hand2', command=self.clear_selection)
        #self.btn_clear.pack(side='left', padx=5)
        #self.btn_clear.config(state='disabled')

        
        self.btn_clear = ttk.Button(button_frame, text=" Limpar Seleção ", command=self.clear_selection, cursor='hand2')
        self.btn_clear.pack(side="left", padx=5)
        #self.btn_clear.config(state="disabled")
        
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

            """
            
            rs.Open(sql, conn)
            
            if not rs.EOF:
                rs.MoveFirst()
                while not rs.EOF:
                    execution_time_raw = rs.Fields('execution_time').Value
                    duration_minutes = self.convert_time_to_minutes(execution_time_raw) if execution_time_raw else 0
                    
                    row_data = {
                        'OS': rs.Fields('OS').Value if rs.Fields('OS').Value else "",
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
                        f"{row_data['execution_time_raw']} ({duration_minutes} min)"
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
                    f"{row['execution_time_raw']} ({row['duration_minutes']} min)"
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
                    IF NOT EXISTS (SELECT 1 FROM [castro_services].dbo.Services WHERE service_code = '{row['Code']}')
                    BEGIN
                        INSERT INTO [castro_services].dbo.Services 
                            (service_code, specification, description, time_execution, is_active)
                        VALUES (
                            '{row['Code']}',
                            '{str(row['Especificação']).replace("'", "''")}',
                            '{str(row['Descrição']).replace("'", "''")}',
                            {duration},
                            1
                        );
                    END
                    ELSE
                    BEGIN
                        UPDATE [castro_services].dbo.Services
                        SET time_execution = {duration},
                            specification = '{str(row['Especificação']).replace("'", "''")}',
                            updated_at = GETDATE()
                        WHERE service_code = '{row['Code']}';
                    END
                """
                conn.Execute(sync_sql)
                
                # Get the local service ID
                rs = win32com.client.Dispatch("ADODB.Recordset")
                rs.Open(f"SELECT id FROM [castro_services].dbo.Services WHERE service_code = '{row['Code']}'", conn)
                
                local_service_id = rs.Fields('id').Value if not rs.EOF else None
                rs.Close()
                
                if local_service_id:
                    # Add to schedule queue
                    notes = f"OS {row['OS']} - {row['Item']} - {row['Especificação']}"
                    
                    schedule_sql = f"""
                        INSERT INTO [castro_services].dbo.Service_Schedule 
                            (service_id, notes, priority)
                        VALUES ({local_service_id}, '{notes.replace("'", "''")}', 10)
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
            
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao agendar:\n{str(e)}")
    
    #def process_fifo(self):
    #    """Process the FIFO queue"""
    #    try:
    #        conn = win32com.client.Dispatch("ADODB.Connection")
    #        conn.Open(self.str_conn)
    #        
    #        rs = win32com.client.Dispatch("ADODB.Recordset")
    #        rs.Open(
    #            "SELECT COUNT(*) as cnt FROM [castro_services].dbo.Service_Schedule WHERE status = 'PENDING'",
    #            conn
    #        )
    #        
    #        pending_count = rs.Fields('cnt').Value if not rs.EOF else 0
    #        rs.Close()
    #        
    #        if pending_count == 0:
    #            messagebox.showinfo("Aviso", "Nenhum serviço pendente na fila!")
    #            conn.Close()
    #            return
    #        
    #        conn.Execute("EXEC [castro_services].dbo.sp_ScheduleNextServices @batch_size = 1000")
    #        conn.Close()
    #        
    #        self.status_label.config(text=f"✅ {pending_count} serviço(s) processado(s) com sucesso!")
    #        
    #    except Exception as e:
    #        messagebox.showerror("Erro", f"Falha ao processar fila:\n{str(e)}")
#
#

status_servidor, cor_status =   verificar_disponibilidade()
status_servidor3, cor_status3 = verificar_disponibilidade3()

root = tk.Tk()
root.title("Assistente de Agendamento - Secretaria Técnica")

#root.geometry("900x500+2020+100")
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
db_viewer0 = ServiceScheduler(notebook, STR_CONN, cor_fundo)   ### ZEBRA
db_viewer1 = ServiceSchedulerLinkedDirect(notebook, STR_CONN_LINKED,STR_CONN, cor_fundo)   ### ZEBRA

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

ttk.Label(content_frame, text=F"Status do Servidor {db}").pack(pady=(20, 0))
label_status = tk.Label(content_frame, text=status_servidor, fg=cor_status, bg=cor_fundo, font=("Segoe UI", 10, "bold"))
label_status.pack(pady=5)

ttk.Label(content_frame, text=F"Status do Servidor {db_linked}").pack(pady=(0, 0))
label_status3 = tk.Label(content_frame, text=status_servidor, fg=cor_status, bg=cor_fundo, font=("Segoe UI", 10, "bold"))
label_status3.pack(pady=5)

info_texto = f"Driver: SQLOLEDB\nVersão: {INFO_VERSAO}"
label_info = ttk.Label(content_frame, text=info_texto, justify="center", font=("Consolas", 9))
label_info.pack(pady=15)

ttk.Label(content_frame, text="Histórico de Atualizações", justify="left", font=("Consolas", 9, "bold")).pack(padx=10, pady=(15, 5))

text_frame = ttk.Frame(content_frame)
text_frame.pack(padx=10, pady=5, fill="both", expand=True)

text_widget = tk.Text(text_frame, wrap="word", font=("Consolas", 10), height=21)
text_scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=text_widget.yview)
text_widget.configure(yscrollcommand=text_scrollbar.set)

text_scrollbar.pack(side="right", fill="y")
text_widget.pack(side="left", expand=True, fill="both")

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

text_widget.insert("1.0", updates_text)
text_widget.config(state="disabled")

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
