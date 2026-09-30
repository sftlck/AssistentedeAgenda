import win32com.client
import tkinter as tk
import win32timezone
from tkinter import filedialog
from tkinter import ttk, messagebox
from datetime import time, datetime, timedelta, date
import random
import webbrowser
import calendar as cal
from dateutil.relativedelta import relativedelta
from getpass import getuser
from urllib.parse import quote
import socket

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.colors import HexColor
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, Paragraph

from reportlab.platypus import Image as RLImage
from reportlab.lib.enums import TA_CENTER, TA_LEFT
import tempfile
import os

import base64
from io import BytesIO
from PIL import Image, ImageTk

LOGO_BASE64 = """iVBORw0KGgoAAAANSUhEUgAAASwAAABECAYAAAA7rQj2AAAACXBIWXMAAAMkAAADJAG0HhTmAAAOqklEQVR4nO2d7XUbNxaGH/vsf7EDIRWYHWhcgZUKzFRgpYLQFUSpYKkKlqrAowoiVRCoA6kC7g9wwhGXHHzMCwxG2fccHckezL0XGODFxcXXh91uR0YsgQYw+78XwKdEWc+A3f/9Ajzu/37s/fslUNYC2AJXibZ0+HkvR4UFrryWHMprCVwkynvlUE5wKKN+eT2eeC8Fnb1m/9P/vw6p5f3EwVa7/90myiqN7hs2HMrGAJeRch72v+3RTzvKungo8tNvy3b/0/++9n9fcfiQgbCuez+pDS0VT7iCHCKuBe4jpxJnH9+B9UgZBldWKzQ2xeJurzsW3TdeMo3d4Brxdv9jJ7LhFJa4Mm0oUzbPHEi8RdcJdSidH3Bt2fI2X7Db7VQ/q91uZ3fTY7sbtrMR6rrx6Br6We5trQEx+Vjsdrt2GjMHsd25b6us0zE/i109beBlt9vd7sbn52ZXR352O5en5UcBEzZ7Jvw38W5uDiw8zxuhrpSebAHcAn8CX4S2jIGvzPpQDKVz4AvwA9cTN4V1r6mrDVwA34BN4vs3uPz8Th35AZendixhrXGVpJZMgZ9ETEFdx1jiGtQ3oQ0KhOajoU6y6uMKVydvC+hqcA37N8qHP0IQ0xGBq5+POKKqMT8XYwhrg/tQtcF6nhuRnlfCg/xwIKup4j1DCCWs66xWaPENl6/YRhuKW+rrrI8R06GuqLd+dnhOJawN8FVoiBK+j6TyEGIqg8FVhhp7LQgPWC9zGpEBn3DlriStbtKmNi/5FELr6Ao3nK21fnZ4TCGsG+olKxj+SMoG10ak3VJvZXjwJ/kbtQ8HT+ET6bGcY3RkNZdyCCGsFY6s5oBowlrixre14pnhYZoR6mkD066p2822gelMRhty4wuuox0D5XKYEnjF/22XzIesIIGwSgQyx8B6no/1sJ6AXzgM8Xww1Bnn68MGppvbcPAYa8YNDVvmQ1YQ5l1tchshho0hrGvqd4Vbz/PURvcAfN6/v4l4b52oryTawHRzJ6wL0r2sW+ZFVuAnrBUzzFMMYa1zWSGE7yPFNrp7HFE1xG+BMNQd6+tgA9PNnbAgfUX/HALsx/C1hXUJI4R4AAglrIZ5sLH1PA+dgr4DfsJV1jbRlnXieyUREufoYPKZUQyXxBHvgvkNmzrYgWfX1L0c4xQshBPWKpsZOhxv9D1GE/D+HziiWjF+b9oc1iy1EWnn0GGFoIlIe0u9s7s+tAPP5lA3j2EB/hWYuPYMvuKviGbg3dv9T8xC0CFMsfE7Fg+Ed0RNPjOKw0Skm8OQ/hSePM9rb8+n0EIYYakbX+cJWXQ77LfEx6+ecSS1Ie5YmhX+2VJ1hXjCfTAVob4QN+NrRHqfceXdetI1e50N+qFL6JBwI9ZbEnbg2ZjjiobwwOGImO4ImmMsSPPU7/ZygwirSVAwpPgGXcOLQVeAz7j40ibi3QXO7hvcx/Y19ibOtLN4ZVwcTQUjkvNIWGyv7f29ovzQzJBvRvwJ18G2hJ9HtuTtkoxjQjC8/UbrAVlm4FkKxrTp43w0vb+7PD/Sm90tSVgxQ5AcsLiKv414x+AKa8WhwfhWhht0XsEt05MV6OpAyukWG9y3+yGyIQTrDDLvOZyCEItT5RZTj/tQzvbeM65Nv/C2frenkx0QQliqYOtGJCcVq4i0BldpT8UwrOddZYWoZaGuKk8phAWuIj9RJvC/QB+7+oXp638OFK+fvlnCRqgrtUcoiQZn51+cr7TWI0PVuLuYwNRYoBuOpRIWlDtRdCWWVxtZmakNGAMfYRmRHt8ev6nR4HrxH/gP1Ws9z6f2RtRQeoxWKCsV1vN8JdT1nbrICrSEtRbKCkIpwrIiOWqscMTwg/Agq49IzAh7+rAiOWOh9BjHQHVEjB14ZtANO7tZ6NqgdByucO3BCGUOotSQsBXJUWHF4UjbmAoacmifqsLX4mEZkRw78n0jsAH8K8BVUK7rU0Jdrz4RPvs7Gj7Cek/DmwWuUF9IP3vbl48mQeY5tEJZY1BDHTDoZl7H7IaIwUYoq3Zc4E4lsWReZDxEWMpgqxXJSUFHVJbxZ2+XGg4+i+QoMMUJrcdYiWzwbd9SXQpyT53eFeTtCC9x4ZUtmYaJQ4SlDLZO6WE16C4JKEVYNXikMO2FHR26RbsKtAPPGpEOqHtGvETd+oKbaV8jPlN/aB2WirBeycPqL7ie19eTlZzlakR6luQpszerhgNgRHpjL+zoY4PO0x8ikkakA+oZzp/CC84DLHHF3G+4NnqDiMSHCMsoFOAqW65tDi1+QlISVltI1yV5jv/ovkMoaTUivSm9usFVcuVi0aFGo/p2/WvYa8WWcndiXgL/4bDTxY4RVmpImAshldmIdPniSsqYX07EfFdVHWgj0l7jvKq/0JLVHcNeXg2TC6WwxXm9JXGF+6a3jBgmDnlYtR+HDP5jNKDcMoM5EDzEX0+mwAq/t5a6kz8Ua4/uErOQteAFRxxT3DfwjcMwcRP78jnCMsnmlIX1PG+Euv6JhKUikFxD3FDc4T9yRYVWKCsn1jjimOK7XOCWFq1wHnVwfHPuhFVq1q60rpywgenmQsA+vOJf1FjD9iNDeh16JG1S4xr4M1GnAlcc1m4FdaTnCKuRmJMfJUnEep7PpYG3genmkh8f1vi/nRHpijkjv49uVjg1BvqEa7OxpPWI25w95d2EF7i8NwSQ1rmg+1wqa8mV5z5d7yXm18HkMqIg7gnbzzdlwN0wjqzADd1T16ptcKQ1JTrS8n6Hc4RlhMbkhPU8L7Vx14j05IaNSNtksqEUnghfIT/lav4lmtllO+LdDY60Ss8c9nGBm70cnEE8R1hzuCHFRyIltxYZkZ7cmGKGcArEDJGMUG8qYSlgR76/wdky9lSNMbjEE288RVjvZTg4xxXuuRHaoJTT/KXxQFw8Z+qAeyPS3QpkWJw9vzKdt/WNgU5kzoRlPc/nuMI9N0IJay75OcZ34oPPUy9pMAK9MbHJENwyrbd1Nh53irBMPjukaD3PS27eNkJduRAzgzU3wnoCPpN2JpMqr6mkofBkrUDGKZkN03hbq3MPThFWk80MDV5xlbMUiZQ8tC8Xnogrj5i0U+IZFywes1nciGxJWQfViHTnXF0/hbd1wZmO5NQ6LCNSek+egtwGyi0186P0Rr4LZfVxy3TDpFz4A82xM6rOpk14Zy77Fy2OXG9wXmyJPbMNJ/J1irDew516RiirFGGFrMguhTmsKZsDqfpgRHJK7V+8xTkMG/LXkZPLG46HhI1QYSuUFYs5xq9q2TRrpjYgEFeMt7UZb8bfmPOShlhdDfAzeWNbzan/PPawlGcCTYk5Lml4j4T1ef97weGbLNGdxXRNPTfTpMSwFPV0qpm8LYczy4p55MeEZURyrUhOKkpOVRuRnloIqxHJeeBt2W33vw06wlpRD2HFwjD9CvexeMHVlzWFjqo5HhKqGnorkpMKI5IT4imqYn5WJGcsjEjOOQK26NYNfaKeWJbJnP4caujo1rhJtuzIRVhTF2KpQ/sakR6YnuQ7lIirbEQ6QH+1fCpMZPpGpHfqttZB7emeHGL3CcvwPq71aoSySs0QTh3z66ME2W8HnsXiWihrDGLrghHprYWw1DiZr2PCyqqsEIxQli8fgzvLI2BFcsai1OyqRRcsviS9k7IiG8DF5WLqgxHoHHMbkRrqobnXw2pEitT7mmJhhLKs53kj1FUDjEjOM/6GtBHpgvRhoRXaAHHr6BQzazV5V4pFvH0U87CmLsRGKKtUXkwhPT6UXBdUy7BQ2cF+I6zhlp7cMhzOcG/QtpEFrvNRn+7RnvrP/rIGI1JkRXJSUerQPiUucQ34Fld+tqDuPhqRnDYgzQu6Cz0vcKSVQoKPaPeC/s7hqjLL4bz1JYchY8mOYcnwue39exTtmb87tD2ZC1w+V+i36pydcewTlmrxVyuSk4IF8504+IL+css74oZLRqTXBqbbol2TlUJYLfBVZEOHK8osprQBaXzeZ/9Go1q2ZJ39jt2QcOpDzFR4L/lQ4Stx082l15Qph4WxQe8OrdCG0mgD0jSZbVDjlQDCMkJlViQrBVMfxlYjQmeRGqHONjBdNyxUISWWZZn2WOBUhC6FqWVhbSgGTxZRe1hTB9xLbnoOTTM1bGA6VdnFBrFrCL5vhDaUQkjdU4ZISuAVz4jgvRGWEckJXd8ydX5DEGqjEemzkem36Hb9fyEtHxvqWrwbgpDvOjfv6gZPu1MTlhXJSUXp65qU3kEuhOZlqk7rhTq8LPU6otx4b4R1T4Cn2xGWKtj6T1nh3kEdg1EjZng2pZetJKxV4ntb6v6Wx7ABaUxmG1QIvkPyI+9nA+9UM4Q1H28SMxxUxTpSCUs1LPxEekNdMZ+h4XvxsGLukOQjZY9iyYnSAfcOLe588RphA9OZCXQeQ+llpQ7vXnBDyilvQA5B6Kxm7YQVRVagJSwrkpOKKZc0rJl+D+UptIHpGpG+McsDaohjgeusGuomLRuQpvYZwnvi75CUDglbkZxUGJGcFE/xBVeOtZFW7TOEfSiHhZeM68A60pp61HAONiBNrd7VK+6uw2sSTpr4yPtZ0qDaD2YT3+tIq5ZFiDFHj9QyS1zDsLDDI65cagzEtwFpTGYbUnCHsys57vuR+e6966MRympHvNuR1hS35R4jpgOZ8m6+PpQTGI1ARhfT+kw93tYTYeW8wRHE1HjF2fETblIj2qvq4yOaYcyvTO9hqaDIxy2uJ/nOdBU9pvEr6sAfjCesR+poZMdocd/zF8p70K97nd9xV2s1Ee+ucDbfUb4e3u91m70dViH0w263686zSdk4CofLFaeEwVX2sd7iE3nG/kvcR1uSZ0f8E67nsvufljjyGFsHNui2t4y1pcMN+TpRg/O8rtF9z44IW9y3fORwNI0KBlcHlzjiW6Dxrl9xtrb739n44MNut8sluzQM48btXSUpgSWHu/pCG+axfSXt/T+G0X3HJiBtn4Qs08+uw9u2E1on2/1vNakO4r+deGrI1z4RLAAAAABJRU5ErkJggg=="""
bool_agenda = 1

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
    f'  ¯|_(ツ)_|¯ ',
    f'  *_*  ',
    f'  Buenas tardes! ', 
    f'  Vai um chimas? '
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


class LabSelector:
    """Master tab with lab selection buttons"""
    
    def __init__(self, parent_notebook, cor_fundo):
        self.cor_fundo = cor_fundo
        self.frame = ttk.Frame(parent_notebook)
        parent_notebook.add(self.frame, text=" Agenda de Serviços ")
        self._create_lab_buttons()



    def show_loading_screen(self, message="Carregando..."):
        
        self.loading_popup = tk.Toplevel(self.frame)
        self.loading_popup.geometry("550x130")
        
        cor_fundo2 = _from_rgb((0,0,0))
        self.loading_popup.configure(bg=cor_fundo2)
        self.loading_popup.title("")
        self.loading_popup.overrideredirect(True)
        self.loading_popup.attributes('-topmost', True)
        
        # Center on parent
        self.loading_popup.update_idletasks()
        x = self.frame.winfo_rootx() + (self.frame.winfo_width() // 2) - 175
        y = self.frame.winfo_rooty() + (self.frame.winfo_height() // 2) - 65
        self.loading_popup.geometry(f"+{x}+{y}")
        
        frame = tk.Frame(self.loading_popup, bg=cor_fundo2, highlightbackground="#FFFFFF", highlightthickness=2)
        frame.pack(fill='both', expand=True, padx=2, pady=2)

        if bool_agenda == 1:
            tk.Label(frame, text=f"Oioi, {username}! O Assistente de Agenda está carregando:", font=('Segoe UI', 10), background='black', fg="#FFFFFF").pack(pady=(15, 5))
        elif bool_agenda == 0:
            tk.Label(frame, text=f"Oioi, {username}! O Assistente de Dados está carregando:", font=('Segoe UI', 10), background='black', fg="#FFFFFF").pack(pady=(15, 5))

        #tk.Label(frame, text=f"Laboratório de {self.lab_name}", font=('Segoe UI', 13, 'bold'), bg=self.cor_fundo, fg="#FFFFFF").pack(pady=(5, 5))
        
        # Message
        tk.Label(frame, text=message, font=('Segoe UI', 11),bg='white', fg='black').pack(pady=(5, 5))
        
        self.loading_popup.grab_set()
        self.loading_popup.update()


    def hide_loading_screen(self):
        """Hide the loading overlay"""
        if hasattr(self, 'loading_popup') and self.loading_popup.winfo_exists():
            self.loading_popup.grab_release()
            self.loading_popup.destroy()


    def _create_lab_buttons(self):
        main = tk.Frame(self.frame, bg=self.cor_fundo)
        main.pack(fill='both', expand=True, padx=20, pady=35)
        
        # Logo
        try:
            img_data = base64.b64decode(LOGO_BASE64)
            img = Image.open(BytesIO(img_data))
            w = 45
            h = int(w * 4.411764)
            img = img.resize((h, w), Image.LANCZOS)
            self.logo_image = ImageTk.PhotoImage(img)
            tk.Label(main, image=self.logo_image, bg=self.cor_fundo).pack(pady=(0, 10))
        except Exception as e:
            print(f"Logo error: {e}")

        ttk.Label(main, text=f"{buenas}", font=("Segoe UI", 10), background='black').pack(pady=(20, 0))
        
        grid = tk.Frame(main, bg=self.cor_fundo)
        grid.pack(expand=True)
        
        row, col = 0, 0
        max_cols = 1
        
        style = ttk.Style()
        #style.configure('Lab.TButton', font=('Segoe UI', 10), padding=1)
        style.configure('Lab.TButton', 
            font=('Segoe UI', 10),
            padding=2,
            borderwidth=2,           # Border thickness
            relief='groove',         # raised, sunken, flat, ridge, solid, groove
            background='#1B4B9F',    # Background color
            foreground='black',      # Text color
            anchor='center',         # Text alignment: center, w, e, n, s, nw, etc.
            width=45                # Width in characters
            )

        for lab_code, lab_config in LABS.items():
            
            btn = ttk.Button(grid, text=f"Laboratório de {lab_config['name']}", style='Lab.TButton',command=lambda lc=lab_code, lcfg=lab_config: self.open_lab(lc, lcfg),cursor='hand2')
            
            #btn.bind('<Button-3>', lambda e, lcfg=lab_config: self.open_teams_chat(lcfg))
            btn.bind('<Button-3>', lambda e, lc=lab_code, lcfg=lab_config: self.show_lab_context_menu(e, lc, lcfg))
            btn.grid(row=row, column=col, padx=10, pady=2, sticky='ew')
            
            col += 1
            if col >= max_cols:
                col = 0
                row += 1
        btn_export_all = ttk.Button(main, text="Exportar Programação Geral de Serviços Metrológicos",command=self.export_all_to_pdf, cursor='hand2')
        btn_export_all.pack(pady=(20, 0))   

    def export_lab_to_pdf(self):
        """Export current lab's services to PDF - for context menu"""
        lab_name = self.lab_config.get('name', '')
        
        filename = filedialog.asksaveasfilename(defaultextension=".pdf",filetypes=[("PDF files", "*.pdf")],initialfile=f"Programação {lab_name} {datetime.now().strftime('%d-%m-%Y')}.pdf",title=f"Programação - Laboratório de {lab_name}")
        
        if not filename:
            return
        
        try:
            doc = SimpleDocTemplate(filename, pagesize=landscape(A4),leftMargin=15*mm, rightMargin=15*mm,topMargin=15*mm, bottomMargin=15*mm)
            
            elements = []
            styles = getSampleStyleSheet()
            
            # Logo
            try:
                logo_bytes = base64.b64decode(LOGO_BASE64)
                logo_img = RLImage(BytesIO(logo_bytes), width=40*mm, height=20*mm)
                elements.append(logo_img)
            except:
                pass
            
            elements.append(Paragraph(f"Programação de Serviços - Laboratório de {lab_name}", styles['Title']))
            elements.append(Paragraph(f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}", styles['Normal']))
            elements.append(Spacer(1, 10*mm))
            
            headers = ["OS", "Item", "Tipo", "Especificação", "Código", "Duração", "Status"]
            data = [headers]
            
            for item in self.tree.get_children():
                values = self.tree.item(item)['values']
                data.append([
                    str(values[0]) if values[0] else "",                        # OS
                    str(values[1]) if values[1] else "",                        # Item
                    str(values[2])[:50] if values[2] else "",                   # Especificação
                    str(values[4]) if values[4] else "",                        # Código
                    str(values[5]) if values[5] else "",                        # Duração
                    str(values[6]) if len(values) > 6 and values[6] else "",    # Status
                ])
            
            col_widths = [15*mm,15*mm,15*mm, 35*mm, 70*mm, 35*mm, 30*mm, 60*mm]

            table = Table(data, colWidths=col_widths, repeatRows=1)
            
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), HexColor('#1B4B9F')),
                ('TEXTCOLOR', (0, 0), (-1, 0), HexColor('#FFFFFF')),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 7),
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 1), (-1, -1), 6),
                ('GRID', (0, 0), (-1, -1), 0.5, HexColor('#CCCCCC')),
                ('ROWBACKGROUNDS', (0, 1), (-1, -1), [HexColor('#FFFFFF'), HexColor('#F5F7FA')]),
            ]))
            
            elements.append(table)
            elements.append(Spacer(1, 10*mm))
            elements.append(Paragraph(f"Total: {len(data) - 1} serviços", styles['Normal']))
            
            doc.build(elements)
            os.startfile(filename)
            
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao gerar PDF:\n{str(e)}")

    def export_all_to_pdf(self):
        
        filename = filedialog.asksaveasfilename(defaultextension=".pdf",filetypes=[("PDF files", "*.pdf")],initialfile=f"Programação Geral {datetime.now().strftime('%d-%m-%Y_%Hh%M')}.pdf",title=f"Programação Geral de Serviços Metrológicos {datetime.now().strftime('%d-%m-%Y')}.pdf")
        
        if not filename:
            return
        self.show_loading_screen("Programação Geral de Serviços Metrológicos")
        try:
            doc = SimpleDocTemplate(filename, pagesize=landscape(A4),leftMargin=15*mm, rightMargin=15*mm,topMargin=15*mm, bottomMargin=15*mm)
            
            elements = []
            styles = getSampleStyleSheet()
            
            # Custom lab header style
            lab_header_style = ParagraphStyle('LabHeader', parent=styles['Heading2'],fontSize=10, textColor=HexColor('#1B4B9F'), spaceBefore=8*mm, spaceAfter=3*mm,borderPadding=(0, 0, 2, 0))
            
            # Logo
            #try:
            #    logo_bytes = base64.b64decode(LOGO_BASE64)
            #    logo_img = RLImage(BytesIO(logo_bytes), width=25*mm, height=12*mm)
            #    elements.append(logo_img)
            #except Exception as e:
            #    print(f"Logo error: {e}")
            
            #elements.append(Paragraph("Programação Geral de Serviços Metrológicos", styles['Title']))
            #elements.append(Paragraph(f"Exportado em: {datetime.now().strftime('%d/%m/%Y - %H:%M')}", styles['Normal']))
            
            total_all = 0
            k = 0
            for lab_code, lab_config in LABS.items():
                lab_data = []
                has_data = False
                
                try:
                    conn = win32com.client.Dispatch("ADODB.Connection")
                    rs = win32com.client.Dispatch("ADODB.Recordset")
                    conn.Open(STR_CONN_LINKED)
                    
                    sql = f"""
                                            
                    SELECT 
                        CONVERT(VARCHAR(10), os.receiving_date , 103) as 'Recebimento',
                        CONVERT(VARCHAR(10), os.date_finished , 103) as 'Entrega',
                        os.code AS 'OS',
                        i.code AS 'Item',
                        it.name as 'Tipo',
                        sm.specification AS 'Especificação',
                        --sm.description AS 'Descrição',
                        sm.code AS 'Código',
                        sm.execution_time AS 'execution_time',
                        sm.id AS 'service_mode_id'

                    FROM instruments_services iss

                    LEFT JOIN orders_services AS os ON os.id = iss.id_service_order 
                    LEFT JOIN service_modes AS sm ON sm.id = iss.id_service
                    LEFT JOIN instruments AS i ON i.id = iss.id_instrument AND i.id_service_order = os.id
                    LEFT JOIN budgets AS b ON b.id_order_service = os.id
                    LEFT JOIN instrument_types it on it.id = i.id_instrument_type

                    WHERE os.removed = 0
                    AND iss.removed = 0
                    AND sm.removed = 0
                    AND i.removed = 0
                    AND b.is_last_revision = 1
                    
                    AND (i.id_current_sector = {lab_config['sector_id']} or i.id_current_sector = 1)
                    AND sm.code LIKE '%631{lab_code.lower()}%'
                    --AND (i.id_current_sector = 5 or i.id_current_sector = 1)
                    --AND sm.code LIKE '%631CD%'
                    
                    ORDER BY os.receiving_date ASC
                    """

                    rs.Open(sql, conn)
                    
                    if not rs.EOF:
                        rs.MoveFirst()
                        while not rs.EOF:
                            duration = self._convert_hhmm_to_minutes(rs.Fields('execution_time').Value)
                            spec = str(rs.Fields('Especificação').Value or "")
                            spec_display = spec if len(spec) < 20 else spec[:20] + '...'
                            
                            type = str(rs.Fields('Tipo').Value or "")
                            #type_display = type if len(type) < 10 else type[:10] + '...'
                            
                            os_parts = str(rs.Fields('OS')).split('/')
                            os_code = os_parts[0].lstrip('0')
                            os_code = f"{os_code}/{os_parts[1]}"
                            
                            #item = str(rs.Fields('Item').Value or ""),
                            #item_display = item if len(item) < 10 else item[:10] + '...'


                            # In the loop where you build lab_data, wrap text in Paragraphs:
                            lab_data.append([
                                Paragraph(str(rs.Fields('Recebimento').Value or ""),    styles['Normal']),
                                Paragraph(str(rs.Fields('Entrega').Value or ""),        styles['Normal']),
                                Paragraph(str(os_code),                                 styles['Normal']),
                                Paragraph(str(rs.Fields('Item').Value or ""),           styles['Normal']),
                                Paragraph(type,                                         styles['Normal']),
                                Paragraph(spec,                                         styles['Normal']),  # Full text, no truncation
                                Paragraph(str(rs.Fields('Código').Value or ""),         styles['Normal']),
                                Paragraph(f"{duration} min",                            styles['Normal']),
                            ])
                            
                            rs.MoveNext()
                        has_data = True
                    
                    rs.Close()
                    conn.Close()
                except:
                    lab_data.append(["Erro ao carregar", "", "", "", ""])
                    
                    self.hide_loading_screen()
                    has_data = True
                
                if has_data:
                    # Lab header
                    count = len(lab_data)
                    total_all += count
                    
                    elements.append(Paragraph("Programação Geral de Serviços Metrológicos", styles['Title']))
                    elements.append(Paragraph(f"Laboratório de {lab_config['name']}", styles['Title']))
                    elements.append(Paragraph("[Em desenvolvimento]", styles['Title']))
                    elements.append(Paragraph(f"{count} {'serviço' if count == 1 else 'serviços'}",lab_header_style))

                    # Table for this lab
                    headers = ["Recebimento", "Entrega" , "OS", "Item", "Tipo" ,"Especificação", "Código", "Duração"]
                    table_data = [headers] + lab_data
                    
                    col_widths = [23*mm, 27*mm, 23*mm, 30*mm, 65*mm, 35*mm, 23*mm, 20*mm]
                    table = Table(table_data, colWidths=col_widths, repeatRows=1)

                    table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), HexColor('#1B4B9F')),
                    ('TEXTCOLOR', (0, 0), (-1, 0), HexColor('#FFFFFF')),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 8),
                    ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                    ('FONTSIZE', (0, 1), (-1, -1), 7),
                    ('GRID', (0, 0), (-1, -1), 0.5, HexColor('#CCCCCC')),
                    ('ROWBACKGROUNDS', (0, 1), (-1, -1), [HexColor('#FFFFFF'), HexColor('#F5F7FA')]),
                    ('VALIGN', (0, 0), (-1, -1), 'TOP'),       # Align top
                    ('WORDWRAP', (0, 0), (-1, -1), True),       # ← Enable word wrap
                    ]))
                    
                    elements.append(table)
                    elements.append(PageBreak())
            
            # Footer
            elements.append(Spacer(1, 10*mm))
            elements.append(Paragraph(f"Total geral: {total_all} serviços em {len(LABS)} laboratórios", styles['Normal']))
            
            doc.build(elements)
            os.startfile(filename)
            
            self.hide_loading_screen()
            
        except Exception as e:
            
            self.hide_loading_screen()
            messagebox.showerror("Erro", f"Falha ao gerar PDF:\n{str(e)}")

    @staticmethod
    def _convert_hhmm_to_minutes(time_str):
        if not time_str:
            return 0
        try:
            parts = str(time_str).strip().split(':')
            return int(parts[0]) * 60 + int(parts[1])
        except:
            return 0

    def show_lab_context_menu(self, event, lab_code, lab_config):
        """Right-click context menu for lab button"""
        context_menu = tk.Menu(self.frame, tearoff=0)
        
        chat_id = lab_config.get('teams_chat_id', '')
        if chat_id:
            context_menu.add_command(label=f"Iniciar chat do Teams com Laboratório de {lab_config['name']}",command=lambda lcfg=lab_config: self.open_teams_chat(lcfg))
            context_menu.add_command(label=f"Exportar Programa de Serviços para Laboratório de {lab_config['name']} [Em desenvolvimento]",command=lambda lcfg=lab_config: self.export_lab_to_pdf)
        
        planilha_link = lab_config.get('sharepoint_planilha', '')
        #if planilha_link:
        #    context_menu.add_command(label=f"📊 Abrir Planilha de Cálculo",command=lambda lcfg=lab_config: webbrowser.open(lcfg.get('sharepoint_planilha', '')))
        
        context_menu.post(event.x_root, event.y_root)

    def open_teams_chat(self, lab_config):
        chat_id = lab_config.get('teams_chat_id', '')
        if chat_id:
            link = f"https://teams.cloud.microsoft/l/chat/19:{chat_id}@thread.v2/conversations?context=%7B%22contextType%22%3A%22chat%22%7D"
            webbrowser.open(link)   

    def open_lab(self, lab_code, lab_config):
        """Open lab calendar + services in a new window"""
        lab_window = tk.Toplevel(self.frame)
        lab_window.transient(self.frame)
        lab_window.geometry("1150x700")
        lab_window.minsize(900, 600)
        lab_window.title(f"Calendário/Lista de Serviços - Laboratório de {lab_config['name']}")
        lab_window.configure(bg=self.cor_fundo)
        
        lab_notebook = ttk.Notebook(lab_window)
        lab_notebook.pack(fill='both', expand=True)
        
        ServiceScheduler(lab_notebook, STR_CONN, self.cor_fundo,id_sector=lab_config['sector_id'],lab_name=lab_config['name'],lab_config=lab_config)
        ServiceSchedulerLinkedDirect(lab_notebook, STR_CONN_LINKED, STR_CONN, self.cor_fundo, lab_code=lab_code, lab_config=lab_config)

class ProcedureManager:
    """Gerenciador de Procedimentos - Update procedure_name and procedure_url"""
    
    def __init__(self, parent, str_conn, str_conn_primary, cor_fundo, lab_code="", lab_config=None):
        self.str_conn = str_conn
        self.str_conn_primary = str_conn_primary
        self.cor_fundo = 'black'
        self.lab_code = lab_code
        self.lab_config = lab_config or {}
        self.all_rows = []
        
        self._open_window(parent)
    
    def _sanitize_input(self, text):
        """Remove dangerous characters to prevent SQL injection"""
        if not text:
            return ""
        # Remove single quotes, semicolons, double dashes
        dangerous = ["'", ";", "--", "/*", "*/", "xp_", "sp_", "exec", "drop", "delete", "truncate"]
        sanitized = text
        for d in dangerous:
            sanitized = sanitized.replace(d, "")
        return sanitized.strip()
    
    def _open_window(self, parent):
        lab_name = self.lab_config.get('name', self.lab_code)
        
        self.window = tk.Toplevel(parent)
        self.window.geometry("900x550")
        self.window.minsize(800, 450)
        self.window.configure(bg=self.cor_fundo)
        self.window.title(f"Lista de Tempos, Serviços e ITs - Laboratório de {lab_name}")
        self.window.transient(parent)
        
        # ===== TOP FRAME =====
        top_frame = tk.Frame(self.window, bg=self.cor_fundo)
        top_frame.pack(fill='x', padx=15, pady=(15, 5))
        
        tk.Label(top_frame, text=f"Lista de Tempos, Serviços e ITs",font=('Segoe UI', 13, 'bold'), bg=self.cor_fundo, fg='white').pack(anchor='w')
        
        # ===== SEARCH =====
        search_frame = tk.Frame(self.window, bg=self.cor_fundo)
        search_frame.pack(fill='x', padx=15, pady=(10, 5))
        
        #tk.Label(search_frame, text="Buscar:", font=('Segoe UI', 9),bg=self.cor_fundo, fg='white').pack(side='left', padx=(0, 5))
        
        #self.search_var = tk.StringVar()
        #self.search_var.trace('w', lambda *args: self._filter_treeview())
        #ttk.Entry(search_frame, textvariable=self.search_var, width=40).pack(side='left', padx=5)
        
        # ===== TREEVIEW =====
        tree_frame = ttk.Frame(self.window)
        tree_frame.pack(fill='both', expand=True, padx=15, pady=10)
        
        tree_scroll_y = ttk.Scrollbar(tree_frame)
        tree_scroll_y.pack(side='right', fill='y')
        
        tree_scroll_x = ttk.Scrollbar(tree_frame, orient="horizontal")
        tree_scroll_x.pack(side='bottom', fill='x')
        
        columns = ("Código", "Duração", "Serviço", "Descrição","Instrução de Trabalho", "Link")
        
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings",yscrollcommand=tree_scroll_y.set,xscrollcommand=tree_scroll_x.set,selectmode="browse")
        
        tree_scroll_y.config(command=self.tree.yview)
        tree_scroll_x.config(command=self.tree.xview)
        
        self.tree.heading("Código",                 text="Código")
        self.tree.heading("Duração",                text="Duração")
        self.tree.heading("Serviço",                text="Serviço")
        self.tree.heading("Descrição",              text="Descrição")
        self.tree.heading("Instrução de Trabalho",  text="Instrução de Trabalho")
        self.tree.heading("Link",                   text="Link")
        
        self.tree.column("Código",                  width=70,minwidth=65)
        self.tree.column("Duração",                 width=50,minwidth=50)
        self.tree.column("Serviço",                 width=200)
        self.tree.column("Descrição",               width=100)
        self.tree.column("Instrução de Trabalho",   width=120,minwidth=120)
        self.tree.column("Link",                    width=250)
        
        self.tree.pack(expand=True, fill='both')
        self.tree.bind('<<TreeviewSelect>>', self._on_select)
        self.tree.bind('<Double-1>', lambda e: self._edit_selected())
        
        # ===== EDIT FRAME =====
        edit_frame = tk.LabelFrame(self.window, text=" Editar Dados [Em desenvolvimento] ", font=('Segoe UI', 10, 'bold'),bg=self.cor_fundo, fg="#FFFFFF",padx=10, pady=10)
        edit_frame.pack(fill='x', padx=15, pady=(5, 10))
        
        # Current info
        self.lbl_current = tk.Label(edit_frame, text="Selecione um serviço na lista acima para editar.",font=('Segoe UI', 9), bg=self.cor_fundo, fg='#B7D5F5', anchor='w', justify='left')
        self.lbl_current.pack(anchor='w', pady=(0, 5))
        
        # Procedure Name
        name_frame = tk.Frame(edit_frame, bg=self.cor_fundo)
        name_frame.pack(fill='x', pady=3)
        
        tk.Label(name_frame, text="Nome da Instrução de Trabalho:", font=('Segoe UI', 9),bg=self.cor_fundo, fg='white', width=25, anchor='w').pack(side='left')
        
        self.proc_name_var = tk.StringVar()
        self.entry_name = ttk.Entry(name_frame, textvariable=self.proc_name_var, width=60)
        self.entry_name.pack(side='left', padx=5, fill='x', expand=True)
        
        # Procedure URL
        url_frame = tk.Frame(edit_frame, bg=self.cor_fundo)
        url_frame.pack(fill='x', pady=3)
        
        tk.Label(url_frame, text="Link do documento (EmDoc):", font=('Segoe UI', 9),bg=self.cor_fundo, fg='white', width=25, anchor='w').pack(side='left')
        
        self.proc_url_var = tk.StringVar()
        self.entry_url = ttk.Entry(url_frame, textvariable=self.proc_url_var, width=55)
        self.entry_url.pack(side='left', padx=5, fill='x', expand=True)
        
        # Buttons
        btn_frame = tk.Frame(edit_frame, bg=self.cor_fundo)
        btn_frame.pack(fill='x', pady=(10, 0))
        
        self.btn_save = ttk.Button(btn_frame, text="Salvar", command=self._save_procedure, cursor='hand2')
        self.btn_save.pack(side='left', padx=5)
        self.btn_save.config(state='disabled')
        
        ttk.Button(btn_frame, text="Atualizar Lista", command=self._load_data).pack(side='left', padx=5)
        
        
        #self.btn_clear = ttk.Button(btn_frame, text="Limpar", command=self._clear_fields, cursor='hand2')
        #self.btn_clear.pack(side='left', padx=5)
        
        #ttk.Button(btn_frame, text="Fechar", command=self.window.destroy).pack(side='right', padx=5)
        
        # Status
        self.status_label = tk.Label(edit_frame, text="", font=('Segoe UI', 9),bg=self.cor_fundo, fg='#28A745')
        self.status_label.pack(anchor='w', pady=(5, 0))
        
        # Load data
        self._load_data()
        self._clear_fields()
    
    def _load_data(self):
        """Load services with their procedure info"""
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        self.all_rows.clear()
        sector_id = self.lab_config.get('sector_id', 0)
        
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            rs = win32com.client.Dispatch("ADODB.Recordset")
            conn.Open(self.str_conn_primary)
            
            # Get services that exist in Service_Modes_Local
            sql = f"""

                SELECT 
                sml.id, 
                sml.code, 
                sml.execution_time_minutes, 
                sml.specification,
                sml.description, 
                sml.procedure_name, 
                sml.procedure_url      

                FROM [IST-PGE].dbo.Service_Modes_Local sml

                WHERE sml.code LIKE '%631{self.lab_code.lower()}%'

                ORDER BY sml.code ASC
            """
            rs.Open(sql, conn)
            
            if not rs.EOF:
                rs.MoveFirst()
                while not rs.EOF:
                    row_data = {
                        'id': rs.Fields('id').Value,
                        'code': rs.Fields('code').Value if rs.Fields('code').Value else "",
                        'execution_time_minutes': f"{rs.Fields('execution_time_minutes').Value} min"  if rs.Fields('execution_time_minutes').Value else "",
                        'specification': rs.Fields('specification').Value if rs.Fields('specification').Value else "",
                        'description': rs.Fields('description').Value if rs.Fields('description').Value else "",
                        'procedure_name': rs.Fields('procedure_name').Value if rs.Fields('procedure_name').Value else "",
                        'procedure_url': rs.Fields('procedure_url').Value if rs.Fields('procedure_url').Value else "",
                    }
                    self.all_rows.append(row_data)
                    
                    proc_name = row_data['procedure_name'] if row_data['procedure_name'] else "(vazio)"
                    proc_url = row_data['procedure_url'] if row_data['procedure_url'] else "(vazio)"
                    
                    self.tree.insert("", "end", values=(
                        row_data['code'],
                        row_data['execution_time_minutes'],
                        row_data['specification'],
                        row_data['description'],
                        proc_name,
                        proc_url
                    ))
                    rs.MoveNext()
            
            rs.Close()
            conn.Close()
            
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao carregar dados:\n{str(e)}")
    
    def _filter_treeview(self):
        """Filter treeview by search"""
        search_term = self.search_var.get().lower()
        
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        for row in self.all_rows:
            if not search_term or \
               search_term in str(row['code']).lower() or \
               search_term in str(row['execution_time_minutes']).lower() or \
               search_term in str(row['specification']).lower() or \
               search_term in str(row['description']).lower() or \
               search_term in str(row['procedure_name']).lower():
                
                proc_name = row['procedure_name'] if row['procedure_name'] else "(vazio)"
                proc_url = row['procedure_url'] if row['procedure_url'] else "(vazio)"
                
                self.tree.insert("", "end", values=(
                    row['code'], 
                    row['execution_time_minutes'],
                    row['specification'], 
                    row['description'], 
                    proc_name, 
                    proc_url
                ))
    
    def _on_select(self, event):
        """Populate fields when a row is selected"""
        selected = self.tree.selection()
        if not selected:
            self._clear_fields()
            return
        
        all_items = self.tree.get_children()
        idx = all_items.index(selected[0])
        
        if idx < len(self.all_rows):
            row = self.all_rows[idx]
            
            #if len(row['specification']) > 60:
            #    self.lbl_current.config(text=f"Editando: {row['code']} - {row['specification'][:60]}...")
            #else:
            #    self.lbl_current.config(text=f"Editando: {row['code']} - {row['specification']}")

            self.proc_name_var.set(row['procedure_name'])
            self.proc_url_var.set(row['procedure_url'])
            self.btn_save.config(state='normal')
            
            # Store selected row index
            self._selected_idx = idx
    
    def _clear_fields(self):
        """Clear edit fields"""
        #self.lbl_current.config(text="Selecione um serviço na lista acima para editar.")
        self.proc_name_var.set("")
        self.proc_url_var.set("")
        self.btn_save.config(state='disabled')
        self._selected_idx = None
        self.status_label.config(text="")
    
    def _edit_selected(self):
        """Focus on name field when double-clicking"""
        if self.btn_save.cget('state') == 'normal':
            self.entry_name.focus_set()
    
    def _save_procedure(self):
        """Save procedure info to database"""
        if self._selected_idx is None:
            return
        
        row = self.all_rows[self._selected_idx]
        service_id = row['id']
        service_code = row['code']
        
        # Sanitize inputs
        proc_name = self._sanitize_input(self.proc_name_var.get().strip())
        proc_url = self._sanitize_input(self.proc_url_var.get().strip())
        
        # Validate URL format (basic check)
        if proc_url and not (proc_url.startswith('http://') or proc_url.startswith('https://')):
            if not messagebox.askyesno("Aviso", 
                "A URL não começa com http:// ou https://. Deseja salvar mesmo assim?"):
                return
        
        # Confirm
        confirm = messagebox.askyesno("Confirmar",
            f"Atualizar procedimento para:\n\n"
            f"Código: {service_code}\n"
            f"Nome: {proc_name if proc_name else '(vazio)'}\n"
            f"URL: {proc_url if proc_url else '(vazio)'}\n\n"
            f"Continuar?")
        
        if not confirm:
            return
        
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            conn.Open(self.str_conn_primary)
            
            sql = f"""
                UPDATE [IST-PGE].dbo.Service_Modes_Local
                SET procedure_name = {'NULL' if not proc_name else f"'{proc_name.replace(chr(39), chr(39)+chr(39))}'"},
                    procedure_url = {'NULL' if not proc_url else f"'{proc_url.replace(chr(39), chr(39)+chr(39))}'"}
                WHERE id = {service_id}
            """
            conn.Execute(sql)
            conn.Close()
            
            self.status_label.config(text=f"Instrução de trabalho de {service_code} atualizado com sucesso!")
            
            # Refresh data
            self._load_data()
            self._clear_fields()
            
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao salvar:\n{str(e)}")
            self.status_label.config(text="Erro ao salvar!")


class ServiceScheduler: #CALENDÁRIO DE SERVIÇOS DO SETOR
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
        self.refresh_calendar()

    def open_procedure_manager(self):
        """Open the procedure management window"""
        ProcedureManager(
            self.frame_certificados,
            self.str_conn,
            self.str_conn_primary,
            self.cor_fundo,
            lab_code=self.lab_code,
            lab_config=self.lab_config
        )

    def on_tab_changed(self, event):
        notebook = event.widget
        current_tab = notebook.select()
        current_tab_text = notebook.tab(current_tab, "text")
        
        #if self.lab_name in current_tab_text:
        #    self.refresh_calendar()
        if current_tab_text:
            self.refresh_calendar()


    def show_loading_screen(self, message="Carregando..."):
        
        self.loading_popup = tk.Toplevel(self.frame_certificados)
        self.loading_popup.geometry("550x130")
        cor_fundo2 = _from_rgb((0,0,0))
        self.loading_popup.configure(bg=cor_fundo2)
        self.loading_popup.title("")
        self.loading_popup.overrideredirect(True)
        self.loading_popup.attributes('-topmost', True)
        
        # Center on parent
        self.loading_popup.update_idletasks()
        x = self.frame_certificados.winfo_rootx() + (self.frame_certificados.winfo_width() // 2) - 175
        y = self.frame_certificados.winfo_rooty() + (self.frame_certificados.winfo_height() // 2) - 65
        self.loading_popup.geometry(f"+{x}+{y}")
        
        frame = tk.Frame(self.loading_popup, bg=cor_fundo2, highlightbackground="#FFFFFF", highlightthickness=2)
        frame.pack(fill='both', expand=True, padx=2, pady=2)

        
        if bool_agenda == 1:
            tk.Label(frame, text=f"Oioi, {username}! O Assistente de Agenda está carregando:", font=('Segoe UI', 10), background='black', fg="#FFFFFF").pack(pady=(15, 5))
        elif bool_agenda == 0:
            tk.Label(frame, text=f"Oioi, {username}! O Assistente de Dados está carregando:", font=('Segoe UI', 10), background='black', fg="#FFFFFF").pack(pady=(15, 5))


        #tk.Label(frame, text=f"Laboratório de {self.lab_name}", font=('Segoe UI', 13, 'bold'), bg=self.cor_fundo, fg="#FFFFFF").pack(pady=(5, 5))
        
        # Message
        tk.Label(frame, text=message, font=('Segoe UI', 11),bg='white', fg='black').pack(pady=(5, 5))
        
        self.loading_popup.grab_set()
        self.loading_popup.update()


    def hide_loading_screen(self):
        """Hide the loading overlay"""
        if hasattr(self, 'loading_popup') and self.loading_popup.winfo_exists():
            self.loading_popup.grab_release()
            self.loading_popup.destroy()


    def reschedule_selected_services_manual(self, tree, selected_items):
        """Manually reschedule only selected services"""
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
        confirm = messagebox.askyesno("Info agendamento", 
            f"Deseja reagendar manualmente {'o serviço selecionado' if count == 1 else f'os {count} serviços selecionados'} de {date_str}?")
        if not confirm:
            return
        
        # Open manual popup — but pass selected_schedule_ids instead of full day
        self._open_manual_reschedule_popup_selected(calendar_key, selected_schedule_ids)


    def _open_manual_reschedule_popup_selected(self, calendar_key, selected_schedule_ids):
        """Manual reschedule popup for selected services only"""
        popup = tk.Toplevel(self.frame_certificados)
        popup.geometry("380x280")
        popup.configure(bg=self.cor_fundo)
        popup.title("Reagendamento Manual")
        
        count = len(selected_schedule_ids)
        date_str = f"{calendar_key[0]:02d}/{calendar_key[1]:02d}/{calendar_key[2]}"
        
        tk.Label(popup, text=f"Reagendar {count} {'serviço selecionado' if count == 1 else 'serviços selecionados'} de {date_str}",
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
        
        tk.Label(popup, text="Horário de início (HH:MM):", font=('Segoe UI', 9),bg=self.cor_fundo, fg='white').pack(anchor='w', padx=30, pady=(10, 0))
        
        time_var = tk.StringVar(value="08:00")
        ttk.Entry(popup, textvariable=time_var, width=10).pack(anchor='w', padx=30, pady=(2, 5))
        
        tk.Label(popup, text="⚠ Os serviços selecionados serão agendados\n   em sequência a partir desta data e horário.",font=('Segoe UI', 8), bg=self.cor_fundo, fg='#FFD700').pack(pady=(10, 5))
        
        btn_frame = tk.Frame(popup, bg=self.cor_fundo)
        btn_frame.pack(pady=15)
        
        ttk.Button(btn_frame, text=" Confirmar Reagendamento ",command=lambda: self._execute_manual_reschedule_selected(calendar_key, popup, selected_schedule_ids,day_var, month_var, year_var, time_var)).pack(side='left', padx=5)
        ttk.Button(btn_frame, text=" Cancelar ", command=popup.destroy).pack(side='left', padx=5)


    def _execute_manual_reschedule_selected(self, calendar_key, popup, selected_schedule_ids, day_var, month_var, year_var, time_var):
        """Execute manual reschedule for selected services only"""
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
            
            popup.destroy()
            
            count = len(selected_schedule_ids)
            
            conn = win32com.client.Dispatch("ADODB.Connection")
            conn.Open(self.str_conn)
            
            id_list = ','.join(str(sid) for sid in selected_schedule_ids)
            conn.Execute(f"UPDATE [IST-PGE].dbo.Service_Schedule SET status = 'PENDING', scheduled_start = NULL, scheduled_end = NULL, updated_at = GETDATE() WHERE id IN ({id_list})")
            conn.Execute(f"DELETE FROM [IST-PGE].dbo.Time_Slots WHERE schedule_id IN ({id_list})")
            conn.Close()
            
            self.show_loading_screen(f"Agendamento de {count} {'item' if count == 1 else 'itens'} em Laboratório de {self.lab_name}")
            self.schedule_all_pending(from_date=new_datetime)
            
            self.refresh_calendar()
            self.hide_loading_screen()
            
            new_date_br = new_date.strftime('%d/%m/%Y')
            messagebox.showinfo('Info agendamento', f"{'Serviço reagendado' if count == 1 else 'Serviços reagendados'} para {new_date_br} a partir das {new_time.strftime('%H:%M')}")
            
        except ValueError:
            messagebox.showwarning("Aviso", "Data ou horário inválido!")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao reagendar:\n{str(e)}")

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
                    ts.slot_date, 
                    ts.start_time, 
                    ts.end_time,
                    s.code, 
                    s.specification, 
                    s.description,
                    s.execution_time_minutes,
                    s.procedure_name,
                    s.procedure_url,
                    ss.notes, 
                    ss.status,
                    ss.priority, 
                    ss.id as schedule_id

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
                    
                    
                    procedure_name = rs.Fields('procedure_name').Value if len(str(rs.Fields('procedure_name').Value)) != None else '(vazio)'

                    order_data = {
                        'start_time': rs.Fields('start_time').Value,
                        'end_time': rs.Fields('end_time').Value,
                        'code': rs.Fields('code').Value,
                        'specification': rs.Fields('specification').Value,
                        'description': rs.Fields('description').Value,
                        'execution_time_minutes': rs.Fields('execution_time_minutes').Value,
                        'procedure_name': procedure_name,
                        'procedure_url': rs.Fields('procedure_url').Value,
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
        
        meses_pt = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho','Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
        
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

        nav_row2 = tk.Frame(nav_frame, bg=self.cor_fundo)
        nav_row2.pack(fill='x')
        
        lab_label = tk.Label(nav_row2, text=f"Calendário de Serviços",font=('Segoe UI', 12, 'bold'), bg=self.cor_fundo, fg='white',anchor='w')
        #lab_label = tk.Label(nav_row2, text=f"Calendário de Serviços - Laboratório de {self.lab_name}",font=('Segoe UI', 12, 'bold'), bg=self.cor_fundo, fg='white',anchor='w')
        lab_label.pack(side='left', padx=5)

        btn_prev = tk.Label(nav_row, text="◀", anchor='center', justify='center',font=('Segoe UI', 12, 'bold'), bg=self.cor_fundo, fg='white',cursor='hand2', padx=0)
        btn_prev.pack(side='left', padx=5)
        btn_prev.bind('<Button-1>', lambda e: self.change_month(-1))

        month_label = tk.Label(nav_row, text=f"{meses_pt[month-1]} {year}",font=('Segoe UI', 12, 'bold'), bg=self.cor_fundo, fg='white')
        month_label.pack(side='left', padx=5)

        btn_next = tk.Label(nav_row, text="▶", anchor='center', justify='center',font=('Segoe UI', 12, 'bold'), bg=self.cor_fundo, fg='white',cursor='hand2', padx=0)
        btn_next.pack(side='left', padx=5)
        btn_next.bind('<Button-1>', lambda e: self.change_month(1))



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
        
        if bool_agenda == 1:
            self.buscar = ttk.Button(grid_frame, text=" Limpar agenda ", command=self.clear_all_scheduled, cursor='hand2')
            self.buscar.pack(side="left", padx=5)

    def show_day_context_menu(self, event, calendar_key):
        context_menu = tk.Menu(self.calendar_frame, tearoff=0)
        
        count = len(self.calendar_data.get(calendar_key, []))
        
        if count == 1:
            context_menu.add_command(label=f"Ver serviço", command=lambda k=calendar_key: self.show_day_orders(k))
            context_menu.add_separator()
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
        
        if count == 1:
            
            context_menu.add_separator()
            context_menu.add_command(label=f"Reagendar o serviço para um dia específico...", command=lambda k=calendar_key: self.reschedule_day_services_manual(k))
            context_menu.add_command(label=f"Reagendar o serviço automaticamente", command=lambda k=calendar_key: self.reschedule_day_services(k))
        elif count > 1:
            
            context_menu.add_separator()
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
            rs.Close()
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
            
            self.show_loading_screen(f"Agendamento de {len(schedule_ids)} {'itens' if len(schedule_ids) > 1 else 'item'} em Laboratório de {self.lab_name}")

            self.next_start, self.next_end = self.schedule_all_pending(from_date=reschedule_date + timedelta(days=1))
            
            self.refresh_calendar()
            
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
        
        tk.Label(popup, text=f"Reagendar {count} {'serviço' if count == 1 else 'serviços'} de {date_str}",font=('Segoe UI', 10, 'bold'), bg=self.cor_fundo, fg='white').pack(pady=(15, 10))
        
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
        
        tk.Label(popup, text="Horário de início (HH:MM):", font=('Segoe UI', 9),bg=self.cor_fundo, fg='white').pack(anchor='w', padx=30, pady=(10, 0))
        
        time_var = tk.StringVar(value="08:00")
        ttk.Entry(popup, textvariable=time_var, width=10).pack(anchor='w', padx=30, pady=(2, 5))
        
        tk.Label(popup, text="⚠ Os serviços serão agendados em sequência\n   a partir desta data e horário.",font=('Segoe UI', 8), bg=self.cor_fundo, fg='#FFD700').pack(pady=(10, 5))
        
        btn_frame = tk.Frame(popup, bg=self.cor_fundo)
        btn_frame.pack(pady=15)
        ttk.Button(btn_frame, text=" Confirmar Reagendamento ",command=lambda: self._execute_manual_reschedule(calendar_key, popup, day_var, month_var, year_var, time_var)).pack(side='left', padx=5)
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
            
            self.show_loading_screen(f"Agendando {len(schedule_ids)} {'itens' if len(schedule_ids) > 1 else 'item'} em Laboratório de {self.lab_name}")

            self.next_start, self.next_end = self.schedule_all_pending(from_date=new_datetime)
            
            self.refresh_calendar()
            self.hide_loading_screen()
            
            new_date_br = new_date.strftime('%d/%m/%Y')
            messagebox.showinfo('Info agendamento', f"{'Serviços reagendados' if len(schedule_ids) > 1 else 'Serviço reagendado'} para {new_date_br} a partir das {new_time.strftime('%H:%M')}")
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
        popup.title(f"Serviços - {date_str} (Laboratório de {self.lab_name})")
        
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
                    f"{start_str} - {end_str}", 
                    os_code, 
                    code,
                    order['specification'], 
                    order['description'],
                    f"{order['execution_time_minutes']} min", 
                    order['code']
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
            
            self.refresh_calendar()
            
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

                tree.insert("", 
                            "end", 
                            values=(f"{start_str} - {end_str}", 
                                    os_code, 
                                    code, 
                                    order['specification'], 
                                    order['description'], 
                                    f"{order['execution_time_minutes']} min", 
                                    order['code']), 
                                    tags=(tag,))
        
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
        
                # NEW: View details of selected service
        if selected_count == 1:
            context_menu.add_command(label="Visualizar o serviço selecionado [Em desenvolvimento]",command=lambda t=tree: self.show_service_details(t))
        else:
            context_menu.add_command(label=f"Visualizar os {selected_count} serviços selecionados [Em desenvolvimento]",command=lambda t=tree: self.show_service_details(t))
        
        context_menu.add_separator()
        context_menu.add_command(label=f"Iniciar chat do Teams com {self.lab_name}", command=self.open_teams_chat)        

        if os_code:
            sharepoint_url = self.get_sharepoint_url(os_code)
            if sharepoint_url:
                context_menu.add_command(label="Abrir link do SharePoint", command=lambda u=sharepoint_url: webbrowser.open(u))
                #context_menu.add_command(label="Copiar link do SharePoint", command=lambda u=sharepoint_url: self.copy_to_clipboard(u) if hasattr(self, 'copy_to_clipboard') else None)

        context_menu.add_command(label="Abrir diretório da planilha de cálculo", command=lambda: self.open_teams_link_planilha(None))
        
        if calendar_key and total_count > 0:
            context_menu.add_separator()
            if selected_count == 1:
                context_menu.add_command(label="Reagendar o serviço selecionado automaticamente",                                   command=lambda t=tree, s=selected: self.reschedule_selected_services(t, s))
                context_menu.add_command(label="Reagendar o serviço selecionado para um dia específico...",                         command=lambda t=tree, s=selected: self.reschedule_selected_services_manual(t, s))
            else:
                context_menu.add_command(label=f"Reagendar {selected_count} serviços selecionados automaticamente",                 command=lambda t=tree, s=selected: self.reschedule_selected_services(t, s))
                context_menu.add_command(label=f"Reagendar {selected_count} serviços selecionados para um dia específico...",       command=lambda t=tree, s=selected: self.reschedule_selected_services_manual(t, s))
                
        context_menu.post(event.x_root, event.y_root)

    def show_service_details(self, tree):
        """Show detailed view of selected service(s)"""
        selected = tree.selection()
        if not selected:
            return
        
        detail_window = tk.Toplevel(self.frame_certificados)
        detail_window.title(f"Detalhes {'do serviço' if len(selected)==1 else 'dos serviços'}")
        detail_window.geometry("700x450")
        detail_window.configure(bg=self.cor_fundo)
        #detail_window.attributes('-topmost', True)
        detail_window.transient(self.frame_certificados)
        text_frame = ttk.Frame(detail_window)
        text_frame.pack(pady=10, padx=10, expand=True, fill="both")
        
        text_widget = tk.Text(text_frame, wrap="word", font=("Consolas", 10), bg='white', fg='black')
        scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side="right", fill="y")
        text_widget.pack(side="left", expand=True, fill="both")
        
        output_text = f"{'='*70}\n"
        output_text += f"Total de itens selecionados: {len(selected)}\n"
        #output_text += f"Laboratório de {self.lab_name}\n"
        output_text += f"{'='*70}\n\n"
        
        for i, item_id in enumerate(selected, 1):
            values = tree.item(item_id)['values']
            
            os_code = values[1]
            if os_code:
                sharepoint_url = self.get_sharepoint_url(os_code)
                output_text += f"{sharepoint_url}\n" #SHAREPOINT
                
            output_text += f"{values[1]}\n" #OS

            output_text += f"\n" #CERTIFICADO
            output_text += f"\n" #TAG
            output_text += f"\n" #DATA CALIBRAÇÃO
            
            #output_text += f"Horário:    {values[0]}\n"
            #output_text += f"Item:       {values[2]}\n"
            #output_text += f"Serviço:    {values[3]}\n"
            #output_text += f"Descrição:  {values[4]}\n"
            #output_text += f"Duração:    {values[5]}\n"
            #output_text += f"Código:     {values[6]}\n"
            
            #SHAREPOINT, OS, CERTIFICADO, TAG, DATA CALIBRAÇÃO, CLIENTE, ENDEREÇO, CONTATO



            output_text += f"\n{'-'*40}\n\n"
        
        text_widget.insert("1.0", output_text)
        text_widget.config(state="disabled")
        
        # Close button
        ttk.Button(detail_window, text="Fechar", command=detail_window.destroy).pack(pady=(0, 10))

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

            self.show_loading_screen(f"Agendando {len(selected_indices)} {'itens' if len(selected_indices) > 1 else 'item'} em Laboratório de {self.lab_name}")
            first_start, last_end = self.schedule_all_pending(from_date=reschedule_date + timedelta(days=1))
            
            self.refresh_calendar()

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
        
        confirm = messagebox.askyesno("Info agendamento", f"Deseja remover {'o bloqueio' if count > 1 and count < 3 else f'os {count-1} bloqueios'} do dia {date_str}?")
        if not confirm:
            return
        
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            conn.Open(self.str_conn)
            date_str_sql = exc_date.strftime('%Y-%m-%d')
            conn.Execute(f"DELETE FROM [IST-PGE].dbo.Calendar_Exceptions WHERE exception_date = '{date_str_sql}' AND id_sector = {self.id_sector} AND (notes IS NULL OR notes NOT LIKE '%Intervalinho%')")
            conn.Close()
            
            self.refresh_calendar()

            self.status_label.config(text=f"{'Bloqueios removidos' if count > 2 and count < 4 else 'bloqueio removido'} de {date_str}")

            messagebox.showinfo("Info agendamento", f"{'Bloqueios removidos' if count > 2 and count < 4 else 'Bloqueio removido'} de {date_str}")
        except Exception as e:
            messagebox.showerror("Info agendamento", f"Falha ao remover bloqueios:\n{str(e)}")

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
            
            rs.Open("""SELECT 
                    id, 
                    code, 
                    specification, 
                    execution_time_minutes, 
                    description,
                    procedure_name,
                    procedure_url 

                    FROM [IST-PGE].dbo.Service_Modes_Local 

                    ORDER BY code ASC""", conn)
            
            self.services_data = []
            if not rs.EOF:
                rs.MoveFirst()
                while not rs.EOF:
                    
                    procedure_name = rs.Fields('procedure_name').Value if len(str(rs.Fields('procedure_name').Value)) != None else '(vazio)'

                    self.services_data.append({
                        'id': rs.Fields('id').Value,
                        'code': rs.Fields('code').Value,
                        'specification': rs.Fields('specification').Value,
                        'description': rs.Fields('description').Value,
                        'execution_time_minutes': rs.Fields('execution_time_minutes').Value,
                        'procedure_name': procedure_name,
                        'procedure_url': rs.Fields('procedure_url').Value

                    })
                    rs.MoveNext()
            rs.Close()
            conn.Close()
        except Exception as e:
            self.status_label.config(text=f"Erro ao carregar serviços: {str(e)}")

    def clear_all_scheduled(self):
        
        confirm = messagebox.askyesno("Info agendamento", f"Tem certeza de que deseja limpar a agenda de serviços de Laboratório de {self.lab_name}?")

        if not confirm:
            return
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            conn.Open(self.str_conn)
            conn.Execute(f"DELETE FROM [IST-PGE].dbo.Time_Slots WHERE id_sector = {self.id_sector}")
            conn.Execute(f"DELETE FROM [IST-PGE].dbo.Service_Schedule WHERE id_sector = {self.id_sector}")
            conn.Close()
            
            self.refresh_calendar()
            self.status_label.config(text=f"Os agendamentos de serviços da agenda de {self.lab_name} foram removidos")
            messagebox.showinfo("Sucesso", f"Todos os agendamentos de serviços foram removidos de Laboratório de {self.lab_name}!")

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


class ServiceSchedulerLinkedDirect: ### SELETOR DE SERVIÇOS PARA AGENDAMENTO
    """ ABA: Serviços para Agendamento (Linked Server) - Multi-Laboratório """
    
    def __init__(self, parent_notebook, str_conn, str_conn_primary, cor_fundo, lab_code="", lab_config=None):
        self.str_conn = str_conn
        self.str_conn_primary = str_conn_primary
        self.cor_fundo = cor_fundo
        self.lab_code = lab_code
        self.lab_config = lab_config or {}
        
        lab_name = self.lab_config.get('name', lab_code)
        sector_id = self.lab_config.get('sector_id', 0)
        
        tab_name = f"Serviços de {lab_name}"
        self.frame_certificados = ttk.Frame(parent_notebook)
        parent_notebook.add(self.frame_certificados, text=tab_name)

        parent_notebook.bind('<<NotebookTabChanged>>', self.on_tab_changed)
    
        self.selected_services = []
        self.all_rows = []
        
        self.setup_ui()
        
        try:
            self.load_services()
        except Exception as e:
            self.status_label.config(text=f"Servidor vinculado indisponível: {str(e)[:80]}")
        self.cleanup_pending_services()
        
    def open_procedure_manager(self):
        """Open the procedure management window"""
        ProcedureManager(
            self.frame_certificados,
            self.str_conn,
            self.str_conn_primary,
            self.cor_fundo,
            lab_code=self.lab_code,
            lab_config=self.lab_config
        )

    def on_tab_changed(self, event):
        notebook = event.widget
        current_tab = notebook.select()
        current_tab_text = notebook.tab(current_tab, "text")
        
        if self.lab_config.get('name', '') in current_tab_text:
            self.load_services()

    def open_teams_chat(self):
        try:
            chat_id = self.lab_config.get('teams_chat_id', '')
            if chat_id:
                link = f"https://teams.cloud.microsoft/l/chat/19:{chat_id}@thread.v2/conversations?context=%7B%22contextType%22%3A%22chat%22%7D"
                webbrowser.open(link)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao abrir chat:\n{str(e)}")

    def setup_ui(self):
        lab_name = self.lab_config.get('name', '')
        
        main_container = tk.Frame(self.frame_certificados, bg=self.cor_fundo)
        main_container.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Top frame
        top_frame = tk.Frame(main_container, bg=self.cor_fundo)
        top_frame.pack(fill='x', pady=(0, 10))
        
        self.greetings = ttk.Label(top_frame, text=f"{buenas}", font=("Segoe UI", 10), background='black')
        self.greetings.pack(side='left', padx=5)
        
        #tk.Label(top_frame, text=f"meme Serviços em Laboratório de {lab_name}",font=('Segoe UI', 12, 'bold'), bg=self.cor_fundo, fg='white').pack(side='left', padx=20)
        tk.Label(top_frame, text=f"Lista de Serviços",font=('Segoe UI', 12, 'bold'), bg=self.cor_fundo, fg='white').pack(side='left', padx=20)
        
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
        
        #columns = ("OS", "Item", "Especificação", "Descrição", "Código", "Duração")
        columns = ("Entrega","OS", "Item", "Tipo","Especificação", "Código", "Duração","Instrução de Trabalho")
        
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings",yscrollcommand=tree_scroll_y.set,xscrollcommand=tree_scroll_x.set,selectmode="extended")
        
        tree_scroll_y.config(command=self.tree.yview)
        tree_scroll_x.config(command=self.tree.xview)
        
        column_widths = {"Entrega":                 40, 
                        "OS":                       80, 
                        "Item":                     120, 
                        "Tipo":                     200, 
                        "Especificação":            200, 
                        "Descrição":                250, 
                        "Código":                   50, 
                        "Duração":                  70, 
                        "Status de Agendamento":    110,
                        "Instrução de Trabalho":    70
                        }
        
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=column_widths.get(col, 80), minwidth=10)
        
        self.tree.pack(expand=True, fill='both')
        
        self.tree.bind('<Button-3>', self.show_context_menu)
        self.tree.bind('<<TreeviewSelect>>', self.on_tree_select)
        self.tree.bind('<Control-a>', self.select_all)
        self.tree.bind('<Control-A>', self.select_all)
        
        button_frame = tk.Frame(main_container, bg=self.cor_fundo)
        button_frame.pack(fill='x', pady=(10, 0))

        self.btn_select_all = ttk.Button(button_frame, text=" Selecionar Todos ", command=self.select_all, cursor='hand2')
        self.btn_select_all.pack(side="left", padx=5)

        if bool_agenda == 1:
            self.btn_schedule = ttk.Button(button_frame, text=" Agendar Selecionados ", command=self.schedule_selected, cursor='hand2')
            self.btn_schedule.pack(side="left", padx=5)
            self.btn_schedule.config(state="disabled")

        btn_refresh = ttk.Button(button_frame, text=" Atualizar Lista ", command=self.load_services, cursor='hand2')
        btn_refresh.pack(side="left", padx=5)

        self.btn_clear = ttk.Button(button_frame, text=" Limpar Seleção ", command=self.clear_selection, cursor='hand2')
        self.btn_clear.pack(side="left", padx=5)
        
        if bool_agenda == 1:
            self.btn_exception = ttk.Button(button_frame, text=" Adicionar Bloqueio de Agenda ", command=self.add_calendar_exception, cursor='hand2')
            self.btn_exception.pack(side="left", padx=5)
            
            
        self.btn_procedures = ttk.Button(button_frame, text=" Lista de Tempos, Serviços e ITs ", command=self.open_procedure_manager, cursor='hand2')
        self.btn_procedures.pack(side="left", padx=5)

        #self.status_label = ttk.Label(self.frame_certificados,text="Selecione os serviços e use o botão direito para agendar",font=("Segoe UI", 12))
        #self.status_label.pack(pady=5)

        #self.btn_export = ttk.Button(button_frame, text="Exportar Relatório de Programação de Serviços ", command=self.export_to_pdf, cursor='hand2')
        #self.btn_export.pack(side="left", padx=5)

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
        self.loading_popup.geometry("550x130")
        
        cor_fundo2 = _from_rgb((0,0,0))
        self.loading_popup.configure(bg=cor_fundo2)
        self.loading_popup.title("")
        self.loading_popup.overrideredirect(True)
        self.loading_popup.attributes('-topmost', True)
        
        # Center on parent
        self.loading_popup.update_idletasks()
        x = self.frame_certificados.winfo_rootx() + (self.frame_certificados.winfo_width() // 2) - 175
        y = self.frame_certificados.winfo_rooty() + (self.frame_certificados.winfo_height() // 2) - 65
        self.loading_popup.geometry(f"+{x}+{y}")
        
        frame = tk.Frame(self.loading_popup, bg=cor_fundo2, highlightbackground="#FFFFFF", highlightthickness=2)
        frame.pack(fill='both', expand=True, padx=2, pady=2)

        if bool_agenda == 1:
            tk.Label(frame, text=f"Oioi, {username}! O Assistente de Agenda está carregando:", font=('Segoe UI', 10), background='black', fg='white').pack(pady=(15, 5))
        elif bool_agenda == 0:
            tk.Label(frame, text=f"Oioi, {username}! O Assistente de Dados está carregando:", font=('Segoe UI', 10), background='black', fg='white').pack(pady=(15, 5))

        #tk.Label(frame, text=f"Laboratório de {self.lab_}", font=('Segoe UI', 13, 'bold'), bg=self.cor_fundo, fg="#FFFFFF").pack(pady=(5, 5))
        
        # Message
        tk.Label(frame, text=message, font=('Segoe UI', 11),bg='white', fg='black').pack(pady=(5, 5))
        
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
            # 1. Query linked server
            conn = win32com.client.Dispatch("ADODB.Connection")
            rs = win32com.client.Dispatch("ADODB.Recordset")
            conn.Open(self.str_conn)

            sql = f"""
                SELECT 
                    CONVERT(VARCHAR(10),os.date_finished, 103) as 'Entrega',
                    os.code AS 'OS', i.code AS 'Item',
                    it.name as 'Tipo', sm.specification AS 'Especificação',
                    sm.description AS 'Descrição', sm.code AS 'Code',
                    sm.execution_time AS 'execution_time', sm.id AS 'service_mode_id'
                FROM instruments_services iss
                LEFT JOIN orders_services AS os ON os.id = iss.id_service_order 
                LEFT JOIN service_modes AS sm ON sm.id = iss.id_service
                LEFT JOIN instruments AS i ON i.id = iss.id_instrument AND i.id_service_order = os.id
                LEFT JOIN budgets AS b ON b.id_order_service = os.id
                LEFT JOIN instrument_types as it on it.id = i.id_instrument_type
                WHERE os.removed = 0 AND iss.removed = 0 AND sm.removed = 0
                AND i.removed = 0 AND b.is_last_revision = 1
                AND (i.id_current_sector = {sector_id} or i.id_current_sector = 1)
                AND sm.code LIKE '%631{lab_code.lower()}%'
                ORDER BY os.receiving_date ASC
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
                    os_code_num = f"{os_parts[0].lstrip('0')}/{os_year}"

                    row_data = {
                        'Entrega': rs.Fields('Entrega').Value if rs.Fields('Entrega').Value else "",
                        'OS': os_code_num,
                        'Item': rs.Fields('Item').Value if rs.Fields('Item').Value else "",
                        'Tipo': rs.Fields('Tipo').Value if rs.Fields('Tipo').Value else "",
                        'Especificação': rs.Fields('Especificação').Value if rs.Fields('Especificação').Value else "",
                        'Descrição': rs.Fields('Descrição').Value if rs.Fields('Descrição').Value else "",
                        'Code': rs.Fields('Code').Value if rs.Fields('Code').Value else "",
                        'execution_time_raw': str(execution_time_raw) if execution_time_raw else "00:00",
                        'duration_minutes': duration_minutes,
                        'service_mode_id': rs.Fields('service_mode_id').Value,
                    }
                    self.all_rows.append(row_data)
                    rs.MoveNext()
            
            rs.Close()
            conn.Close()
            
            # 2. Query local IST-PGE for scheduled status
            scheduled_map = self._get_scheduled_status(sector_id)
            
            # 3. Query local IST-PGE for procedure info
            procedure_map = self._get_procedure_map()
            
            # 4. Update treeview columns
            self._update_treeview_columns()
            
            # 5. Insert rows with status and procedure
            for row_data in self.all_rows:
                item_code = row_data['Item']
                service_code = row_data['Code']
                status = scheduled_map.get(item_code, "Pendente")
                
                proc = procedure_map.get(service_code, {})
                proc_display = proc.get('name', '')[:30] if proc.get('name') else ""
                
                self.tree.insert("", "end", values=(
                    row_data['Entrega'],
                    row_data['OS'], 
                    row_data['Item'], 
                    row_data['Tipo'],
                    row_data['Especificação'],
                    row_data['Code'], 
                    f"{row_data['duration_minutes']} min",
                    status,
                    proc_display
                ))
            
            count = len(self.all_rows)
            
        except Exception as e:
            self.status_label.config(text=f"Erro: {str(e)}")

    def _get_scheduled_status(self, sector_id):
        """Get mapping of item_code → status from IST-PGE"""
        status_map = {}
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            rs = win32com.client.Dispatch("ADODB.Recordset")
            conn.Open(self.str_conn_primary)
            
            sql = f"""

                SELECT 
                ss.notes, 
                ss.status, 
                ss.scheduled_start, 
                ss.scheduled_end

                FROM [IST-PGE].dbo.Service_Schedule ss

                WHERE ss.id_sector = {sector_id}
                AND ss.status IN ('SCHEDULED', 'PENDING')

            """
            rs.Open(sql, conn)
            
            if not rs.EOF:
                rs.MoveFirst()
                while not rs.EOF:
                    notes = rs.Fields('notes').Value if rs.Fields('notes').Value else ""
                    status = rs.Fields('status').Value if rs.Fields('status').Value else ""
                    scheduled_start = rs.Fields('scheduled_start').Value
                    
                    # Extract item code from notes (format: "OS X - ITEM - Spec")
                    parts = notes.split(' - ')
                    if len(parts) >= 2:
                        item_code = parts[1].strip()
                        
                        if scheduled_start:
                            if isinstance(scheduled_start, str):
                                full_str = scheduled_start[:16]  # '2026-05-22 08:00'
                                formatted = datetime.strptime(full_str, '%Y-%m-%d %H:%M').strftime('%d/%m/%Y - %H:%M')
                            elif isinstance(scheduled_start, datetime):
                                formatted = scheduled_start.strftime('%d/%m/%Y - %H:%M')
                            else:
                                formatted = str(scheduled_start)[:10]
                            status_map[item_code] = datetime.strptime(full_str, '%Y-%m-%d %H:%M').strftime('%d/%m/%Y - %H:%M')
                        else:
                            formatted = "None"
                            #full_str = scheduled_start[:16]  # '2026-05-22 08:00'
                            #formatted = datetime.strptime(full_str, '%Y-%m-%d %H:%M').strftime('%d/%m/%Y - %H:%M')

                    rs.MoveNext()
            
            rs.Close()
            conn.Close()
            
        except Exception as e:
            print(f"Error getting scheduled status: {e}")
        
        return status_map

    def _get_procedure_map(self):
        """Get mapping of service_code → procedure info from IST-PGE"""
        procedure_map = {}
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            rs = win32com.client.Dispatch("ADODB.Recordset")
            conn.Open(self.str_conn_primary)
            
            rs.Open("""
                    
                SELECT 
                code, 
                procedure_name, 
                procedure_url 
                    
                FROM [IST-PGE].dbo.Service_Modes_Local
                    
            """, conn)

            if not rs.EOF:
                rs.MoveFirst()
                while not rs.EOF:
                    code = rs.Fields('code').Value
                    name = rs.Fields('procedure_name').Value
                    url = rs.Fields('procedure_url').Value
                    
                    # Show "(vazio)" if name is None or empty
                    if not name or str(name).strip() == "":
                        display_name = "[Vazio]"
                    else:
                        display_name = str(name)
                    

                    procedure_map[code] = {
                        'name': display_name,
                        'url': str(url) if url else ""
                    }
                    rs.MoveNext()
            
            rs.Close()
            conn.Close()
            
        except Exception as e:
            print(f"Error getting procedure map: {e}")
        
        return procedure_map
    
    def open_procedure_url(self):
        """Open the procedure document for the selected service"""
        selected = self.tree.selection()
        if not selected:
            return
        
        all_items = self.tree.get_children()
        idx = all_items.index(selected[0])
        
        if idx < len(self.all_rows):
            service_code = self.all_rows[idx]['Code']
            procedure_map = self._get_procedure_map()
            proc = procedure_map.get(service_code, {})
            url = proc.get('url', '')
            
            if url:
                webbrowser.open(url)
            else:
                service_name = self.all_rows[idx].get('Especificação', service_code)
                messagebox.showinfo("Aviso", f"Nenhum procedimento cadastrado para:\n{service_code} - {service_name}")
                
    def _update_treeview_columns(self):
        """Update treeview to include Status column"""
        # Reconfigure columns
        #self.tree["columns"] = ("Entrega","OS", "Item", "Tipo", "Especificação", "Descrição", "Código", "Duração", "Status de Agendamento")
        self.tree["columns"] = ("Entrega","OS", "Item", "Tipo", "Especificação", "Código", "Duração", "Status de Agendamento", "Instrução de Trabalho")
        
        self.tree.heading("Entrega", text="Entrega")
        self.tree.heading("OS", text="OS")
        self.tree.heading("Item", text="Item")
        self.tree.heading("Tipo", text="Tipo")
        self.tree.heading("Especificação", text="Especificação")
        #self.tree.heading("Descrição", text="Descrição")
        self.tree.heading("Código", text="Código")
        self.tree.heading("Duração", text="Duração")
        self.tree.heading("Status de Agendamento", text="Status de Agendamento")
        self.tree.heading("Instrução de Trabalho", text="Instrução de Trabalho")
        
        self.tree.column("Entrega", width=20)
        self.tree.column("OS", width=20)
        self.tree.column("Item", width=70)
        self.tree.column("Tipo", width=70)
        self.tree.column("Especificação", width=160)
        #self.tree.column("Descrição", width=200)
        self.tree.column("Código", width=30)
        self.tree.column("Duração", width=20)
        self.tree.column("Status de Agendamento", width=110)
        self.tree.column("Instrução de Trabalho", width=70)
        
        ## Color tags for status
        #self.tree.tag_configure('Disponível', background="#E0EC30")  # Green
        #self.tree.tag_configure('PENDING', background='#FFF3CD')    # Yellow


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
               search_term in str(row['Status']).lower() or \
               search_term in str(row['Code']).lower():
                
                self.tree.insert("", "end", values=(
                    row['OS'], 
                    row['Item'], 
                    row['Especificação'],
                    row['Descrição'], 
                    row['Code'], 
                    f"{row['duration_minutes']} min",
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
        lab_name = self.lab_config.get('name', '')

        if bool_agenda == 1:
            if count == 1:  
                context_menu.add_command(label=f"Agendar serviço automaticamente", command=self.schedule_selected)
            elif count > 1:
                context_menu.add_command(label=f"Agendar serviços automaticamente", command=self.schedule_selected)
            context_menu.add_separator()
        
        context_menu.add_command(label=f"Iniciar chat do Teams com Laboratório de {lab_name}", command=self.open_teams_chat)
        context_menu.add_command(label="Abrir Instrução de Trabalho no EmDoc", command=self.open_procedure_url)
        context_menu.add_command(label="Editar link para Instrução de Trabalho no EmDoc", command=self.open_procedure_url)
        context_menu.add_separator()
        context_menu.add_command(label="Selecionar todos", command=self.select_all)
        context_menu.add_command(label="Limpar seleção", command=self.clear_selection)
        
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
                item_code = row['Item']
                
                service_mode_id = row.get('service_mode_id', None)

                # Sync to Service_Modes_Local
                sync_sql = f"""
                        IF NOT EXISTS (SELECT 1 FROM [IST-PGE].dbo.Service_Modes_Local WHERE id = {service_mode_id})
                        BEGIN
                            INSERT INTO [IST-PGE].dbo.Service_Modes_Local (id, code, specification, description, execution_time_minutes)
                            VALUES ({service_mode_id}, '{row['Code']}', '{str(row['Especificação']).replace("'", "''")}', 
                                    '{str(row['Descrição']).replace("'", "''")}', {duration});
                        END
                        ELSE
                        BEGIN
                            UPDATE [IST-PGE].dbo.Service_Modes_Local
                            SET execution_time_minutes = {duration},
                                specification = '{str(row['Especificação']).replace("'", "''")}',
                                description = '{str(row['Descrição']).replace("'", "''")}'
                            WHERE id = {service_mode_id};
                        END
                    """
                conn.Execute(sync_sql)
                
                rs = win32com.client.Dispatch("ADODB.Recordset")
                rs.Open(f"SELECT id FROM [IST-PGE].dbo.Service_Modes_Local WHERE id = {service_mode_id}", conn)
                local_service_id = rs.Fields('id').Value if not rs.EOF else None
                rs.Close()
                
                if local_service_id:
                    
                    conn.Execute(f"""
                        DELETE FROM [IST-PGE].dbo.Time_Slots 
                        WHERE schedule_id IN (
                            SELECT id FROM [IST-PGE].dbo.Service_Schedule 
                            WHERE notes LIKE '% - {item_code} - %' AND id_sector = {sector_id}
                        )""")

                    conn.Execute(f"""
                        DELETE FROM [IST-PGE].dbo.Service_Schedule 
                        WHERE notes LIKE '% - {item_code} - %' AND id_sector = {sector_id}
                    """)
                    
                    # Insert new schedule
                    notes = f"OS {row['OS']} - {row['Item']} - {row['Especificação']}"
                    schedule_sql = f"""
                        INSERT INTO [IST-PGE].dbo.Service_Schedule (service_id, notes, id_sector)
                        VALUES ({local_service_id}, '{notes.replace("'", "''")}', {sector_id})
                    """
                    conn.Execute(schedule_sql)
                    added_count += 1
            
            conn.Close()
            
            #if added_count == 1:
            #    self.status_label.config(text=f"{added_count} serviço adicionado à fila!")
            #elif added_count == 0:
            #    self.status_label.config(text=f"Nenhum serviço adicionado à fila!")
            #else:
            #    self.status_label.config(text=f"{added_count} serviços adicionados à fila!")
                
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
                SELECT 
                ss.id as schedule_id, 
                ss.service_id,
                s.execution_time_minutes, 
                ss.priority, 
                ss.requested_at, 
                ss.notes

                FROM [IST-PGE].dbo.Service_Schedule ss

                JOIN [IST-PGE].dbo.Service_Modes_Local s ON ss.service_id = s.id
                
                WHERE ss.status = 'PENDING' AND ss.id_sector = {sector_id}
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
            
            self.show_loading_screen(f"Agendamento de {pending_count} {'serviços' if pending_count > 1 else 'serviço' } em Laboratório de {lab_name}...")
            self.frame_certificados.update()

            current_time = datetime.now()
            scheduled_count = 0
            for service in pending_services:
                next_start, next_end = self.find_next_available_slot(current_time, service['execution_time_minutes'])
                

                if next_start is None:
                    continue
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
                            {service['service_id']}, {sector_id})
                """)

                current_time = next_end
                scheduled_count += 1
            
            self.hide_loading_screen()
            conn.Close()
            
            if scheduled_count == 1:
                self.load_services()
                messagebox.showinfo("Info agendamento", f"{scheduled_count} serviço agendado em Laboratório de {lab_name}!")                
            else:
                self.load_services()
                messagebox.showinfo("Info agendamento", f"{scheduled_count} serviços agendados em Laboratório de {lab_name}!")
              
        except Exception as e:
            self.hide_loading_screen()
            messagebox.showerror("Erro", f"Falha ao processar fila:\n{str(e)}")
        
    # ===== FIFO ENGINE (same as ServiceScheduler) =====
    
    def find_next_available_slot(self, from_time, duration_minutes):
        check_date = from_time.date()
        max_days = 365
        max_iterations_per_window = 100  # Safety limit
        
        availability = self._get_availability()
        exceptions = self._get_exceptions()
        
        for day_offset in range(max_days):
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
                
                iterations = 0
                while proposed_start + timedelta(minutes=duration_minutes) <= window_end_dt:
                    iterations += 1
                    if iterations > max_iterations_per_window:
                        break  # Safety break
                    
                    proposed_end = proposed_start + timedelta(minutes=duration_minutes)
                    
                    # Check time-specific exceptions
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
                    
                    # Check conflicts
                    if self._has_time_slot_conflict(check_date, proposed_start.time(), proposed_end.time()):
                        next_free = self._get_next_free_time(check_date, proposed_start.time())
                        if next_free is None:
                            break  # No free time found, exit this window
                        new_start = datetime.combine(check_date, next_free)
                        if new_start <= proposed_start:  # No progress
                            break  # Prevent infinite loop
                        proposed_start = new_start
                        continue
                    
                    return proposed_start, proposed_end
                
                # Break out of while without finding slot — try next window
            
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
    
    def add_calendar_exception(self):
        popup = tk.Toplevel(self.frame_certificados)
        popup.geometry("400x400")
        popup.minsize(400, 400)
        popup.maxsize(400, 400)
        popup.configure(bg='black')
        
        lab_name = self.lab_config.get('name', '')
        popup.title(f"Adicionar Bloqueio de Agenda - Laboratório de {lab_name}")
        
        popup.transient(self.frame_certificados)
        
        tk.Label(popup, text=f"Adicionar Bloqueio de Agenda",font=('Segoe UI', 13, 'bold'), bg='black', fg='white').pack(anchor='w',padx=15, pady=(15, 0))
        
        tk.Label(popup, text="Data:", font=('Segoe UI', 9), bg='black', fg='white').pack(anchor='w', padx=20, pady=(15, 0),)
        
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
            #self.status_label.config(text=f"Bloqueio de calendário adicionado em: {exception_date_br}")
        except ValueError:
            messagebox.showwarning("Aviso", "Data inválida!")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao salvar Bloqueio de Calendário:\n{str(e)}")


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
                context_menu.add_separator()
                context_menu.add_command(label=f"Iniciar chat do Teams com {setor_teams}",command=lambda i=item: self.open_teams_chat(i))
                context_menu.add_separator()
                context_menu.add_command(label="Copiar link", command=lambda u=sharepoint_url: self.copy_to_clipboard(u))
                context_menu.add_command(label="Copiar linha", command=lambda i=item: self.copy_row_data(i))
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
            elif status == "Esta semana":
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
        print_window.transient(self.frame_certificados)
        
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


class DatabaseViewerOrcamentos:  ### ORÇAMENTOS POR CLIENTE
    
    def __init__(self, parent_notebook, str_conn, cor_fundo):
        self.str_conn = str_conn
        self.cor_fundo = cor_fundo
        self.selected_items = []
        self.frame_certificados = ttk.Frame(parent_notebook)
        parent_notebook.add(self.frame_certificados, text=" Orçamentos ")
        
        self.setup_ui()
    
    def on_entry_focus_in(self, event):
        if self.search_entry.get() == self.placeholder_text:
            self.search_entry.delete(0, 'end')
            self.search_entry.configure(style='TEntry')

    def on_entry_focus_out(self, event):
        if not self.search_entry.get():
            self.search_entry.insert(0, self.placeholder_text)
            style = ttk.Style()
            self.search_entry.configure(style='Placeholder.TEntry')
        
    def setup_ui(self):
        search_frame = ttk.Frame(self.frame_certificados)
        search_frame.pack(pady=10, padx=10, fill="x")
        
        self.greetings = ttk.Label(search_frame, text=f"{buenas}", font=("Segoe UI", 10), background='black')
        self.greetings.pack(side="left", padx=5)
        
        # Search by customer name
        tk.Label(search_frame, text="Cliente:", font=('Segoe UI', 9), 
                bg=self.cor_fundo, fg='white').pack(side="left", padx=(10, 5))
        
        self.search_entry = ttk.Entry(search_frame, width=40)
        self.search_entry.pack(side="left", padx=5)
        self.search_entry.bind('<Return>', lambda e: self.perform_search())

        style = ttk.Style()
        style.configure('Placeholder.TEntry', foreground='#7D7D7D')

        self.placeholder_text = "Digite o nome do cliente..."
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
            "Cliente",
            "Instrumento",
            "Código",
            "Modelo",
            "Nº Série",
            "Serviço",
            "Aprovação",
            "Orçamento",
            "Valor (R$)",
            #"Desconto (%)"
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
            "Instrumento": 120,
            "Código": 80,
            "Modelo": 100,
            "Nº Série": 80,
            "Serviço": 180,
            "Aprovação": 80,
            "Orçamento": 80,
            "Valor (R$)": 80,
            #"Desconto (%)": 80
        }
        
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=column_widths.get(col, 80), minwidth=50)
        
        self.tree.pack(expand=True, fill="both")
        
        self.tree.bind('<Button-3>', self.show_context_menu)
        self.tree.bind('<<TreeviewSelect>>', self.on_select)
        
        button_frame = ttk.Frame(self.frame_certificados)
        button_frame.pack(pady=10, padx=10, fill="x")
        self.btn_visualizar = ttk.Button(button_frame, text="Visualizar Selecionado(s)", command=self.print_selected)
        self.btn_visualizar.pack(side="left", padx=5)
        self.btn_visualizar.config(state="disabled")

        self.btn_limpar = ttk.Button(button_frame, text="Limpar Seleção", command=self.clear_selection)
        self.btn_limpar.pack(side="left", padx=5)
        self.btn_limpar.config(state="disabled")

        self.status_label = ttk.Label(self.frame_certificados, text="Busque por um cliente para visualizar orçamentos", font=("Segoe UI", 9))
        self.status_label.pack(pady=5)
    
    def show_context_menu(self, event):
        item = self.tree.identify_row(event.y)
        
        if item:
            self.tree.selection_set(item)
            
            context_menu = tk.Menu(self.tree, tearoff=0)
            
            context_menu.add_command(label="Visualizar Selecionado", command=self.print_selected)
            context_menu.add_command(label="Copiar linha", command=lambda i=item: self.copy_row_data(i))
            
            context_menu.post(event.x_root, event.y_root)

    def copy_to_clipboard(self, text):
        try:
            self.frame_certificados.clipboard_clear()
            self.frame_certificados.clipboard_append(text)
            self.status_label.config(text="Copiado!")
        except Exception:
            pass

    def copy_row_data(self, item):
        try:
            values = self.tree.item(item)['values']
            row_text = "\n".join([f"{val}" for val in values])
            self.frame_certificados.clipboard_clear()
            self.frame_certificados.clipboard_append(row_text)
            self.status_label.config(text="Linha copiada!")
            messagebox.showinfo("Sucesso", "Dados da linha copiados!")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao copiar:\n{str(e)}")
    
    def execute_query(self, where_clause=None):
        try:
            conn = win32com.client.Dispatch("ADODB.Connection")
            rs = win32com.client.Dispatch("ADODB.Recordset")
            
            conn.Open(self.str_conn)
            
            sql = """
                SELECT 

                    c.name AS 'Cliente',
                    it.name AS 'Instrumento',
                    i.code AS 'Código',
                    im.name AS 'Modelo',
                    i.serial_number AS 'Nº Série',
                    sm.specification AS 'Serviço',
                    CONVERT(VARCHAR(10), iss.approve_date, 103) AS 'Aprovação',
                    b.code AS 'Orçamento',
                    ROUND((bi.unitary_value - CASE WHEN b.id_discount_type = 1 THEN (bi.unitary_value * bi.discount / 100.0) ELSE 0 END), 2) AS 'Valor'
                    --((((CASE WHEN b.id_discount_type = 1 THEN (bi.unitary_value * bi.discount / 100.0) ELSE 0 END)/bi.unitary_value)*100), 2) AS 'Desconto'

                FROM instruments i
                LEFT JOIN instruments_models im ON im.id = i.id_instrument_model
                LEFT JOIN instrument_types it ON it.id = i.id_instrument_type
                LEFT JOIN instruments_services iss ON iss.id_instrument = i.id
                LEFT JOIN service_modes sm ON sm.id = iss.id_service
                LEFT JOIN budgets_items_details bid ON bid.id = iss.id_budget_item_detail
                LEFT JOIN budgets_items bi ON bi.id = bid.id_budget_item
                LEFT JOIN budgets b ON b.id = bi.id_budget
                LEFT JOIN customers c on c.id = b.id_customer
                WHERE i.removed = 0
            """
            
            if where_clause:
                sql += f" AND {where_clause}"
            
            sql += " ORDER BY it.name ASC"
            
            rs.Open(sql, conn)
            
            results = []
            if not rs.EOF:
                rs.MoveFirst()
                while not rs.EOF:
                    row = []
                    for i in range(rs.Fields.Count):
                        value = rs.Fields(i).Value
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
        self.status_label.config(text="Carregando orçamentos...", font=("Segoe UI", 12))
        self.frame_certificados.update()
        
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        results = self.execute_query()
        
        for row in results:
            self.tree.insert("", "end", values=row)
            
        if len(results) == 0:
            self.status_label.config(text="Nenhum orçamento encontrado")
        elif len(results) == 1:
            self.status_label.config(text=f"1 orçamento encontrado")
        else:
            self.status_label.config(text=f"{len(results)} orçamentos encontrados")
    
    def perform_search(self):
        search_term = self.search_entry.get().strip()
        
        if not search_term or search_term == self.placeholder_text:
            self.load_all_records()
            return
        
        self.status_label.config(text=f"Buscando por '{search_term}'...")
        self.frame_certificados.update()
        
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Search by customer name from the linked server
        where_clause = f"""
            b.id IN (
                SELECT b2.id FROM budgets b2
                LEFT JOIN budgets_items bi2 ON bi2.id_budget = b2.id
                LEFT JOIN budgets_items_details bid2 ON bid2.id_budget_item = bi2.id
                LEFT JOIN instruments_services iss2 ON iss2.id_budget_item_detail = bid2.id
                LEFT JOIN instruments i2 ON i2.id = iss2.id_instrument
                LEFT JOIN customers c2 ON c2.id = i2.id_customer
                WHERE c2.name LIKE '%{search_term}%'
            )
        """
        
        results = self.execute_query(where_clause)
        
        for row in results:
            self.tree.insert("", "end", values=row)
        
        self.status_label.config(text=f"Encontrados {len(results)} itens para '{search_term}'")
    
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
                'cliente':          values[0] if len(values) > 0 else '',
                'instrumento':      values[1] if len(values) > 1 else '',
                'codigo':           values[2] if len(values) > 2 else '',
                'modelo':           values[3] if len(values) > 3 else '',
                'serie':            values[4] if len(values) > 4 else '',
                'servico':          values[5] if len(values) > 5 else '',
                'aprovacao':        values[6] if len(values) > 6 else '',
                'orcamento':        values[7] if len(values) > 7 else '',
                #'valor':            round(values[7],2) if len(values) > 7 else '',
                'valor':            round(values[8],2) if len(values) > 8 else '',
                #'desconto':         values[8] if len(values) > 8 else ''
            })
        
        return selected_data
    
    def print_selected(self):
        selected_data = self.get_selected_data()
        
        if not selected_data:
            messagebox.showinfo("Aviso", "Nenhum item selecionado para visualizar.")
            return
        
        print_window = tk.Toplevel(self.frame_certificados)
        print_window.title("Orçamentos Selecionados")
        print_window.geometry("700x450")
        print_window.configure(bg=self.cor_fundo)
        print_window.attributes('-topmost', True)
        
        text_frame = ttk.Frame(print_window)
        text_frame.pack(pady=10, padx=10, expand=True, fill="both")
        
        text_widget = tk.Text(text_frame, wrap="word", font=("Consolas", 10), bg='white', fg='black')
        scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side="right", fill="y")
        text_widget.pack(side="left", expand=True, fill="both")
        
        output_text = f"{'='*60}\n"
        output_text += f"ORÇAMENTOS SELECIONADOS\n"
        output_text += f"Total: {len(selected_data)}\n"
        output_text += f"{'='*60}\n\n"
        
        for i, item in enumerate(selected_data, 1):
            output_text += f"--- Item {i} ---\n\n"
            output_text += f"Cliente:    {item['cliente']}\n"
            output_text += f"Instrumento:    {item['instrumento']}\n"
            output_text += f"Código:         {item['codigo']}\n"
            output_text += f"Modelo:         {item['modelo']}\n"
            output_text += f"Nº Série:       {item['serie']}\n"
            output_text += f"Serviço:        {item['servico']}\n"
            output_text += f"Aprovação:      {item['aprovacao']}\n"
            output_text += f"Orçamento:      {item['orcamento']}\n"
            output_text += f"Valor:          R$ {round(item['valor'],2)}\n"
            #output_text += f"Desconto:       {item['desconto']}%\n"
            output_text += f"\n{'-'*40}\n\n"

        text_widget.insert("1.0", output_text)
        text_widget.config(state="disabled")
        
        ttk.Button(print_window, text="Fechar", command=print_window.destroy).pack(pady=(0, 10))
        
    def clear_selection(self):
        for item in self.tree.selection():
            self.tree.selection_remove(item)
        self.status_label.config(text="Seleção limpa")


# ==================== MAIN ====================
status_servidor, cor_status = verificar_disponibilidade3()
status_servidor3, cor_status3 = status_servidor, cor_status

root = tk.Tk()

if bool_agenda == 1:
    root.title("Assistente de Agenda")
elif bool_agenda == 0:
    root.title("Assistente de Dados")

root.geometry("550x675")
root.minsize(650, 675)
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
LabSelector(notebook, cor_fundo)


db_viewer6 = DatabaseViewer6(notebook, STR_CONN_LINKED, cor_fundo)         ### VISTA GERAL DE ORDENS DE SERVIÇO
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
