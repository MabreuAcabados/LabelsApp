import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import messagebox, simpledialog
import psycopg2
import os, sys, subprocess
import os
import sys
import pandas as pd
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle
import tempfile
import math
from collections import defaultdict
import hashlib
import time
from datetime import datetime, timedelta
from PIL import Image, ImageTk
from typing import Any
import ctypes
import threading

# Control de logs para evitar ruido en consola en producción
DEBUG_LOGS = False

def debug_log(*args: Any, **kwargs: Any):
    if DEBUG_LOGS:
        try:
            print(*args, **kwargs)
        except Exception:
            pass

# Silenciar stdout por defecto para evitar mensajes en consola (mantiene stderr)
if not DEBUG_LOGS:
    try:
        sys.stdout = open(os.devnull, 'w')
    except Exception:
        pass

APP_VERSION = "5.0.0"
URL_VERSION = "https://labelsapp.onrender.com/Version2.txt"
URL_EXE = "https://labelsapp.onrender.com/LabelsApp2.exe"

def obtener_sucursal_usuario(usuario_id):
    """Detecta la sucursal basándose en el ID del usuario"""
    if not usuario_id:
        return 'principal'
    
    usuario_lower = usuario_id.lower()
    
    # Mapeo completo de usuarios a sucursales
    if 'alameda' in usuario_lower:
        return 'alameda'
    elif 'churchill' in usuario_lower:
        return 'churchill'
    elif 'bavaro' in usuario_lower:
        return 'bavaro'
    elif 'bellavista' in usuario_lower:
        return 'bellavista'
    elif 'tiradentes' in usuario_lower:
        return 'tiradentes'
    elif 'la_vega' in usuario_lower or 'vega' in usuario_lower:
        return 'la_vega'
    elif 'luperon' in usuario_lower:
        return 'luperon'
    elif 'puertoplata' in usuario_lower:
        return 'puertoplata'
    elif 'puntacana' in usuario_lower:
        return 'puntacana'
    elif 'romana' in usuario_lower:
        return 'romana'
    elif 'santiago' in usuario_lower:
        return 'santiago1'
    elif 'sanisidro' in usuario_lower:
        return 'sanisidro'
    elif 'villamella' in usuario_lower:
        return 'villamella'
    elif 'terrenas' in usuario_lower:
        return 'terrenas'
    elif 'arroyohondo' in usuario_lower:
        return 'arroyohondo'
    elif 'bani' in usuario_lower:
        return 'bani'
    elif 'rafaelvidal' in usuario_lower:
        return 'rafaelvidal'
    elif 'sanfrancisco' in usuario_lower:
        return 'sanfrancisco'
    elif 'sanmartin' in usuario_lower:
        return 'sanmartin'
    elif 'zonaoriental' in usuario_lower:
        return 'zonaoriental'
    else:
        return 'principal'

def version_tuple(v):
    return tuple(int(x) for x in v.strip().split(".") if x.isdigit())

def is_newer(latest, current):
    return version_tuple(latest) > version_tuple(current)

def run_windows_updater(new_exe_path, current_exe_path):
    # Creamos un .bat temporal que espera a que termine el proceso actual,
    # reemplaza el exe y lanza la nueva versión.
    bat = tempfile.NamedTemporaryFile(delete=False, suffix=".bat", mode="w", encoding="utf-8")
    new_p = new_exe_path.replace("/", "\\")
    cur_p = current_exe_path.replace("/", "\\")
    exe_name = os.path.basename(cur_p)
    bat_contents = f"""@echo off
timeout /t 2 /nobreak > nul
:waitloop
tasklist /FI "IMAGENAME eq {exe_name}" | find /I "{exe_name}" > nul
if %ERRORLEVEL%==0 (
  timeout /t 1 > nul
  goto waitloop
)
move /Y "{new_p}" "{cur_p}"
start "" "{cur_p}"
del "%~f0"
"""
    bat.write(bat_contents)
    bat.close()
    # lanzar el .bat y salir
    subprocess.Popen(["cmd", "/c", bat.name], creationflags=subprocess.CREATE_NEW_CONSOLE)
    sys.exit(0)

def _is_frozen_exe():
    return getattr(sys, "frozen", False)

def _current_binary_path():
    return sys.executable if _is_frozen_exe() else sys.argv[0]

def check_update():
    try:
        is_frozen = _is_frozen_exe()
        headers = {
            "User-Agent": f"LabelsApp/{APP_VERSION}",
            "Cache-Control": "no-cache",
            "Pragma": "no-cache",
        }
        # Importación perezosa
        import requests  # type: ignore

        r = requests.get(URL_VERSION, timeout=5, headers=headers)
        r.raise_for_status()
        latest = r.text.strip()
        if not latest:
            return

        if is_newer(latest, APP_VERSION):
            if os.name == "nt" and not is_frozen:
                # Modo desarrollo: notificar, no actualizar
                try:
                    msg = (
                        f"Hay una nueva versión disponible: {latest}\n\n"
                        f"Versión actual: {APP_VERSION}\n\n"
                        f"Ejecuta el EXE para actualizar automáticamente o descarga la última versión."
                    )
                    ctypes.windll.user32.MessageBoxW(0, msg, "Actualización disponible", 0x40)
                except Exception:
                    pass
                return

            if os.name == "nt" and is_frozen:
                base_path = os.path.dirname(_current_binary_path()) or os.getcwd()
                new_exe = os.path.join(base_path, "LabelsApp_new.exe")
                with requests.get(URL_EXE, stream=True, timeout=20, headers=headers) as resp:
                    resp.raise_for_status()
                    with open(new_exe, "wb") as f:
                        for chunk in resp.iter_content(8192):
                            if chunk:
                                f.write(chunk)
                try:
                    if os.path.getsize(new_exe) < 100_000:
                        return
                except Exception:
                    return
                run_windows_updater(new_exe, _current_binary_path())
                return
            # Otros SO: permanecer en silencio
            return
    except Exception:
        return

if __name__ == "__main__":
    # Verificación en segundo plano si no se pasa --no-update
    if "--no-update" not in sys.argv:
        try:
            threading.Thread(target=check_update, daemon=True).start()
        except Exception:
            pass

# === SISTEMA DE LOGIN INTEGRADO ===
class SistemaLoginIntegrado:
    """Sistema de login integrado para LabelsApp"""
    
    def __init__(self):
        self.db_config = {
            "host": "dpg-d1b18u8dl3ps73e68v1g-a.oregon-postgres.render.com",
            "port": 5432,
            "database": "labels_app_db",
            "user": "admin",
            "password": "KCFjzM4KYzSQx63ArufESIXq03EFXHz3",
            "sslmode": "require"
        }
        self.usuario_actual = None
        self.sucursal = None
        
    def conectar_bd(self):
        """Conecta a la base de datos"""
        try:
            return psycopg2.connect(**self.db_config)
        except Exception as e:
            debug_log(f"❌ Error conectando a BD: {e}")
            return None
    
    def hash_password(self, password):
        """Encripta la contraseña usando SHA-256"""
        return hashlib.sha256(password.encode()).hexdigest()
    
    def verificar_credenciales(self, username, password):
        """Verifica las credenciales del usuario"""
        conn = self.conectar_bd()
        if not conn:
            return {"error": "No se pudo conectar a la base de datos"}
        
        try:
            cur = conn.cursor()
            password_hash = self.hash_password(password)
            
            query = """
            SELECT u.id, u.username, u.password_hash, u.nombre_completo, u.rol, 
                   u.activo, u.sucursal_id, s.nombre as sucursal_nombre
            FROM usuarios u
            LEFT JOIN sucursales s ON u.sucursal_id = s.id
            WHERE u.username = %s AND u.activo = true
            """
            
            cur.execute(query, (username,))
            usuario = cur.fetchone()
            
            if not usuario:
                return {"error": "Usuario no encontrado o inactivo"}
            
            if usuario[2] != password_hash:
                return {"error": "Contraseña incorrecta"}
            
            # Verificar que el rol sea apropiado para LabelsApp
            roles_permitidos = ['facturador', 'administrador']
            if usuario[4] not in roles_permitidos:
                return {"error": f"Rol '{usuario[4]}' no tiene acceso a LabelsApp. Se requiere rol de facturador o administrador."}
            
            return {
                "id": usuario[0],
                "username": usuario[1],
                "nombre_completo": usuario[3],
                "rol": usuario[4],
                "sucursal_id": usuario[6],
                "sucursal_nombre": usuario[7] or "SUCURSAL PRINCIPAL"
            }
            
        except Exception as e:
            return {"error": f"Error en verificación: {e}"}
        finally:
            cur.close()
            conn.close()
    
    def mostrar_login(self):
        """Muestra la ventana de login"""
        ventana_login = tk.Tk()
        ventana_login.title("LabelsApp - Login")
        ventana_login.geometry("450x700")
        ventana_login.resizable(False, False)
        ventana_login.configure(bg="#f5f5f5")
        
        # Configurar icono si existe usando ruta absoluta
        icono_path = obtener_ruta_absoluta("icono.ico")
        if os.path.exists(icono_path):
            try:
                ventana_login.iconbitmap(icono_path)
            except:
                pass
        
        # Centrar ventana
        ventana_login.update_idletasks()
        x = (ventana_login.winfo_screenwidth() // 2) - (450 // 2)
        y = (ventana_login.winfo_screenheight() // 2) - (700 // 2)
        ventana_login.geometry(f"450x700+{x}+{y}")
        
        # Frame principal con sombra
        main_frame = tk.Frame(
            ventana_login, 
            bg="white", 
            relief="flat", 
            bd=0
        )
        main_frame.pack(fill="both", expand=True, padx=30, pady=30)
        
        # Frame de contenido interno
        content_frame = tk.Frame(main_frame, bg="white")
        content_frame.pack(fill="both", expand=True, padx=30, pady=30)
        
        # Logo grande centrado usando ruta absoluta
        logo_path = obtener_ruta_absoluta("logo.png")
        if os.path.exists(logo_path):
            try:
                # Cargar y redimensionar logo
                logo_img = Image.open(logo_path)
                logo_img = logo_img.resize((220, 220), Image.Resampling.LANCZOS)
                self.logo_photo = ImageTk.PhotoImage(logo_img)
                
                # Mostrar logo centrado
                logo_label = tk.Label(content_frame, image=self.logo_photo, bg="white")
                logo_label.pack(pady=(15, 25))
            except Exception as e:
                debug_log(f"Error cargando logo: {e}")
                # Si no se puede cargar el logo, mostrar título alternativo
                tk.Label(
                    content_frame,
                    text="LABELS APP",
                    font=("Segoe UI", 28, "bold"),
                    fg="#1565C0",
                    bg="white"
                ).pack(pady=(25, 25))
        else:
            # Título principal sin logo
            tk.Label(
                content_frame,
                text="LABELS APP",
                font=("Segoe UI", 28, "bold"),
                fg="#1565C0",
                bg="white"
            ).pack(pady=(25, 25))
        
        # Frame para el formulario con estilo moderno
        form_frame = tk.Frame(content_frame, bg="white")
        form_frame.pack(fill="x", pady=(0, 20))
        
        # Usuario con estilo moderno
        tk.Label(
            form_frame, 
            text="Usuario", 
            font=("Segoe UI", 11, "normal"), 
            fg="#333333",
            bg="white"
        ).pack(anchor="w", pady=(0, 8))
        
        entry_usuario = tk.Entry(
            form_frame, 
            font=("Segoe UI", 12), 
            relief="solid",
            bd=1,
            highlightthickness=2,
            highlightcolor="#1565C0",
            bg="white",
            fg="#333333"
        )
        entry_usuario.pack(fill="x", ipady=8, pady=(0, 20))
        
        # Contraseña con estilo moderno
        tk.Label(
            form_frame, 
            text="Contraseña", 
            font=("Segoe UI", 11, "normal"), 
            fg="#333333",
            bg="white"
        ).pack(anchor="w", pady=(0, 8))
        
        entry_password = tk.Entry(
            form_frame, 
            font=("Segoe UI", 12), 
            show="*",
            relief="solid",
            bd=1,
            highlightthickness=2,
            highlightcolor="#1565C0",
            bg="white",
            fg="#333333"
        )
        entry_password.pack(fill="x", ipady=8, pady=(0, 15))
        
        # Área de mensajes de error/estado (inicialmente oculta)
        mensaje_frame = tk.Frame(form_frame, bg="white")
        mensaje_frame.pack(fill="x", pady=(0, 15))
        
        label_mensaje = tk.Label(
            mensaje_frame,
            text="",
            font=("Segoe UI", 10),
            bg="white",
            wraplength=300,
            justify="center"
        )
        label_mensaje.pack()
        
        def mostrar_mensaje(mensaje, tipo="error"):
            if tipo == "error":
                label_mensaje.configure(text=mensaje, fg="#d32f2f")
            elif tipo == "exito":
                label_mensaje.configure(text=mensaje, fg="#388e3c")
            
            tiempo = 5000 if tipo == "error" else 2000
            
            def limpiar_login():
                try:
                    label_mensaje.configure(text="")
                except:
                    pass  # Ignorar si la ventana ya se cerró
            
            ventana_login.after(tiempo, limpiar_login)
        
        def procesar_login():
            username = entry_usuario.get().strip()
            password = entry_password.get()
            
            if not username or not password:
                mostrar_mensaje("Por favor ingresa usuario y contraseña")
                return
            
            resultado = self.verificar_credenciales(username, password)
            
            if "error" in resultado:
                mostrar_mensaje(resultado["error"])
                return
            
            # Login exitoso
            self.usuario_actual = resultado
            self.sucursal = resultado.get('sucursal_nombre', 'SUCURSAL PRINCIPAL')
            
            mostrar_mensaje(f"¡Bienvenido {resultado['nombre_completo']}!", "exito")
            
            def cerrar_ventana():
                try:
                    ventana_login.destroy()
                except:
                    pass  # Ignorar si la ventana ya se cerró
            
            ventana_login.after(1500, cerrar_ventana)
        
        # Botón de login moderno con hover effect
        btn_login = tk.Button(
            form_frame,
            text="INICIAR SESIÓN",
            font=("Segoe UI", 13, "bold"),
            bg="#1565C0",
            fg="white",
            relief="flat",
            bd=0,
            cursor="hand2",
            command=procesar_login
        )
        btn_login.pack(fill="x", ipady=12, pady=(0, 15))
        
        # Agregar efectos hover al botón
        btn_login.bind("<Enter>", lambda e: btn_login.configure(bg="#1976D2"))
        btn_login.bind("<Leave>", lambda e: btn_login.configure(bg="#1565C0"))
        
        # Botón de debug temporal (se puede quitar después)
        btn_debug = tk.Button(
            form_frame,
            text="🔧 DEBUG: Verificar BD",
            font=("Segoe UI", 9),
            bg="#ffc107",
            fg="black",
            relief="flat",
            bd=0,
            cursor="hand2",
            command=lambda: self.debug_verificar_bd()
        )
        btn_debug.pack(fill="x", ipady=8, pady=(5, 0))
        
        # Información de credenciales por defecto con diseño moderno
        info_frame = tk.Frame(content_frame, bg="#f8f9fa", relief="flat", bd=0)
        info_frame.pack(fill="x", pady=(10, 20), padx=0)
        
        tk.Label(
            info_frame,
            text="💡 Credenciales por defecto",
            font=("Segoe UI", 10, "bold"),
            fg="#495057",
            bg="#f8f9fa"
        ).pack(pady=(15, 5))
        
        tk.Label(
            info_frame,
            text="Facturador: caja_churchill  |  Contraseña: hello",
            font=("Segoe UI", 9),
            fg="#6c757d",
            bg="#f8f9fa"
        ).pack(pady=(0, 5))
        
        tk.Label(
            info_frame,
            text="Admin: admin  |  Contraseña: admin123",
            font=("Segoe UI", 9),
            fg="#6c757d",
            bg="#f8f9fa"
        ).pack(pady=(0, 15))
        
        # Información de roles con diseño más limpio
        roles_frame = tk.Frame(content_frame, bg="white")
        roles_frame.pack(fill="x", pady=(0, 10))
        
        tk.Label(
            roles_frame,
            text="Roles disponibles:",
            font=("Segoe UI", 10, "bold"),
            fg="#333333",
            bg="white"
        ).pack(anchor="w", pady=(0, 8))
        
        roles_info = [
            "👑 Administrador: Acceso completo al sistema",
            "🧾 Facturador: Impresión de etiquetas y gestión"
        ]
        
        for info in roles_info:
            tk.Label(
                roles_frame,
                text=info,
                font=("Segoe UI", 9),
                fg="#666666",
                bg="white"
            ).pack(anchor="w", pady=1)
        
        # Bind Enter
        entry_password.bind("<Return>", lambda e: procesar_login())
        entry_usuario.bind("<Return>", lambda e: entry_password.focus())
        
        entry_usuario.focus()
        ventana_login.mainloop()
        
        return self.usuario_actual is not None
    
    def debug_verificar_bd(self):
        """Método de debug para verificar la base de datos"""
        # Verificación silenciosa de base de datos
        conn = self.conectar_bd()
        if not conn:
            return
        
        try:
            cur = conn.cursor()
            # Verificación rápida sin mensajes
            cur.execute("SELECT COUNT(*) FROM usuarios WHERE activo = true")
            usuarios_activos = cur.fetchone()[0]
            cur.close()
            conn.close()
                
        except Exception as e:
            pass  # Verificación silenciosa
        finally:
            if conn:
                conn.close()

# Función para ejecutar el login
def ejecutar_login():
    """Ejecuta el sistema de login y retorna la información del usuario"""
    sistema_login = SistemaLoginIntegrado()
    
    if sistema_login.mostrar_login():
        return sistema_login.usuario_actual, sistema_login.sucursal
    else:
        # Si se cancela el login, salir de la aplicación
        sys.exit(0)

   # Intentar importar win32print con manejo de errores
try:
    import win32print
    import win32api
    WIN32_AVAILABLE = True
except ImportError as e:
    WIN32_AVAILABLE = False


# === DB: carga productos desde ProductSW ===
def obtener_productos_desde_db():
    try:
        conn = psycopg2.connect(
            host="dpg-d1b18u8dl3ps73e68v1g-a.oregon-postgres.render.com",
            port=5432,
            database="labels_app_db",
            user="admin",
            password="KCFjzM4KYzSQx63ArufESIXq03EFXHz3",
            sslmode="require"
        )
        cur = conn.cursor()
        # Solo obtener productos activos
        cur.execute("SELECT codigo, nombre, base, ubicacion FROM ProductSW WHERE activo = TRUE;")
        datos = cur.fetchall()
        cur.close()
        conn.close()
        return datos
    except Exception as e:
        return []
    

# === Consulta a la base de datos ===
def obtener_datos_por_pintura(pintura_id):
    try:
        conn = psycopg2.connect(
            host="dpg-d1b18u8dl3ps73e68v1g-a.oregon-postgres.render.com",
            port=5432,
            database="labels_app_db",
            user="admin",
            password="KCFjzM4KYzSQx63ArufESIXq03EFXHz3",
            sslmode="require"
        )
        with conn.cursor() as cur:
            cur.execute("""
                SELECT 
                    p.id AS codigo,
                    p.base,
                    c.nombre AS colorante,
                    pr.tipo,
                    pr.oz,
                    pr._32s,
                    pr._64s,
                    pr._128s
                FROM presentacion pr
                JOIN pintura p ON pr.id_pintura = p.id
                JOIN colorante c ON pr.id_colorante = c.id
                WHERE p.id = %s;
            """, (pintura_id,))
            return cur.fetchall()
    except Exception as e:
        return []

def obtener_datos_por_tinte(tinte_id):
    try:
        conn = psycopg2.connect(
            host="dpg-d1b18u8dl3ps73e68v1g-a.oregon-postgres.render.com",
            port=5432,
            database="labels_app_db",
            user="admin",
            password="KCFjzM4KYzSQx63ArufESIXq03EFXHz3",
            sslmode="require"
        )
        with conn.cursor() as cur:
            cur.execute("""
                SELECT 
                    t.id AS codigo,
                    t.nombre_color,
                    c.nombre AS colorante,
                    p.tipo,
                    p.cantidad
                FROM presentacion_tintes p
                JOIN tintes t ON p.id_tinte = t.id
                JOIN colorantes_tintes c ON p.id_colorante_tinte = c.codigo
                WHERE t.id = %s;
            """, (tinte_id,))
            return cur.fetchall()
    except Exception as e:
        return []




def obtener_sufijo_presentacion(presentacion):
    """Devuelve el sufijo correspondiente a la presentación seleccionada"""
    sufijos = {
        "Medio Galón": "1/2",
        "Cuarto": "QT",
        "Galón": "1",
        "Cubeta": "5",
        "1/8": "1/8"
    }
    return sufijos.get(presentacion, "")



def mostrar_codigo_base():
    base = base_var.get()
    producto = producto_var.get()
    terminacion = terminacion_var.get()
    presentacion = presentacion_var.get()

    if not base or not producto or not terminacion:
        aviso_var.set("Completa todos los campos")
        return

    # Obtener código base
    resultado = obtener_codigo_base(base, producto, terminacion)
    
    # Agregar sufijo de presentación si está seleccionada
    if presentacion and resultado != "No encontrado" and resultado != "No Aplica":
        sufijo_presentacion = obtener_sufijo_presentacion(presentacion)
        if sufijo_presentacion:
            resultado += sufijo_presentacion

    # Copiar al portapeles y mostrar avisos
    app.clipboard_clear()
    app.clipboard_append(resultado)
    codigo_base_var.set(resultado)
    aviso_var.set("Código facturación copiado en el portapapeles")
    aviso_var.set("Copiado al portapapeles")
    actualizar_vista()

    # Borrar el mensaje después de 3 segundos
    limpiar_mensaje_despues(3000)


# === Gestión de temporizadores ===
timer_id = None  # Variable global para almacenar el ID del temporizador actual

def limpiar_mensaje_despues(milisegundos):
    """Función centralizada para limpiar mensajes después de un tiempo"""
    global timer_id
    
    # Cancelar temporizador anterior si existe
    if timer_id is not None:
        try:
            app.after_cancel(timer_id)
        except:
            pass  # Ignorar errores si el temporizador ya no existe
    
    # Crear nuevo temporizador
    def limpiar():
        global timer_id
        aviso_var.set("")
        timer_id = None
    
    timer_id = app.after(milisegundos, limpiar)

# === Rutas y configuración ===
def obtener_ruta_absoluta(rel_path):
    """Obtiene la ruta correcta de un archivo tanto para scripts como para ejecutables"""
    try:
        # Si es un ejecutable (PyInstaller) con recursos embebidos
        if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
            # PyInstaller crea una carpeta temporal con los recursos
            base_path = sys._MEIPASS
            ruta_recurso = os.path.join(base_path, rel_path)
            if os.path.exists(ruta_recurso):
                return ruta_recurso
        
        # Si es un ejecutable sin _MEIPASS, buscar en directorio del ejecutable
        if getattr(sys, 'frozen', False):
            base_path = os.path.dirname(sys.executable)
            ruta_recurso = os.path.join(base_path, rel_path)
            if os.path.exists(ruta_recurso):
                return ruta_recurso
        
        # Si es un script de Python normal
        base_path = os.path.dirname(os.path.abspath(__file__))
        ruta_recurso = os.path.join(base_path, rel_path)
        if os.path.exists(ruta_recurso):
            return ruta_recurso
        
        # Buscar en directorio de usuario como respaldo
        user_path = os.path.expanduser("~/.etiquetas_app")
        os.makedirs(user_path, exist_ok=True)
        return os.path.join(user_path, rel_path)
        
    except Exception as e:
        # En caso de error, usar directorio de usuario
        user_path = os.path.expanduser("~/.etiquetas_app")
        os.makedirs(user_path, exist_ok=True)
        return os.path.join(user_path, rel_path)

def cargar_sucursal():
    """Carga la sucursal desde parámetros del sistema de login o archivo local"""
    try:
        # Primero verificar si se pasó información desde el sistema de login
        if len(sys.argv) > 1:
            try:
                # Formato esperado: "usuario_id|username|sucursal_nombre"
                params = sys.argv[1].split('|')
                if len(params) >= 3 and params[2].strip():
                    sucursal_desde_login = params[2].strip()
                    return sucursal_desde_login
                else:
                    pass
            except Exception as e:
                pass
        
        # Fallback: usar archivo local para la sucursal
        config_path = obtener_ruta_absoluta("sucursal.txt")
        if os.path.exists(config_path):
            with open(config_path, "r", encoding="utf-8") as f:
                s = f.read().strip()
                if s:
                    return s
        
        # Último recurso: usar sucursal por defecto para pruebas
        # Cuando se ejecuta directamente sin login, no solicitar sucursal
        return "SUCURSAL PRINCIPAL"
                
    except Exception as e:
        # Fallback final al archivo local
        config_path = obtener_ruta_absoluta("sucursal.txt")
        if os.path.exists(config_path):
            with open(config_path, "r", encoding="utf-8") as f:
                s = f.read().strip()
                if s:
                    return s
        # Usar sucursal por defecto sin solicitar al usuario
        return "SUCURSAL PRINCIPAL"

def obtener_icono_path():
    if getattr(sys, 'frozen', False):
        return os.path.join(sys._MEIPASS, "icono.ico")
    else:
        return os.path.abspath("icono.ico")



# === Configuración inicial ===
CSV_PATH = obtener_ruta_absoluta("etiquetas_guardadas.csv")
IMPRESORA_CONF_PATH = obtener_ruta_absoluta("config_impresora.txt")
PERSONALIZADOS_PATH = obtener_ruta_absoluta("productos_personalizados.csv")
LOGO_PATH = obtener_ruta_absoluta("logo.png")
SUCURSAL = cargar_sucursal()

# Variables globales para información de usuario desde login integrado
USUARIO_ID = None
USUARIO_USERNAME = None
SUCURSAL_USUARIO = None
USUARIO_ROL = None

# Ejecutar sistema de login integrado
# Iniciando sistema de login integrado sin mensajes
usuario_info, sucursal_info = ejecutar_login()

if usuario_info:
    USUARIO_ID = str(usuario_info['id'])
    USUARIO_USERNAME = usuario_info['username']
    SUCURSAL_USUARIO = sucursal_info
    USUARIO_ROL = usuario_info['rol']
    # Sobreescribir SUCURSAL con la del usuario autenticado
    SUCURSAL = sucursal_info
else:
    sys.exit(1)

def cargar_productos_personalizados():
    """Carga productos personalizados desde archivo CSV local"""
    try:
        if os.path.exists(PERSONALIZADOS_PATH):
            df = pd.read_csv(PERSONALIZADOS_PATH)
            productos = []
            for _, row in df.iterrows():
                # Convertir todos los valores a string y filtrar NaN
                codigo = str(row['codigo']) if pd.notna(row['codigo']) else ""
                nombre = str(row['nombre']) if pd.notna(row['nombre']) else ""
                base = str(row['base']) if pd.notna(row['base']) else ""
                ubicacion = str(row['ubicacion']) if pd.notna(row['ubicacion']) else ""
                
                if codigo and nombre:  # Solo agregar si tienen código y nombre válidos
                    productos.append((codigo, nombre, base, ubicacion))
            
            return productos
        return []
    except Exception as e:
        return []

# === Carga de datos ===
datos = obtener_productos_desde_db()
# Filtrar valores None, NaN y convertir a string
codigos = [str(r[0]) for r in datos if r[0] is not None]
nombres = [str(r[1]) for r in datos if r[1] is not None and str(r[1]) != 'nan']

data_por_codigo = {}
data_por_nombre = {}

for r in datos:
    if r[0] is not None and r[1] is not None and str(r[1]) != 'nan':
        codigo = str(r[0])
        nombre = str(r[1])
        base = str(r[2]) if r[2] is not None else ""
        ubicacion = str(r[3]) if r[3] is not None else ""
        
        data_por_codigo[codigo] = {"nombre": nombre, "base": base, "ubicacion": ubicacion}
        data_por_nombre[nombre] = {"codigo": codigo, "base": base, "ubicacion": ubicacion}

# Cargar productos personalizados inmediatamente y combinarlos
productos_personalizados = cargar_productos_personalizados()
for producto in productos_personalizados:
    codigo, nombre, base, ubicacion = producto
    codigo = str(codigo)
    nombre = str(nombre)
    
    if codigo not in data_por_codigo:  # Evitar duplicados
        codigos.append(codigo)
        nombres.append(nombre)
        data_por_codigo[codigo] = {"nombre": nombre, "base": base, "ubicacion": ubicacion}
        data_por_nombre[nombre] = {"codigo": codigo, "base": base, "ubicacion": ubicacion}

def recargar_productos():
    """Recarga los productos activos desde la base de datos"""
    global datos, codigos, nombres, data_por_codigo, data_por_nombre
    
    datos = obtener_productos_desde_db()
    # Filtrar valores None, NaN y convertir a string
    codigos = [str(r[0]) for r in datos if r[0] is not None]
    nombres = [str(r[1]) for r in datos if r[1] is not None and str(r[1]) != 'nan']
    
    data_por_codigo = {}
    data_por_nombre = {}
    
    for r in datos:
        if r[0] is not None and r[1] is not None and str(r[1]) != 'nan':
            codigo = str(r[0])
            nombre = str(r[1])
            base = str(r[2]) if r[2] is not None else ""
            ubicacion = str(r[3]) if r[3] is not None else ""
            
            data_por_codigo[codigo] = {"nombre": nombre, "base": base, "ubicacion": ubicacion}
            data_por_nombre[nombre] = {"codigo": codigo, "base": base, "ubicacion": ubicacion}
    
    # Cargar productos personalizados y combinarlos
    productos_personalizados = cargar_productos_personalizados()
    for producto in productos_personalizados:
        codigo, nombre, base, ubicacion = producto
        codigo = str(codigo)
        nombre = str(nombre)
        
        if codigo not in data_por_codigo:  # Evitar duplicados
            codigos.append(codigo)
            nombres.append(nombre)
            data_por_codigo[codigo] = {"nombre": nombre, "base": base, "ubicacion": ubicacion}
            data_por_nombre[nombre] = {"codigo": codigo, "base": base, "ubicacion": ubicacion}
    
    # Actualizar las listas de autocompletado con filtrado seguro solo si existen
    if 'codigo_entry' in globals() and hasattr(codigo_entry, 'lista'):
        codigo_entry.lista = sorted(set([c for c in codigos if c and str(c) != 'nan']))
    if 'descripcion_entry' in globals() and hasattr(descripcion_entry, 'lista'):
        descripcion_entry.lista = sorted(set([n for n in nombres if n and str(n) != 'nan']))

def guardar_producto_personalizado(codigo, nombre, base, ubicacion):
    """Guarda un nuevo producto personalizado"""
    try:
        # Cargar datos existentes o crear DataFrame vacío
        if os.path.exists(PERSONALIZADOS_PATH):
            df = pd.read_csv(PERSONALIZADOS_PATH)
        else:
            df = pd.DataFrame(columns=['codigo', 'nombre', 'base', 'ubicacion'])
        
        # Verificar si el código ya existe
        if codigo in df['codigo'].values:
            return False, "El código ya existe en productos personalizados"
        
        # Agregar nuevo producto
        nuevo_producto = pd.DataFrame([{
            'codigo': codigo,
            'nombre': nombre,
            'base': base,
            'ubicacion': ubicacion
        }])
        
        df = pd.concat([df, nuevo_producto], ignore_index=True)
        df.to_csv(PERSONALIZADOS_PATH, index=False)
        
        return True, "Producto personalizado guardado exitosamente"
        
    except Exception as e:
        return False, f"Error al guardar producto: {e}"

def eliminar_producto_personalizado(codigo):
    """Elimina un producto personalizado"""
    try:
        if not os.path.exists(PERSONALIZADOS_PATH):
            return False, "No hay productos personalizados"
        
        df = pd.read_csv(PERSONALIZADOS_PATH)
        if codigo not in df['codigo'].values:
            return False, "Código no encontrado en productos personalizados"
        
        df = df[df['codigo'] != codigo]
        df.to_csv(PERSONALIZADOS_PATH, index=False)
        
        return True, "Producto personalizado eliminado"
        
    except Exception as e:
        return False, f"Error al eliminar producto: {e}"

def abrir_ventana_personalizados():
    """Abre la ventana para gestionar productos personalizados"""
    ventana_pers = tk.Toplevel(app)
    ventana_pers.title("Productos Personalizados")
    ventana_pers.geometry("650x450")  # Aumentado de 600 a 650 para dar más espacio
    ventana_pers.resizable(False, False)
    ventana_pers.configure(bg="#f5f5f5")
    
    # Agregar ícono a la ventana
    try:
        ventana_pers.iconbitmap(ICONO_PATH)
    except Exception as e:
        pass
    
    # Centrar ventana
    ventana_pers.update_idletasks()
    x = (ventana_pers.winfo_screenwidth() // 2) - (650 // 2)  # Actualizado para nuevo ancho
    y = (ventana_pers.winfo_screenheight() // 2) - (450 // 2)
    ventana_pers.geometry(f"650x450+{x}+{y}")
    
    # Frame principal
    main_frame = ttk.Frame(ventana_pers, padding=15)  # Reducido de 20 a 15
    main_frame.pack(fill="both", expand=True)
    
    # Título
    ttk.Label(main_frame, text="Gestión de Productos Personalizados", 
              font=("Segoe UI", 16, "bold")).pack(pady=(0, 15))  # Reducido de 20 a 15
    
    # Frame para agregar nuevo producto
    add_frame = ttk.LabelFrame(main_frame, text="Agregar Nuevo Producto", padding=12)  # Reducido de 15 a 12
    add_frame.pack(fill="x", pady=(0, 15))  # Reducido de 20 a 15
    
    # Variables para los campos
    codigo_pers_var = tk.StringVar()
    nombre_pers_var = tk.StringVar()
    
    # Campos de entrada en una sola fila
    ttk.Label(add_frame, text="Código:").grid(row=0, column=0, sticky="w", pady=10, padx=(0, 5))
    ttk.Entry(add_frame, textvariable=codigo_pers_var, width=25).grid(row=0, column=1, pady=10, padx=5)
    
    ttk.Label(add_frame, text="Descripción:").grid(row=0, column=2, sticky="w", pady=10, padx=(15, 5))
    ttk.Entry(add_frame, textvariable=nombre_pers_var, width=35).grid(row=0, column=3, pady=10, padx=5)
    
    def agregar_producto():
        codigo = codigo_pers_var.get().strip().upper()
        nombre = nombre_pers_var.get().strip()
        base = "custom"  # Siempre 'custom' para productos personalizados
        # Calcular ubicación incremental
        if os.path.exists(PERSONALIZADOS_PATH):
            df = pd.read_csv(PERSONALIZADOS_PATH)
            ubicacion = str(len(df) + 1)
        else:
            ubicacion = "1"
        
        if not all([codigo, nombre]):
            messagebox.showwarning("Campos incompletos", "Por favor complete todos los campos")
            return
        
        # Verificar si el código ya existe en la base de datos principal
        if codigo in data_por_codigo:
            messagebox.showwarning("Código existente", "Este código ya existe en la base de datos principal")
            return
        
        exito, mensaje = guardar_producto_personalizado(codigo, nombre, base, ubicacion)
        
        if exito:
            messagebox.showinfo("Éxito", mensaje)
            # Limpiar campos
            codigo_pers_var.set("")
            nombre_pers_var.set("")
            # Actualizar lista
            actualizar_lista_personalizados()
            # Recargar productos en la aplicación principal
            recargar_productos()
        else:
            messagebox.showerror("Error", mensaje)
    
    # Botón agregar centrado
    btn_frame_form = ttk.Frame(add_frame)
    btn_frame_form.grid(row=1, column=0, columnspan=4, pady=15)
    
    ttk.Button(btn_frame_form, text="Agregar Producto", command=agregar_producto,
               style="BotonImprimir.TButton", width=20).pack()
    
    # Frame para lista de productos existentes
    list_frame = ttk.LabelFrame(main_frame, text="Productos Personalizados Existentes", padding=12)  # Reducido de 15 a 12
    list_frame.pack(fill="both", expand=True)
    
    # Treeview para mostrar productos
    columns = ("Código", "Descripción", "Base", "Ubicación")
    tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=14)  # Aumentado de 12 a 14
    
    # Ajustar el ancho de las columnas para mejor distribución
    tree.heading("Código", text="Código")
    tree.column("Código", width=100)
    
    tree.heading("Descripción", text="Descripción")
    tree.column("Descripción", width=200)
    
    tree.heading("Base", text="Base")
    tree.column("Base", width=80)
    
    tree.heading("Ubicación", text="Ubicación")
    tree.column("Ubicación", width=100)
    
    tree.pack(side="left", fill="both", expand=True)
    
    # Scrollbar para el treeview
    scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=tree.yview)
    scrollbar.pack(side="right", fill="y")
    tree.configure(yscrollcommand=scrollbar.set)
    
    def actualizar_lista_personalizados():
        """Actualiza la lista de productos personalizados"""
        for item in tree.get_children():
            tree.delete(item)
        
        productos = cargar_productos_personalizados()
        for producto in productos:
            tree.insert("", "end", values=producto)
    
    def eliminar_seleccionado():
        """Elimina el producto seleccionado"""
        seleccion = tree.selection()
        if not seleccion:
            messagebox.showwarning("Sin selección", "Por favor seleccione un producto para eliminar")
            return
        
        item = tree.item(seleccion[0])
        codigo = item['values'][0]
        
        if messagebox.askyesno("Confirmar eliminación", 
                              f"¿Está seguro de eliminar el producto {codigo}?"):
            exito, mensaje = eliminar_producto_personalizado(codigo)
            
            if exito:
                messagebox.showinfo("Éxito", mensaje)
                # Limpiar campos del formulario
                codigo_pers_var.set("")
                nombre_pers_var.set("")
                # Actualizar lista
                actualizar_lista_personalizados()
                # Recargar productos en la aplicación principal
                recargar_productos()
            else:
                messagebox.showerror("Error", mensaje)
    
    # Frame para botones
    btn_frame = ttk.Frame(list_frame)
    btn_frame.pack(fill="x", pady=(10, 0))
    
    ttk.Button(btn_frame, text="Eliminar", command=eliminar_seleccionado,
               style="BotonGrande.TButton").pack(side="left", padx=5)
    
    # Cargar lista inicial
    actualizar_lista_personalizados()

# === Autocomplete personalizado ===
class AutoCompleteEntry(tk.Entry):
    def __init__(self, master, lista, callback=None, **kwargs):
        super().__init__(master, **kwargs)
        self.lista = sorted(set(lista))
        self.callback = callback
        self.listbox = None
        self.bind('<KeyRelease>', self.check_input)
        self.bind('<Down>', self.focus_listbox)
        self.bind('<Return>', self.select_listbox)

    def check_input(self, event=None):
        txt = self.get().lower()
        if not txt:
            self.close_listbox(); return
        matches = [i for i in self.lista if txt in i.lower()]
        if matches:
            self.show_listbox(matches)
        else:
            self.close_listbox()

    def show_listbox(self, matches):
        if self.listbox: self.listbox.destroy()
        lb = tk.Listbox(self.master, height=5)
        lb.place(x=self.winfo_x(), y=self.winfo_y()+self.winfo_height())
        for m in matches: lb.insert('end', m)
        lb.bind('<<ListboxSelect>>', self.on_select)
        lb.bind('<Return>', self.on_select_keyboard)
        lb.bind('<Escape>', lambda e: self.close_listbox())
        self.listbox = lb

    def on_select(self, e=None):
        sel = self.listbox.get(self.listbox.curselection())
        self.delete(0, 'end')
        self.insert(0, sel)
        self.close_listbox()
        if self.callback:
            self.callback()
        self.focus_set()

    def on_select_keyboard(self, e=None):
        self.on_select(e)

    def focus_listbox(self, event=None):
        if self.listbox:
            self.listbox.focus()
            self.listbox.select_set(0)

    def select_listbox(self, event=None):
        if self.listbox:
            self.on_select()
        else:
            self.event_generate('<Tab>')

    def close_listbox(self):
        if self.listbox:
            self.listbox.destroy()
            self.listbox = None
            

# === ZPL + impresión ===
def generar_zpl(codigo, descripcion, producto, terminacion, presentacion, cantidad=1):
    w, h = 406, 203  # 2x1 pulgadas a 203 dpi
    
    # Ajuste dinámico de fuentes según longitud del contenido
    font_codigo = 70 if len(codigo) <= 6 else 70 if len(codigo) <= 8 else 30
    font_desc = 24 if len(descripcion) > 25 else 28
    
    # Construir texto del producto sin presentación
    producto_completo = '/'.join([x for x in [producto, terminacion] if x])
    font_producto = 20 if len(producto_completo) > 25 else 22 if len(producto_completo) > 20 else 26
    
    # Posiciones optimizadas - movido más a la derecha
    margin = 65  # Incrementado de 55 a 65
    y_cod = 25  # Bajado de 15 a 25
    y_desc = y_cod + font_codigo + 5  # Reducido de 8 a 5
    
    # Calcular posición de producto/terminación dinámicamente
    desc_lines = 1 if len(descripcion) <= 32 else 2
    y_producto = y_desc + (font_desc * desc_lines) + 12  # Reducido de 18 a 12
    
    # === Borde decorativo ===
    border_thickness = 2
    
    # === Sucursal lateral vertical optimizada ===
    sucursal_font_size = 16  # Reducido de 20 a 16
    x_sucursal = 18  # Movido de 8 a 18
    y_sucursal_start = 30
    
    # === Base/Ubicación en la parte inferior ===
    base = base_var.get() if base_var.get() else ""
    ubicacion = ubicacion_var.get() if ubicacion_var.get() else ""
    
    # Productos que no deben mostrar la base
    productos_sin_base = ['laca', 'uretano', 'esmalte kem', 'esmalte multiuso', 'monocapa']
    mostrar_base = not any(prod.lower() in producto.lower() for prod in productos_sin_base)
    
    if mostrar_base:
        info_adicional = f"{base} | {ubicacion}" if base and ubicacion else base or ubicacion
    else:
        # Solo mostrar ubicación para productos sin base
        info_adicional = ubicacion if ubicacion else ""
    
    font_info = 16
    y_info = h - 25

    zpl = (
        "^XA\n"
        "^CI28\n"  # Codificación UTF-8
        f"^PW{w}\n^LL{h}\n^LH0,0\n"
        
        # === BORDE DECORATIVO ===
        f"^FO0,0^GB{w},{border_thickness},B^FS\n"  # Borde superior
        f"^FO0,{h-border_thickness}^GB{w},{border_thickness},B^FS\n"  # Borde inferior
        f"^FO{w-border_thickness},0^GB{border_thickness},{h},B^FS\n"  # Borde derecho
        f"^FO15,0^GB{border_thickness},{h},B^FS\n"  # Borde izquierdo movido hacia la izquierda
        
        # === LÍNEA DECORATIVA SUPERIOR ===
        f"^FO15,15^GB{w-30},1,B^FS\n"  # Línea arriba del código que toca los bordes
        
        # === CÓDIGO PRINCIPAL (Destacado y centrado) ===
        f"^CF0,{font_codigo}\n"
        f"^FO{margin},{y_cod}^FB{w-margin*2-5},1,0,C,0^FD{codigo}^FS\n"
        
        # === DESCRIPCIÓN (Centrada, máximo 2 líneas) ===
        f"^CF0,{font_desc}\n"
        f"^FO{margin},{y_desc}^FB{w-margin*2-5},{desc_lines},0,C,0^FD{descripcion}^FS\n"
        
        # === PRODUCTO/TERMINACIÓN/PRESENTACIÓN (Destacado y centrado) ===
        f"^CF0,{font_producto}\n"
        f"^FO{margin-10},{y_producto}^FB{w-margin*2+15},1,0,C,0^FD{'/'.join([x.upper() for x in [producto, terminacion, presentacion] if x])}^FS\n"
    )
    
    # === INFORMACIÓN ADICIONAL (Base/Ubicación) ===
    if info_adicional:
        zpl += (
            f"^CF0,{font_info}\n"
            f"^FO{margin},{y_info}^FB{w-margin*2-5},1,0,C,0^FD{info_adicional}^FS\n"
        )
    
    # === LÍNEA SEPARADORA ENTRE PRODUCTO Y BASE ===
    y_linea_separadora = y_producto + font_producto + 5  # Subido 5 píxeles
    zpl += f"^FO{margin+20},{y_linea_separadora}^GB{w-margin*2-50},1,B^FS\n"  # Línea más pequeña
    
    # === SUCURSAL LATERAL (Rotada 90°) ===
    if SUCURSAL:
        # Calcular el centro vertical real de la etiqueta
        centro_etiqueta = h // 2  # Centro absoluto de la etiqueta (203/2 = 101.5)
        
        # Calcular la longitud del texto para centrarlo perfectamente
        longitud_texto = len(SUCURSAL) * (sucursal_font_size * 0.5)  # Ajustado de 0.6 a 0.5
        y_inicio_centrado = centro_etiqueta - (longitud_texto // 2)  # Cambiado + por -
        
        zpl += (
            f"^A0R,{sucursal_font_size},{sucursal_font_size}\n"
            f"^FO{x_sucursal},{y_inicio_centrado}^FD{SUCURSAL.upper()}^FS\n"
        )
    
    # === LÍNEA SEPARADORA DECORATIVA ===
    y_linea = y_producto + font_producto + 8
    # Esta línea ya no es necesaria porque agregamos la línea separadora arriba
    # if not info_adicional or y_linea < y_info - 10:
    #     zpl += f"^FO{margin + 10},{y_linea}^GB{w-margin*2-55},1,B^FS\n"
    
    zpl += "^XZ\n"
    
    return zpl * int(cantidad)


# === Generar PDF ===
def generar_pdf_ficha(data, filename="ficha_pintura.pdf"):
    if not data:
        return

    codigo = data[0][0]
    base = data[0][1]

    c = canvas.Canvas(filename, pagesize=landscape(A4))
    width, height = landscape(A4)

    # Título
    c.setFont("Helvetica-Bold", 20)
    c.drawString(40, height - 40, f"Fórmula - Código: {codigo}")
    c.setFont("Helvetica", 14)
    c.drawString(40, height - 65, f"Base: {base}")

    # Encabezado de tabla
    encabezado = [
        ["COLORANTE", "CUARTOS", "", "", "", "GALONES", "", "", "", "CUBETAS", "", "", ""],
        ["", "oz", "32s", "64s", "128s", "oz", "32s", "64s", "128s", "oz", "32s", "64s", "128s"]
    ]

    # Datos organizados por colorante y tipo
    filas = {}
    for _, _, colorante, tipo, oz, _32s, _64s, _128s in data:
        if colorante not in filas:
            filas[colorante] = {
                "cuarto": ["", "", "", ""],
                "galon": ["", "", "", ""],
                "cubeta": ["", "", "", ""]
            }
        filas[colorante][tipo] = [oz, _32s, _64s, _128s]

    # Cuerpo de la tabla
    cuerpo = []
    for colorante, valores in filas.items():
        fila = [colorante]
        for tipo in ["cuarto", "galon", "cubeta"]:
            for i in range(4):
                val = valores[tipo][i]
                if val is None or str(val).lower() == "nan" or (isinstance(val, float) and math.isnan(val)):
                    fila.append("")
                else:
                    try:
                        num = float(val)
                        if num.is_integer():
                            fila.append(str(int(num)))
                        else:
                            fila.append(str(num))
                    except:
                        fila.append(str(val))
        cuerpo.append(fila)

    tabla = encabezado + cuerpo

    # Estilos
    t = Table(tabla, colWidths=[80] + [40]*12)
    t.setStyle(TableStyle([
     ("BACKGROUND", (0, 0), (-1, 0), colors.darkblue),
     ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
     ("BACKGROUND", (1, 1), (4, 1), colors.orange),
     ("BACKGROUND", (5, 1), (8, 1), colors.lightblue),
     ("BACKGROUND", (9, 1), (12, 1), colors.gold),
     ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
     ("ALIGN", (0, 0), (-1, -1), "CENTER"),
     ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
     ("FONTNAME", (0, 0), (-1, 1), "Helvetica-Bold"),
 
       # Span y centrado para CUARTOS, GALONES, CUBETAS
      ("SPAN", (1, 0), (4, 0)),
      ("SPAN", (5, 0), (8, 0)),
      ("SPAN", (9, 0), (12, 0)),
      ("ALIGN", (1, 0), (12, 0), "CENTER"),
      ("VALIGN", (1, 0), (12, 0), "MIDDLE"),
    ]))


    t.wrapOn(c, width, height)
    t.drawOn(c, 40, height - 150 - 25 * len(cuerpo))

    c.save()

def generar_pdf_tinte(data, filename="ficha_tinte.pdf"):
    if not data:
        return

    codigo = data[0][0]
    nombre_color = data[0][1]

    c = canvas.Canvas(filename, pagesize=landscape(A4))
    width, height = landscape(A4)

    # Título
    c.setFont("Helvetica-Bold", 20)
    c.drawString(40, height - 40, f"Tinte - Código: {codigo}")
    c.setFont("Helvetica", 14)
    c.drawString(40, height - 65, f"Nombre del color: {nombre_color}")

    # Orden deseado de las unidades
    orden_tipos = ["1/8", "QT", "1/2", "GALON"]

    # Construimos estructura: {colorante: {tipo: cantidad}}
    estructura = defaultdict(dict)
    tipos_encontrados = set()
    for _, _, colorante, tipo, cantidad in data:
        tipos_encontrados.add(tipo)
        try:
            num = float(cantidad)
            cantidad_str = str(int(num)) if num.is_integer() else str(num)
        except:
            cantidad_str = str(cantidad)
        estructura[colorante][tipo] = cantidad_str

    # Usar solo los tipos en el orden deseado que existan en los datos
    tipos = [t for t in orden_tipos if t in tipos_encontrados]
    colorantes = sorted(estructura.keys())

    # Encabezado: COLORANTE | 1/8 | QT | 1/2 | GALON
    encabezado = ["COLORANTE"] + tipos
    cuerpo = []

    for colorante in colorantes:
        fila = [colorante]
        for tipo in tipos:
            fila.append(estructura[colorante].get(tipo, ""))
        cuerpo.append(fila)

    tabla = [encabezado] + cuerpo

    # Tamaño de columnas dinámico   
    col_widths = [130] + [80] * (len(encabezado) - 1)

    t = Table(tabla, colWidths=col_widths)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.darkblue),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("ALIGN", (1, 1), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
    ]))

    t.wrapOn(c, width, height)
    t.drawOn(c, 40, height - 150 - 25 * len(cuerpo))

    c.save()

def generar_pdf_por_cada_tinte():
    try:
        conn = psycopg2.connect(
            host="dpg-d1b18u8dl3ps73e68v1g-a.oregon-postgres.render.com",
            port=5432,
            database="labels_app_db",
            user="admin",
            password="KCFjzM4KYzSQx63ArufESIXq03EFXHz3",
            sslmode="require"
        )
        with conn.cursor() as cur:
            cur.execute("SELECT id FROM tintes;")
            ids = cur.fetchall()

        for (tinte_id,) in ids:
            data = obtener_datos_por_tinte(tinte_id)
            generar_pdf_tinte(data, filename=f"tinte_{tinte_id}.pdf")

    except Exception as e:
        pass


def imprimir_zebra_zpl(zpl_code):
    if not WIN32_AVAILABLE:
        messagebox.showerror("Error de impresión",
                           "Módulos de impresión no disponibles.\n"
                           "Instala pywin32: pip install pywin32")
        return

    try:
        pr = printer_var.get()
        guardar_impresora(pr)

        # Verificar lista de impresoras disponibles
        try:
            available = [p[2] for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
        except Exception:
            available = []

        if not pr:
            messagebox.showwarning("Impresora no seleccionada", "Por favor selecciona una impresora antes de imprimir.")
            return

        if available and pr not in available:
            messagebox.showwarning("Impresora no encontrada", f"La impresora seleccionada ('{pr}') no está entre las impresoras detectadas.\nLista detectada: {available}")

        # Intentar enviar ZPL en RAW
        h = win32print.OpenPrinter(pr)
        try:
            win32print.StartDocPrinter(h, 1, ("Etiqueta", None, "RAW"))
            win32print.StartPagePrinter(h)
            win32print.WritePrinter(h, zpl_code.encode())
            win32print.EndPagePrinter(h)
            win32print.EndDocPrinter(h)
        finally:
            try:
                win32print.ClosePrinter(h)
            except:
                pass
    except Exception as e:
        # Mostrar error detallado y ofrecer escribir ZPL a archivo como fallback
        try:
            # Guardar ZPL en archivo temporal para envío manual
            import tempfile
            tmp = tempfile.mktemp(suffix='.zpl')
            with open(tmp, 'w', encoding='utf-8') as f:
                f.write(zpl_code)
            messagebox.showerror("Error impresión", f"No se pudo imprimir:\n{e}\nSe guardó el ZPL en: {tmp}")
        except Exception as e2:
            messagebox.showerror("Error impresión", f"No se pudo imprimir:\n{e}\nAdemás, no se pudo crear archivo de fallback: {e2}")

def guardar_impresora(nombre):
    try:
        with open(IMPRESORA_CONF_PATH,'w',encoding='utf-8') as f:
            f.write(nombre)
    except: pass

def cargar_impresora_guardada():
    if not WIN32_AVAILABLE:
        return ''
        
    if os.path.exists(IMPRESORA_CONF_PATH):
        try:
            n = open(IMPRESORA_CONF_PATH,'r',encoding='utf-8').read().strip()
            printers = [p[2] for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL|win32print.PRINTER_ENUM_CONNECTIONS)]
            if n in printers:
                return n
        except: pass
    return ''
def on_btn_imprimir_click():
    codigo = codigo_entry.get()
    if codigo:
        imprimir_ficha_pintura(codigo)
    else:
        messagebox.showinfo("Campo vacío", "Por favor ingrese un código de pintura.")


# === UI ===
app = ttk.Window(themename='flatly')
app.title(f"Acabados & Pinturas {APP_VERSION} - {SUCURSAL}")
app.geometry("900x540")  # Ventana ligeramente más compacta
app.resizable(False, False)

# Configurar icono de la aplicación
ICONO_PATH = obtener_icono_path()
try:
    app.iconbitmap(ICONO_PATH)
except Exception as e:
    pass
    
aviso_var = tk.StringVar()

printer_var = tk.StringVar(value=cargar_impresora_guardada())


# Obtener lista de impresoras con manejo de errores
if WIN32_AVAILABLE:
    try:
        printers = [p[2] for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
    except Exception as e:
        printers = []
else:
    printers = ["Sin impresoras disponibles (instalar pywin32)"]

if not printers:
    messagebox.showwarning("Sin impresoras", "No se detectaron impresoras")

# Variables

descripcion_var = tk.StringVar()
producto_var = tk.StringVar()
terminacion_var = tk.StringVar()
presentacion_var = tk.StringVar()
spin = tk.IntVar(value=1)
base_var = tk.StringVar()
ubicacion_var = tk.StringVar()
codigo_base_var = tk.StringVar()

# Lista temporal de productos para factura múltiple
lista_productos_factura = []

# Actualiza la vista previa al cambiar producto o terminación
producto_var.trace_add('write', lambda *args: actualizar_vista())
terminacion_var.trace_add('write', lambda *args: actualizar_vista())

# Diccionario de terminaciones válidas por producto
TERMINACIONES_POR_PRODUCTO = {
    'laca': ['Mate', 'Semimate', 'Brillo'],
    'esmalte multiuso': ['Mate', 'Satin', 'Gloss'],
    'excello premium': ['Mate', 'Satin', 'Semigloss', 'Semisatin'],
    'excello voc': ['Mate', 'Satin'],
    'master paint': ['Mate'],
    'tinte al thinner': ['Claro', 'Intermedio', 'Especial'],
    'super paint': ['Mate', 'Satin', 'Gloss'],
    'esmalte kem': ['Mate', 'Semimate', 'Brillo'],
    'excello pastel': ['Mate'],
    'water blocking': ['Mate'],
    'kem aqua': ['Satin'],
    'emerald': ['Satin', 'SemiGloss'],
    'monocapa': ['Mate', 'Semimate', 'Brillo'],
    'uretano': ['Mate', 'Semimate', 'Brillo'],
    'airpuretec': ['Mate', 'Satin'],
    'kem pro': ['Mate'],
    'sanitizing': ['Satin'],
    'scuff tuff-wb' : ['Mate', 'Satin',],
    'armoseal t-p' : ['Semigloss'],
    'armoseal 1000hs' : ['Gloss'],
    'pro industrial dtm' : ['Gloss'],
    'promar® 400 voc' : ['Satin'],
    'h&c heavy-shield' : ['Gloss'],
    'h&c silicone-acrylic' : ['Mate'],
    'promar® 200 voc' : ['Satin', 'Mate'],

    
    
}

def actualizar_terminaciones(*args):
    """Actualiza las terminaciones disponibles según el producto seleccionado"""
    producto = producto_var.get().lower()
    base = base_var.get().lower()
    
    # Buscar terminaciones válidas para el producto
    terminaciones_validas = []
    for key, terminaciones in TERMINACIONES_POR_PRODUCTO.items():
        if key in producto:
            terminaciones_validas = terminaciones
            break
    
    # Si no se encuentra el producto, usar todas las terminaciones
    if not terminaciones_validas:
        terminaciones_validas = ['Mate', 'Satin', 'Semigloss', 'Semimate', 'Gloss', 'Brillo', 
                                "N/A", "ESPECIAL", "CLARO", "INTERMEDIO", "MADERA", "PERLADO", "METALICO", "SEMISATIN"]
    
    # Lógica especial para Excello Premium con bases Ultra Deep / Ultra Deep II
    if 'excello premium' in producto:
        # Detectar Ultra Deep y Ultra Deep II (distintas variantes)
        es_ultra_deep_ii = any(k in base for k in ['ultra deep ii', 'ultradeep ii', 'ultra-deep ii', 'ultra deep 2'])
        es_ultra_deep = ('ultra deep' in base)
        if es_ultra_deep or es_ultra_deep_ii:
            # Solo permitir Semisatin para ambas bases
            terminaciones_validas = ['Semisatin']

    # Actualizar el combobox
    terminaciones_combobox['values'] = terminaciones_validas
    
    # Limpiar la selección actual si no es válida
    terminacion_actual = terminacion_var.get()
    if terminacion_actual and terminacion_actual not in terminaciones_validas:
        terminacion_var.set('')
    
    # Si solo hay una terminación válida, seleccionarla automáticamente
    if len(terminaciones_validas) == 1:
        terminacion_var.set(terminaciones_validas[0])
    
    # Actualizar vista previa
    actualizar_vista()

def actualizar_presentaciones(*args):
    """Actualiza las presentaciones disponibles según el producto seleccionado"""
    producto = producto_var.get().lower()
    
    # Presentaciones por defecto (incluye Medio Galón)
    presentaciones_disponibles = ['Cuarto', 'Medio Galón', 'Galón', 'Cubeta']
    
    # Si es laca, agregar octavos (1/8)
    if 'laca' in producto:
        presentaciones_disponibles = ['1/8', 'Cuarto', 'Medio Galón', 'Galón', 'Cubeta']
    
    # Actualizar el combobox
    presentacion_combobox['values'] = presentaciones_disponibles
    
    # Limpiar la selección actual si no es válida
    presentacion_actual = presentacion_var.get()
    if presentacion_actual and presentacion_actual not in presentaciones_disponibles:
        presentacion_var.set('')

# Actualiza la vista previa al cambiar producto o terminación
producto_var.trace_add('write', actualizar_terminaciones)
producto_var.trace_add('write', actualizar_presentaciones)
terminacion_var.trace_add('write', lambda *args: actualizar_vista())
presentacion_var.trace_add('write', lambda *args: actualizar_vista())
# Agregar listener para cuando cambie la base (importante para Excello Premium)
base_var.trace_add('write', actualizar_terminaciones)

# Layout
labels = ["Código", "Descripción", "Producto", "Terminación", "Presentación", "Ubicación", "Base", "Codigo Base", "Cantidad"]
for i, l in enumerate(labels):
    ttk.Label(app, text=f"{l}:").place(x=30, y=20 + i * 50)

codigo_entry = AutoCompleteEntry(app, sorted(set([c for c in codigos if c and str(c) != 'nan'])), callback=lambda: completar_datos())
codigo_entry.place(x=150, y=20, width=200)

descripcion_entry = AutoCompleteEntry(app, sorted(set([n for n in nombres if n and str(n) != 'nan'])), callback=lambda: completar_datos(), textvariable=descripcion_var)
descripcion_entry.place(x=150, y=70, width=200)

# Guarda las referencias de los combobox
producto_combobox = ttk.Combobox(app, textvariable=producto_var,
    values=['Excello Premium','Laca', 'Esmalte Kem', "Excello VOC","Master Paint","Tinte al Thinner", 'Super Paint', 'Esmalte Multiuso','Excello Pastel', 'Water Blocking', 'Kem Aqua', 'Emerald', 'Monocapa', 'Uretano', 'Airpuretec', 'Kem Pro', 'Sanitizing', 'h&c silicone-acrylic', 'h&c heavy-shield', 'promar® 200 voc', 'promar® 400 voc', 'pro industrial dtm', 'armoseal 1000hs','armoseal t-p', 'scuff tuff-wb' ],
    state='readonly')
producto_combobox.place(x=150, y=120, width=200)

terminaciones_combobox = ttk.Combobox(app, textvariable=terminacion_var,
    values=['Mate', 'Satin', 'Semigloss', 'Semimate', 'Gloss', 'Brillo', "N/A", "ESPECIAL", "CLARO", "INTERMEDIO", "ESPECIAL", "MADERA", "PERLADO", "METALICO"],
    state='readonly')
terminaciones_combobox.place(x=150, y=170, width=200)

# Combobox de presentación
presentacion_combobox = ttk.Combobox(app, textvariable=presentacion_var,
    values=['1/8', 'Cuarto', 'Medio Galón', 'Galón', 'Cubeta'],
    state='readonly')
presentacion_combobox.place(x=150, y=220, width=200)

# Eliminado el selector de impresora (se conserva la lectura/guardado silencioso)


# Función para soporte de teclado en combobox
def combobox_keydown(event, combobox):
    if event.keysym == 'Down':
        combobox.event_generate('<Button-1>')
    elif event.keysym == 'Return':
        combobox.event_generate('<Tab>')

# Bind a todos los combobox
producto_combobox.bind('<Key>', lambda e: combobox_keydown(e, producto_combobox))
terminaciones_combobox.bind('<Key>', lambda e: combobox_keydown(e, terminaciones_combobox))
# Eliminado el bind del selector de impresora

ttk.Entry(app, textvariable=ubicacion_var, state='readonly').place(x=150, y=270, width=200)
ttk.Entry(app, textvariable=base_var, state='readonly').place(x=150, y=320, width=200)
ttk.Entry(app, textvariable=codigo_base_var, state='readonly').place(x=150, y=370, width=200)
ttk.Spinbox(app, from_=1, to=100, textvariable=spin).place(x=150, y=420, width=60)





# Vista Previa dentro de un LabelFrame igual que Acciones
frame_vista = ttk.LabelFrame(app, text="Vista Previa", padding=14, style="Acciones.TLabelframe")
frame_vista.place(x=420, y=15, width=440, height=260)
# Canvas para dibujar una vista previa que imita la ZPL del gestor (406x203)
vista_canvas = tk.Canvas(frame_vista, width=406, height=203, bg="#ffffff", highlightthickness=0)
vista_canvas.pack(anchor='center')
_vista_imgs_cache = {}

# Label para avisos debajo de la vista previa
aviso_label = ttk.Label(app, textvariable=aviso_var, font=('Segoe UI', 10), foreground="#1976d2", background="#fff")
aviso_label.place(x=500, y=270, width=380)  # Subido más para dejar espacio al cuadro de acciones

codigo_base_actual = ""

def actualizar_vista():
    """Dibuja una vista previa alineada a la etiqueta del gestor (layout ZPL simulado)."""
    # Limpiar canvas
    vista_canvas.delete("all")

    # Medidas de etiqueta ZPL
    w, h = 406, 203
    margin = 65
    border = 2

    # Datos actuales
    c = codigo_entry.get().strip().upper()
    d = descripcion_var.get().strip()
    p = (producto_var.get() or '').strip()
    t = (terminacion_var.get() or '').strip()
    pr = (presentacion_var.get() or '').strip()
    b = (base_var.get() or '').strip()
    u = (ubicacion_var.get() or '').strip()

    # Regla del gestor: mostrar PRODUCTO/TERMINACIÓN (sin presentación)
    producto_linea = '/'.join([x for x in [p, t] if x])

    # Productos que no deben mostrar la base (coincide con lógica ZPL)
    productos_sin_base = ['laca', 'uretano', 'esmalte kem', 'esmalte multiuso', 'monocapa']
    mostrar_base = not any(prod.lower() in p.lower() for prod in productos_sin_base)
    info_adicional = f"{b} | {u}" if mostrar_base and b and u else (b if mostrar_base else (u or ""))

    # Cálculo de líneas de descripción (máx 2)
    desc_lines = 1 if len(d) <= 32 else 2
    desc1 = d[:32]
    desc2 = d[32:64] if desc_lines == 2 else ''

    # Tipografías aproximadas
    from tkinter import font as tkfont
    font_codigo = tkfont.Font(family='Consolas', size=28, weight='bold')
    font_desc = tkfont.Font(family='Segoe UI', size=14, weight='bold')
    font_prod = tkfont.Font(family='Segoe UI', size=16, weight='bold')
    font_info = tkfont.Font(family='Segoe UI', size=12)
    font_sucursal = tkfont.Font(family='Segoe UI', size=12, weight='bold')

    # Fondo y bordes
    vista_canvas.create_rectangle(0, 0, w, h, outline='#000000', width=border)
    # Línea decorativa superior interna
    vista_canvas.create_line(15, 15, w-15, 15, fill='#000000')

    # Código principal (centrado)
    y_cod = 25
    vista_canvas.create_text(w//2, y_cod, text=c or '', font=font_codigo, fill='#000000', anchor='n')

    # Descripción (1-2 líneas centradas)
    y_desc = y_cod + 28 + 5
    vista_canvas.create_text(w//2, y_desc, text=desc1, font=font_desc, fill='#000000', anchor='n')
    if desc2:
        vista_canvas.create_text(w//2, y_desc + font_desc.metrics('linespace'), text=desc2, font=font_desc, fill='#000000', anchor='n')

    # Producto/Terminación (centrado, sin presentación)
    y_prod = y_desc + (font_desc.metrics('linespace') * desc_lines) + 12
    vista_canvas.create_text(w//2, y_prod, text=(producto_linea.upper()), font=font_prod, fill='#000000', anchor='n')

    # Info adicional (Base/Ubicación) centrado al pie si aplica
    if info_adicional:
        y_info = h - 25
        vista_canvas.create_text(w//2, y_info, text=info_adicional, font=font_info, fill='#000000', anchor='s')
        # Línea separadora sobre info
        vista_canvas.create_line(margin+20, y_prod + font_prod.metrics('linespace') + 5, w - (margin+20), y_prod + font_prod.metrics('linespace') + 5, fill='#000000')

    # Sin texto lateral de sucursal para una vista previa más limpia


def completar_datos():
    c = codigo_entry.get().strip()
    d = descripcion_entry.get().strip()

    if c in data_por_codigo:
        info = data_por_codigo[c]
        descripcion_var.set(info["nombre"])
        base_var.set(info["base"])
        ubicacion_var.set(info["ubicacion"])
    elif d in data_por_nombre:
        info = data_por_nombre[d]
        codigo_entry.delete(0, 'end')
        codigo_entry.insert(0, info["codigo"])
        descripcion_var.set(d)
        base_var.set(info["base"])
        ubicacion_var.set(info["ubicacion"])
    
    # Limpiar código base para que solo aparezca al presionar el botón
    codigo_base_var.set("")
    
    # Actualizar terminaciones después de completar datos
    actualizar_terminaciones()
    actualizar_vista()

def mostrar_ventana_factura():
    """Muestra ventana emergente para ingresar ID de factura y prioridad"""
    ventana_factura = tk.Toplevel()
    ventana_factura.title("Información de Factura")
    ventana_factura.geometry("400x250")
    ventana_factura.resizable(False, False)
    ventana_factura.grab_set()  # Hacer modal
    
    # Centrar la ventana respecto a la ventana principal
    ventana_factura.transient(app)
    
    # Centrar la ventana en pantalla
    x = (ventana_factura.winfo_screenwidth() // 2) - (400 // 2)
    y = (ventana_factura.winfo_screenheight() // 2) - (250 // 2)
    ventana_factura.geometry(f"400x250+{x}+{y}")
    
    # Variables para almacenar los valores
    id_factura_var = tk.StringVar()
    prioridad_var = tk.StringVar(value="Media")
    resultado = {"continuar": False, "id_factura": "", "prioridad": ""}
    
    # Título
    ttk.Label(ventana_factura, text="Información del Pedido", 
             font=("Segoe UI", 14, "bold")).pack(pady=15)
    
    # Frame para los campos
    frame_campos = ttk.Frame(ventana_factura)
    frame_campos.pack(pady=10, padx=20, fill="x")
    
    # Campo ID Factura
    ttk.Label(frame_campos, text="ID Factura:", font=("Segoe UI", 10)).grid(row=0, column=0, sticky="w", pady=5)
    entry_factura = ttk.Entry(frame_campos, textvariable=id_factura_var, font=("Segoe UI", 10), width=25)
    entry_factura.grid(row=0, column=1, pady=5, padx=(10, 0))
    entry_factura.focus()
    
    # Campo Prioridad
    ttk.Label(frame_campos, text="Prioridad:", font=("Segoe UI", 10)).grid(row=1, column=0, sticky="w", pady=5)
    combo_prioridad = ttk.Combobox(frame_campos, textvariable=prioridad_var, 
                                  values=["Alta", "Media", "Baja"], 
                                  state="readonly", font=("Segoe UI", 10), width=22)
    combo_prioridad.grid(row=1, column=1, pady=5, padx=(10, 0))
    
    # Frame para botones
    frame_botones = ttk.Frame(ventana_factura)
    frame_botones.pack(pady=20)
    
    def aceptar():
        if not id_factura_var.get().strip():
            messagebox.showwarning("Campo requerido", "Debe ingresar el ID de la factura.")
            entry_factura.focus()
            return
        
        resultado["continuar"] = True
        resultado["id_factura"] = id_factura_var.get().strip()
        resultado["prioridad"] = prioridad_var.get()
        ventana_factura.destroy()
    
    def cancelar():
        resultado["continuar"] = False
        ventana_factura.destroy()
    
    # Botones
    ttk.Button(frame_botones, text="Aceptar", command=aceptar, 
              bootstyle="success").pack(side="left", padx=10)
    ttk.Button(frame_botones, text="Cancelar", command=cancelar, 
              bootstyle="secondary").pack(side="left", padx=10)
    
    # Bind Enter para aceptar
    ventana_factura.bind('<Return>', lambda e: aceptar())
    # Bind Escape para cancelar
    ventana_factura.bind('<Escape>', lambda e: cancelar())
    
    # Esperar hasta que se cierre la ventana
    ventana_factura.wait_window()
    
    return resultado

def imprimir_guardar():
    c = codigo_entry.get()
    d = descripcion_var.get()
    p = producto_var.get()
    t = terminacion_var.get()
    pr = presentacion_var.get()
    q = spin.get()
    
    if not c:
        return  # Salir silenciosamente
    
    # VALIDACIÓN OBLIGATORIA: Presentación debe estar seleccionada
    if not pr:
        messagebox.showwarning("Presentación requerida", 
                              "Debe seleccionar una presentación antes de enviar el producto.\n\n" +
                              "Las presentaciones disponibles son:\n" +
                              "• 1/8 (para lacas)\n" +
                              "• Cuarto\n" +
                              "• Medio Galón\n" +
                              "• Galón\n" +
                              "• Cubeta")
        return
    
    # Mostrar ventana para ID factura y prioridad
    datos_factura = mostrar_ventana_factura()
    
    # Si el usuario canceló, no continuar
    if not datos_factura["continuar"]:
        return
    
    id_factura = datos_factura["id_factura"]
    prioridad = datos_factura["prioridad"]
    
    # Guardar en CSV (para compatibilidad)
    reg = {'Codigo': c, 'Descripcion': d, 'Producto': p, 'Terminacion': t, 'Presentacion': pr, 'ID_Factura': id_factura, 'Prioridad': prioridad}
    df = pd.read_csv(CSV_PATH) if os.path.exists(CSV_PATH) else pd.DataFrame()
    df = pd.concat([df, pd.DataFrame([reg])], ignore_index=True)
    df.to_csv(CSV_PATH, index=False)
    
    # Limpiar campos inmediatamente para velocidad
    limpiar_campos()
    
    # Operaciones de BD en paralelo sin bloquear interfaz
    import threading
    def operaciones_bd():
        try:
            # Registrar en base de datos para estadísticas
            registrar_impresion(c, d, p, t, pr, q, SUCURSAL, USUARIO_ID, id_factura, prioridad)
            # Agregar a lista de espera con presentación correcta
            agregar_a_lista_espera(c, p, t, id_factura, prioridad, q, base=base_var.get(), presentacion=pr, ubicacion=ubicacion_var.get())
        except:
            pass  # Operación silenciosa
    
    threading.Thread(target=operaciones_bd, daemon=True).start()

def imprimir_pdf(path_pdf):
    try:
        if WIN32_AVAILABLE:
            # Intentar imprimir usando ShellExecute (funciona con visores asociados)
            try:
                # Preferir ShellExecute desde win32api
                win32api.ShellExecute(0, "print", path_pdf, None, ".", 0)
                return
            except Exception:
                # Fallback a os.startfile con verbo 'print'
                try:
                    os.startfile(path_pdf, 'print')
                    return
                except Exception as e:
                    messagebox.showerror("Error impresión PDF", f"No se pudo imprimir el PDF: {e}")
                    return
        else:
            # Abrir el PDF si no hay soporte Win32 (solo mostrar)
            try:
                os.startfile(path_pdf)
                return
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo abrir el PDF: {e}")
                return
    except Exception as e:
        messagebox.showerror("Error impresión", f"Error al intentar imprimir/abrir el PDF: {e}")

def limpiar_campos():
    global codigo_base_actual
    codigo_entry.delete(0, 'end')
    descripcion_var.set('')
    producto_var.set('')
    terminacion_var.set('')
    presentacion_var.set('')
    spin.set(1)
    base_var.set('')
    ubicacion_var.set('')
    vista_canvas.delete('all')
    codigo_base_actual = ''
    codigo_base_var.set('')
    
def calcular_tiempo_compromiso(producto, cantidad):
    """Calcula el tiempo de compromiso basado en el tipo de producto y cantidad"""
    producto_lower = producto.lower()
    
    # Productos que requieren 10 minutos por unidad
    productos_10_min = ['laca', 'esmalte', 'industrial']
    
    # Verificar si el producto es de 10 minutos
    tiempo_por_unidad = 10 if any(tipo in producto_lower for tipo in productos_10_min) else 6
    
    return tiempo_por_unidad * cantidad

def obtener_tiempo_acumulado_sucursal(sucursal):
    """Obtiene el tiempo acumulado de todas las órdenes pendientes y en proceso de una sucursal"""
    try:
        conn = psycopg2.connect(
            host="dpg-d1b18u8dl3ps73e68v1g-a.oregon-postgres.render.com",
            port=5432,
            database="labels_app_db",
            user="admin",
            password="KCFjzM4KYzSQx63ArufESIXq03EFXHz3",
            sslmode="require"
        )
        cur = conn.cursor()
        
        # Obtener la suma de tiempos estimados de órdenes pendientes y en proceso
        cur.execute("""
            SELECT COALESCE(SUM(tiempo_estimado), 0)
            FROM lista_espera 
            WHERE sucursal = %s AND estado IN ('Pendiente', 'En Proceso')
        """, (sucursal,))
        
        tiempo_acumulado = cur.fetchone()[0]
        cur.close()
        conn.close()
        return tiempo_acumulado
        
    except Exception as e:
        print(f"Error al obtener tiempo acumulado: {e}")
        return 0

def calcular_fecha_compromiso(tiempo_total_minutos):
    """Calcula la fecha de compromiso basada en el tiempo total en minutos"""
    from datetime import datetime, timedelta
    
    # Obtener fecha y hora actual
    ahora = datetime.now()
    
    # Agregar los minutos al tiempo actual
    fecha_compromiso = ahora + timedelta(minutes=tiempo_total_minutos)
    
    return fecha_compromiso

def generar_id_profesional(id_factura=None):
    """Genera un ID profesional basado en factura: todos los productos de la misma factura tendrán el mismo ID"""
    from datetime import datetime
    import random
    import time
    
    # Si no hay ID de factura, usar método anterior como fallback
    if not id_factura:
        # Obtener fecha actual
        ahora = datetime.now()
        
        # Iniciales de días de la semana
        dias_semana = {
            0: 'L',  # Lunes
            1: 'M',  # Martes  
            2: 'X',  # Miércoles
            3: 'J',  # Jueves
            4: 'V',  # Viernes
            5: 'S',  # Sábado
            6: 'D'   # Domingo
        }
        
        inicial_dia = dias_semana[ahora.weekday()]
    else:
        # Usar ID de factura como base
        # Extraer los últimos 4 dígitos de la factura
        import re
        numeros = re.findall(r'\d+', str(id_factura))
        if numeros:
            factura_num = int(numeros[-1])  # Tomar el último número encontrado
            # Usar F + últimos 4 dígitos de la factura
            inicial_dia = 'F'
            # Si el número es muy grande, tomar solo los últimos 4 dígitos
            base_numero = factura_num % 10000
        else:
            # Fallback si no hay números en la factura
            inicial_dia = 'F'
            base_numero = abs(hash(str(id_factura))) % 10000
    
    try:
        # Conectar a base de datos
        conn = psycopg2.connect(
            host="dpg-d1b18u8dl3ps73e68v1g-a.oregon-postgres.render.com",
            port=5432,
            database="labels_app_db",
            user="admin",
            password="KCFjzM4KYzSQx63ArufESIXq03EFXHz3",
            sslmode="require"
        )
        cur = conn.cursor()
        
        # Obtener todas las tablas de pedidos pendientes
        cur.execute("""
            SELECT table_name FROM information_schema.tables 
            WHERE table_name LIKE 'pedidos_pendientes_%' AND table_schema = 'public'
        """)
        tablas = [row[0] for row in cur.fetchall()]
        
        # También incluir lista_espera si existe
        cur.execute("""
            SELECT table_name FROM information_schema.tables 
            WHERE table_name = 'lista_espera' AND table_schema = 'public'
        """)
        if cur.fetchone():
            tablas.append('lista_espera')
        
        # Si tenemos ID de factura, generar ID secuencial basado en la factura
        if id_factura:
            # Obtener el prefijo base para esta factura
            id_base = f"F{base_numero:03d}"
            
            # Buscar cuántos productos ya existen para esta factura
            total_productos_factura = 0
            for tabla in tablas:
                cur.execute(f"""
                    SELECT COUNT(*) FROM {tabla}
                    WHERE id_factura = %s
                """, (id_factura,))
                
                count = cur.fetchone()[0]
                total_productos_factura += count
            
            # Generar el siguiente número secuencial para esta factura
            siguiente_secuencia = total_productos_factura + 1
            
            # Intentar generar IDs secuenciales hasta encontrar uno libre
            max_intentos = 10
            for intento in range(max_intentos):
                # Formato: F001A, F001B, F001C... (base + letra secuencial)
                if siguiente_secuencia <= 26:
                    sufijo = chr(64 + siguiente_secuencia)  # A=65, B=66, etc.
                else:
                    # Si hay más de 26 productos, usar números
                    sufijo = str(siguiente_secuencia - 26).zfill(2)
                
                id_propuesto = f"{id_base}{sufijo}"
                
                # Verificar que no exceda 5 caracteres
                if len(id_propuesto) > 5:
                    # Reducir el número base si es muy largo
                    base_reducido = base_numero % 100  # Solo 2 dígitos
                    id_base = f"F{base_reducido:02d}"
                    id_propuesto = f"{id_base}{sufijo}"
                
                # Verificar si ya existe este ID
                existe = False
                for tabla in tablas:
                    cur.execute(f"""
                        SELECT COUNT(*) FROM {tabla}
                        WHERE id_orden_profesional = %s
                    """, (id_propuesto,))
                    
                    count = cur.fetchone()[0]
                    if count > 0:
                        existe = True
                        break
                
                if not existe:
                    cur.close()
                    conn.close()
                    print(f"✅ Nuevo ID generado para factura {id_factura}: {id_propuesto}")
                    return id_propuesto
                
                # Ya existe, probar con el siguiente
                siguiente_secuencia += 1
        
        # Fallback: usar el método anterior (sin factura o si hay conflicto)
        patron_busqueda = f"{inicial_dia}%"
        ultimo_numero_encontrado = 0
        
        # Buscar el mayor número usado en todas las tablas
        for tabla in tablas:
            cur.execute(f"""
                SELECT id_orden_profesional FROM {tabla}
                WHERE id_orden_profesional LIKE %s
                AND LENGTH(id_orden_profesional) <= 5
                ORDER BY CAST(SUBSTRING(id_orden_profesional FROM '\\d+') AS INTEGER) DESC 
                LIMIT 1
            """, (patron_busqueda,))
            
            resultado = cur.fetchone()
            if resultado and resultado[0]:
                try:
                    import re
                    numeros = re.findall(r'\d+', resultado[0])
                    if numeros:
                        numero = int(numeros[0])
                        ultimo_numero_encontrado = max(ultimo_numero_encontrado, numero)
                except (ValueError, IndexError):
                    pass
        
        if ultimo_numero_encontrado > 0:
            siguiente_numero = ultimo_numero_encontrado + 1
        else:
            siguiente_numero = 1001  # Empezar desde 1001 para tener 4 dígitos
        
        # Intentar generar un ID único ultra corto (máximo 5 caracteres)
        max_intentos = 50
        for intento in range(max_intentos):
            # Formato: L1234 (1 letra + 4 números = 5 caracteres máximo)
            id_propuesto = f"{inicial_dia}{siguiente_numero}"
            
            # Verificar que no exceda 5 caracteres
            if len(id_propuesto) > 5:
                siguiente_numero = 1001  # Reset si se hace muy largo
                id_propuesto = f"{inicial_dia}{siguiente_numero}"
            
            # Verificar si ya existe en cualquier tabla
            existe = False
            for tabla in tablas:
                cur.execute(f"""
                    SELECT COUNT(*) FROM {tabla}
                    WHERE id_orden_profesional = %s
                """, (id_propuesto,))
                
                count = cur.fetchone()[0]
                if count > 0:
                    existe = True
                    break
            
            if not existe:
                # No existe, podemos usarlo
                cur.close()
                conn.close()
                return id_propuesto
            
            # Ya existe, incrementar
            siguiente_numero += 1
        
        # Si llegamos aquí, usar timestamp como fallback ultra corto
        cur.close()
        conn.close()
        timestamp = int(time.time())
        return f"{inicial_dia}{timestamp % 9999}"
        
    except Exception as e:
        debug_log(f"Error generando ID profesional: {e}")
        # Fallback con timestamp ultra corto (máximo 5 caracteres)
        timestamp = int(time.time())
        return f"{inicial_dia}{timestamp % 9999}"

def agregar_a_lista_espera(codigo, producto, terminacion, id_factura, prioridad, cantidad, base=None, presentacion=None, ubicacion=None):
    """Agrega el pedido a la tabla específica de la sucursal del usuario"""
    try:
        # Detectar sucursal automáticamente basándose en el USERNAME (no ID numérico)
        usuario_para_deteccion = USUARIO_USERNAME if 'USUARIO_USERNAME' in globals() and USUARIO_USERNAME else USUARIO_ID
        sucursal = obtener_sucursal_usuario(usuario_para_deteccion)
        tabla_pendientes = f'pedidos_pendientes_{sucursal}'

        # Normalizar prioridad
        prioridad_limpia = (prioridad or 'Media').strip().title()
        if prioridad_limpia not in ['Alta', 'Media', 'Baja']:
            prioridad_limpia = 'Media'

        debug_log(f"🏢 Usuario: {usuario_para_deteccion} → Sucursal: {sucursal} → Tabla: {tabla_pendientes} → Prioridad: {prioridad_limpia}")

        # Preparar datos previos
        base_a_usar = base if base is not None else base_var.get()
        presentacion_a_usar = presentacion if presentacion is not None else presentacion_var.get()
        ubicacion_a_usar = ubicacion if ubicacion is not None else ubicacion_var.get()

        # Calcular código base completo (con sufijo presentación si aplica)
        codigo_base_calculado = ""
        if base_a_usar and producto and terminacion:
            codigo_base_calculado = obtener_codigo_base(base_a_usar, producto, terminacion)
            if presentacion_a_usar and codigo_base_calculado not in ["No encontrado", "No Aplica"]:
                sufijo = obtener_sufijo_presentacion(presentacion_a_usar)
                if sufijo:
                    codigo_base_calculado += sufijo

        # Generar ID profesional único usando la función actualizada (basado en factura)
        id_profesional = generar_id_profesional(id_factura)

        # Calcular tiempo de compromiso
        tiempo_esta_orden = calcular_tiempo_compromiso(producto, cantidad)
        _ = calcular_fecha_compromiso(tiempo_esta_orden)

        # INTENTO CON REINTENTOS Y VERIFICACIÓN
        backoffs = [0.15, 0.4, 0.9]
        last_error = None

        for intento, espera in enumerate(backoffs, start=1):
            conn = None
            try:
                conn = psycopg2.connect(
                    host="dpg-d1b18u8dl3ps73e68v1g-a.oregon-postgres.render.com",
                    port=5432,
                    database="labels_app_db",
                    user="admin",
                    password="KCFjzM4KYzSQx63ArufESIXq03EFXHz3",
                    sslmode="require"
                )
                cur = conn.cursor()

                # verificar si ya existe mismo código+factura+presentación (permitimos duplicado si cambia presentación)
                cur.execute(f"""
                    SELECT id FROM {tabla_pendientes}
                    WHERE id_factura = %s AND codigo = %s AND COALESCE(presentacion,'') = %s
                """, (id_factura, codigo, presentacion_a_usar or ""))
                if cur.fetchone():
                    try:
                        cur.close()
                        conn.close()
                    except Exception:
                        pass
                    debug_log(f"ℹ️ Ya existía en {tabla_pendientes} con misma presentación ({presentacion_a_usar}); se considera éxito.")
                    return True

                # descubrir columnas disponibles de la tabla
                cur.execute("""
                    SELECT column_name FROM information_schema.columns
                    WHERE table_schema='public' AND table_name=%s
                """, (tabla_pendientes,))
                cols_disponibles = {r[0] for r in cur.fetchall()}

                # columnas base (siempre esperadas)
                cols = [
                    'id_orden_profesional','codigo','producto','terminacion','id_factura',
                    'prioridad','cantidad','tiempo_estimado','base','ubicacion','sucursal'
                ]
                vals = [
                    id_profesional, codigo, producto, terminacion, id_factura,
                    prioridad_limpia, cantidad, tiempo_esta_orden, base_a_usar, ubicacion_a_usar, sucursal.title()
                ]

                # opcionales
                if 'presentacion' in cols_disponibles:
                    cols.append('presentacion'); vals.append(presentacion_a_usar)
                if 'codigo_base' in cols_disponibles:
                    cols.append('codigo_base'); vals.append(codigo_base_calculado)
                if 'estado' in cols_disponibles:
                    cols.append('estado'); vals.append('Pendiente')

                placeholders = ", ".join(["%s"] * len(cols))
                sql = f"INSERT INTO {tabla_pendientes} (" + ", ".join(cols) + ") VALUES (" + placeholders + ")"

                cur.execute(sql, tuple(vals))
                conn.commit()

                # verificación inmediata de inserción
                cur.execute(f"""
                    SELECT id_orden_profesional FROM {tabla_pendientes}
                    WHERE id_orden_profesional = %s
                    OR (id_factura=%s AND codigo=%s AND COALESCE(presentacion,'')=%s)
                    LIMIT 1
                """, (id_profesional, id_factura, codigo, presentacion_a_usar or ""))
                ok = cur.fetchone()
                cur.close()
                conn.close()
                if ok:
                    debug_log(f"✅ Insert verificado en {tabla_pendientes}: {id_profesional} [{prioridad_limpia}]")
                    return True
                else:
                    last_error = Exception("Verificación post-inserción no encontró el registro")
                    debug_log(f"⚠️ No se encontró tras inserción; reintento {intento}")
                    time.sleep(espera)
                    continue

            except Exception as e:
                last_error = e
                try:
                    if conn:
                        conn.rollback()
                        conn.close()
                except Exception:
                    pass
                debug_log(f"❌ Error intento {intento} insertando en {tabla_pendientes}: {e}")
                time.sleep(espera)
                continue

        # Si llegamos aquí, todos los intentos fallaron
        debug_log(f"🚫 Falló el envío tras varios intentos para factura {id_factura}, código {codigo} ({prioridad_limpia}). Último error: {last_error}")
        return False

    except Exception as e:
        debug_log(f"❌ Error general en agregar_a_lista_espera: {e}")
        return False

def registrar_impresion(codigo, descripcion, producto, terminacion, presentacion, cantidad, sucursal, usuario_id='sistema', id_factura='', prioridad='Media'):
    """Registra una impresión en la base de datos para estadísticas"""
    try:
        conn = psycopg2.connect(
            host="dpg-d1b18u8dl3ps73e68v1g-a.oregon-postgres.render.com",
            port=5432,
            database="labels_app_db",
            user="admin",
            password="KCFjzM4KYzSQx63ArufESIXq03EFXHz3",
            sslmode="require"
        )
        cur = conn.cursor()
        cur.execute("""
            INSERT INTO historial_impresiones 
            (codigo, descripcion, producto, terminacion, presentacion, cantidad, sucursal, usuario_id, id_factura, prioridad)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s);
        """, (codigo, descripcion, producto, terminacion, presentacion, cantidad, sucursal, usuario_id, id_factura, prioridad))
        conn.commit()
        cur.close()
        conn.close()
        return True
    except Exception as e:
        return False

# === FUNCIONES PARA LISTA DE PRODUCTOS DE FACTURA ===

def agregar_producto_a_lista():
    """Agrega el producto actual a la lista temporal de la factura"""
    global lista_productos_factura
    
    # Validar campos necesarios
    c = codigo_entry.get().strip()
    d = descripcion_var.get().strip()
    p = producto_var.get().strip()
    t = terminacion_var.get().strip()
    pr = presentacion_var.get().strip()
    q = spin.get()
    
    if not c:
        messagebox.showwarning("Campo requerido", "Debe ingresar un código.")
        return
    
    if not p:
        messagebox.showwarning("Campo requerido", "Debe seleccionar un producto.")
        return
    
    if not pr:
        messagebox.showwarning("Presentación requerida", 
                              "Debe seleccionar una presentación antes de agregar el producto a la lista.\n\n" +
                              "Las presentaciones disponibles son:\n" +
                              "• 1/8 (para lacas)\n" +
                              "• Cuarto\n" +
                              "• Medio Galón\n" +
                              "• Galón\n" +
                              "• Cubeta")
        return
    
    # Crear producto para la lista
    producto_item = {
        'codigo': c,
        'descripcion': d,
        'producto': p,
        'terminacion': t,
        'presentacion': pr,
        'cantidad': q,
        'base': base_var.get(),
        'ubicacion': ubicacion_var.get()
    }
    
    # Agregar a la lista directamente (sin validación de duplicados)
    
    # Agregar a la lista
    lista_productos_factura.append(producto_item)
    
    # Mostrar confirmación
    messagebox.showinfo("Producto agregado", f"Código {c} agregado a la lista.\nTotal productos: {len(lista_productos_factura)}")
    
    # Limpiar campos para el siguiente producto
    limpiar_campos()

def abrir_gestor_lista_factura():
    """Abre la ventana para gestionar la lista de productos de la factura"""
    global lista_productos_factura
    
    if not lista_productos_factura:
        messagebox.showinfo("Lista vacía", "No hay productos en la lista.\nAgregue productos usando 'Agregar a Lista'.")
        return
    
    # Crear ventana
    ventana_lista = tk.Toplevel(app)
    ventana_lista.title("Gestionar Lista de Factura")
    ventana_lista.geometry("900x600")
    ventana_lista.resizable(True, True)
    ventana_lista.iconbitmap(ICONO_PATH)
    ventana_lista.grab_set()
    ventana_lista.transient(app)
    
    # Centrar ventana
    x = (ventana_lista.winfo_screenwidth() // 2) - (900 // 2)
    y = (ventana_lista.winfo_screenheight() // 2) - (600 // 2)
    ventana_lista.geometry(f"900x600+{x}+{y}")
    
    # Título
    ttk.Label(ventana_lista, text="Lista de Productos para Factura", 
             font=("Segoe UI", 14, "bold")).pack(pady=10)
    
    # Frame para información de factura
    frame_factura = ttk.LabelFrame(ventana_lista, text="Información de Factura", padding=10)
    frame_factura.pack(fill="x", padx=20, pady=5)
    
    # Variables para factura
    id_factura_var = tk.StringVar()
    prioridad_var = tk.StringVar(value="Media")
    
    # Campos de factura
    ttk.Label(frame_factura, text="ID Factura:").grid(row=0, column=0, sticky="w", padx=5)
    entry_factura = ttk.Entry(frame_factura, textvariable=id_factura_var, width=25)
    entry_factura.grid(row=0, column=1, padx=5, sticky="w")
    
    ttk.Label(frame_factura, text="Prioridad:").grid(row=0, column=2, sticky="w", padx=5)
    combo_prioridad = ttk.Combobox(frame_factura, textvariable=prioridad_var, 
                                  values=["Alta", "Media", "Baja"], state="readonly", width=15)
    combo_prioridad.grid(row=0, column=3, padx=5, sticky="w")
    
    # Frame para lista de productos
    frame_lista = ttk.LabelFrame(ventana_lista, text="Productos en la Lista", padding=10)
    frame_lista.pack(fill="both", expand=True, padx=20, pady=10)
    
    # Treeview para mostrar productos
    columns = ("Código", "Descripción", "Producto", "Terminación", "Presentación", "Cantidad")
    tree_productos = ttk.Treeview(frame_lista, columns=columns, show="headings", height=12)
    
    # Configurar columnas
    for col in columns:
        tree_productos.heading(col, text=col)
        if col == "Código":
            tree_productos.column(col, width=100)
        elif col == "Descripción":
            tree_productos.column(col, width=200)
        elif col == "Cantidad":
            tree_productos.column(col, width=80)
        else:
            tree_productos.column(col, width=120)
    
    # Scrollbar
    scrollbar = ttk.Scrollbar(frame_lista, orient="vertical", command=tree_productos.yview)
    tree_productos.configure(yscrollcommand=scrollbar.set)
    
    tree_productos.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    
    def cargar_productos_en_tree():
        """Carga los productos de la lista en el treeview"""
        for item in tree_productos.get_children():
            tree_productos.delete(item)
        
        for producto in lista_productos_factura:
            tree_productos.insert("", "end", values=(
                producto['codigo'],
                producto['descripcion'],
                producto['producto'],
                producto['terminacion'],
                producto['presentacion'],
                producto['cantidad']
            ))
    
    def eliminar_producto():
        """Elimina el producto seleccionado de la lista"""
        global lista_productos_factura
        selection = tree_productos.selection()
        if not selection:
            messagebox.showwarning("Selección", "Seleccione un producto para eliminar.")
            return
        
        item = tree_productos.item(selection[0])
        vals = item['values']
        codigo_sel = vals[0]
        desc_sel = vals[1]
        prod_sel = vals[2]
        term_sel = vals[3]
        pres_sel = vals[4]
        cant_sel = vals[5]
        
        # Confirmar eliminación
        if messagebox.askyesno("Confirmar", f"¿Eliminar el producto {codigo_sel} ({pres_sel})?"):
            # Eliminar solo la primera coincidencia exacta (permitiendo duplicados restantes)
            idx_to_remove = None
            for i, p in enumerate(lista_productos_factura):
                if (
                    p.get('codigo') == codigo_sel and
                    p.get('presentacion') == pres_sel and
                    p.get('terminacion') == term_sel and
                    p.get('descripcion') == desc_sel and
                    p.get('producto') == prod_sel and
                    str(p.get('cantidad')) == str(cant_sel)
                ):
                    idx_to_remove = i
                    break
            if idx_to_remove is None:
                # Fallback: coincidir por código + presentación + terminación
                for i, p in enumerate(lista_productos_factura):
                    if (
                        p.get('codigo') == codigo_sel and
                        p.get('presentacion') == pres_sel and
                        p.get('terminacion') == term_sel
                    ):
                        idx_to_remove = i
                        break
            if idx_to_remove is not None:
                del lista_productos_factura[idx_to_remove]
            # Recargar tree
            cargar_productos_en_tree()
            messagebox.showinfo("Eliminado", f"Producto {codigo_sel} ({pres_sel}) eliminado.")
    
    def enviar_todos_a_lista_espera():
        """Envía todos los productos a la lista de espera"""
        global lista_productos_factura
        
        # Validar ID de factura silenciosamente
        id_factura = id_factura_var.get().strip()
        if not id_factura:
            return  # Salir silenciosamente
        
        prioridad = prioridad_var.get()
        
        # Envío directo sin confirmación para velocidad
        
        # Operaciones en paralelo para máxima velocidad
        import threading
        def procesar_lista_bd():
            try:
                for producto in lista_productos_factura:
                    try:
                        # Registrar impresión (para estadísticas)
                        registrar_impresion(
                            producto['codigo'], producto['descripcion'], producto['producto'],
                            producto['terminacion'], producto['presentacion'], producto['cantidad'],
                            SUCURSAL, USUARIO_ID, id_factura, prioridad
                        )
                        
                        # Agregar a lista de espera con todos los datos necesarios
                        agregar_a_lista_espera(
                            producto['codigo'], producto['producto'], producto['terminacion'],
                            id_factura, prioridad, producto['cantidad'],
                            base=producto['base'], 
                            presentacion=producto['presentacion'],
                            ubicacion=producto['ubicacion']
                        )
                    except:
                        pass  # Continuar con otros productos
            except:
                pass
        
        threading.Thread(target=procesar_lista_bd, daemon=True).start()
        
        # Limpiar y cerrar sin mensajes para velocidad
        lista_productos_factura = []
        ventana_lista.destroy()
    
    # Cargar productos iniciales
    cargar_productos_en_tree()
    
    # Frame para botones
    frame_botones = ttk.Frame(ventana_lista)
    frame_botones.pack(fill="x", padx=20, pady=15)
    
    ttk.Button(frame_botones, text="❌ Eliminar Seleccionado", 
              command=eliminar_producto, bootstyle="danger").pack(side="left", padx=8, pady=5)
    
    ttk.Button(frame_botones, text="🔄 Limpiar Lista", 
              command=lambda: [setattr(globals(), 'lista_productos_factura', []), 
                              cargar_productos_en_tree(),
                              messagebox.showinfo("Lista limpiada", "Todos los productos han sido eliminados.")], 
              bootstyle="warning").pack(side="left", padx=8, pady=5)
    
    ttk.Button(frame_botones, text="📋 Enviar Todos a Lista de Espera", 
              command=enviar_todos_a_lista_espera, 
              bootstyle="success").pack(side="right", padx=8, pady=5)
    
    ttk.Button(frame_botones, text="❌ Cerrar", 
              command=ventana_lista.destroy, 
              bootstyle="secondary").pack(side="right", padx=8, pady=5)

# === Código Base desde tabla CodigoBase ===
def obtener_codigo_base( base, producto, terminacion):
    try:
        conn = psycopg2.connect(
            host="dpg-d1b18u8dl3ps73e68v1g-a.oregon-postgres.render.com",
            port=5432,
            database="labels_app_db",
            user="admin",
            password="KCFjzM4KYzSQx63ArufESIXq03EFXHz3",
            sslmode="require"
        )
        cur = conn.cursor()
        cur.execute("SELECT base, tath, tath2, tath3 , flat, satin, sgi , flat2, satin3, sg4, satinkq, flatkp, flatmp, flatcov, flatpas, satinem, sgem, flatsp, satinsp, glossp, flatap, satinap, satinsan   FROM CodigoBase WHERE base ILIKE %s", (base,))
        row = cur.fetchone()
        conn.close()

        if not row:
            return "No encontrado"

        _, tath, tath2, tath3 ,flat, satin, sgi , flat2, satin3, sg4, satinkq, flatkp, flatmp, flatcov, flatpas, satinem, sgem, flatsp, satinsp, glossp, flatap, satinap, satinsan = row
        producto = producto.lower()
        terminacion = terminacion.lower()

        es_esmalte = any(p in producto for p in ["esmalte multiuso"])
        es_kempro = any(p in producto for p in ["kem pro"])
        es_kemaqua = any(p in producto for p in ["kem aqua"])
        es_masterpaint = any(p in producto for p in ["master paint"])
        es_pastel = any(p in producto for p in ["excello pastel"])
        es_emerald = any(p in producto for p in ["emerald"])
        es_superpaint = any(p in producto for p in ["super paint"])
        es_superpaintAP = any(p in producto for p in ["airpurtec"])
        es_sanitizing= any(p in producto for p in ["sanitizing"])
        es_laca= any(p in producto for p in ["laca"])
        es_EsmalteIndustrial= any(p in producto for p in ["esmalte kem"])
        es_uretano= any(p in producto for p in ["uretano"])
        es_tintealthinner= any(p in producto for p in ["tinte al thinner"])
        es_monocapa= any(p in producto for p in ["monocapa"])
        es_excellocov= any(p in producto for p in ["excello voc"])
        es_excellopremium= any(p in producto for p in ["excello premium"])
        es_waterblocking= any(p in producto for p in ["water blocking"])
        es_airpuretec= any(p in producto for p in ["airpuretec"])
        es_hcsiloconeacr= any(p in producto for p in ["h&c silicone-acrylic"])
        es_hcheavyshield= any(p in producto for p in ["h&c heavy-shield"]) 
        es_ProMarEgShel= any(p in producto for p in ["promar® 200 voc"])
        es_ProMarEgShel400= any(p in producto for p in ["promar® 400 voc"]) 
        es_proindustrialDTM= any(p in producto for p in ["pro industrial dtm"])                     
        es_armoseal= any(p in producto for p in ["armoseal 1000hs"])                     
        es_armosealtp= any(p in producto for p in ["armoseal t-p"]) 
        es_scufftuff= any(p in producto for p in ["scuff tuff-wb"]) 


        base_color = base_var.get().lower()


        if es_kemaqua:

            if terminacion == "satin":
                return satinkq
            else:
                return "No Aplica"

        if es_airpuretec:

            if  terminacion == "mate":

                if base_color == "extra white":
                 return "A86W00061-" 
                
                elif base_color == "deep":
                    return "A86W00063-"
                    
            elif terminacion == "satin":

                if base_color== "extra white":
                  
                  return "A87W00061-" 
                
                elif base_color == "deep":
                    return "A87W00063-"
            else:
                return "No Aplica"
                
        if es_waterblocking:

            if terminacion == "mate":
                return "LX12WDR50-"
            else:
                return "No Aplica"    
            
        if es_excellocov:

            if terminacion == "mate":

                return flatcov
            
            elif terminacion == "satin" and base_color == "extra white":

                return "A20DR2651-"
            
            else:
                return "No Aplica"    

        if es_laca:

            if terminacion == "mate":
                return "L15-" 
            elif terminacion == "semimate":
                return "L15-" 
            elif terminacion == "brillo":
                return "L15-" 
            else:
                return "No Aplica"
            

        if es_EsmalteIndustrial:

            if terminacion == "mate":
                return "F300-"
                
            elif terminacion == "semimate":
                return "F300-"
            
            elif terminacion == "brillo":
                return "F300-"
            else:
                return "No Aplica"

        if es_hcsiloconeacr:

             if terminacion == "mate":

                if base_color == "extra white":
                 return "20.101214-" 
                
                elif base_color == "deep":
                 return "20.102214-1" 

                elif base_color == "ultra deep":
                 return "20.103214-"                 
                
                else:
                    return "No aplica"
                                 
             else:
                return "No Aplica"
             
        if es_proindustrialDTM:

             if terminacion == "gloss":

                if base_color == "extra white":
                 return "B66W1051-" 
                
                elif base_color == "ultra deep":
                 return "B66T1054-" 
               
                
                else:
                    return "No aplica"
                                 
             else:
                return "No Aplica"

        if es_scufftuff:

             if terminacion == "mate":

                if base_color == "extra white":
                 return "S23W00051-" 
                
                elif base_color == "ultra deep":
                 return "S23T00154-" 
                
                elif base_color == "deep":
                 return "S23W00153-" 
                
                else:
                    return "No aplica"
                
             elif terminacion == "satin" :
                 return "S24W00051-" 
             
               
             elif terminacion == "semigloss" :
                 return "S26W00051-"                                                
             else:
                return "No Aplica"
             
        if es_hcheavyshield:

             if terminacion == "gloss":

                if base_color == "extra white":
                 return "35.100214-" 
                
                elif base_color == "deep":
                 return "35.100314-" 

                elif base_color == "ultra deep":
                 return "35.100414-"                 
                
                else:
                    return "No aplica"
                                 
             else:
                return "No Aplica"
             
        if  es_ProMarEgShel:

             if terminacion == "satin":

                if base_color == "deep":
                 return "B20W02653-" 
                
                elif base_color == "extra white":
                 return "B20W12651-"                
                
                else:
                    return "No aplica"
                                 
             elif terminacion== "mate":

                if base_color == "ultra deep":
                 return "B30T02654-" 
                
                elif base_color == "extra White":
                 return "B30W02651-1"                
                
                elif base_color == "deep":
                 return "B30W02653-"   

             elif terminacion== "semigloss":

                if base_color == "extra White":
                 return "B31W02651-"                
                                                  
                         
             else:

                return "No aplica"

        if  es_ProMarEgShel400:

             if terminacion == "satin":

                if base_color == "extra white":
                 return "B20W04651-" 
 
                else:
                    return "No aplica"

             else:
                  return "No Aplica"   
             
        if es_armoseal:

             if terminacion == "gloss":

                if base_color == "extra white":
                 return "B67W2001-" 
                
                elif base_color == "ultra deep":
                 return "B67T2004-" 
               
                
                else:
                    return "No aplica"
                                 
             else:
                return "No Aplica"

        if es_armosealtp:

             if terminacion == "semigloss":

                if base_color == "extra white":
                 return "B90T104-" 
                
                elif base_color == "ultra deep":
                 return "B90W111-" 
               
                
                else:
                    return "No aplica"
                                 
             else:
                return "No Aplica"



        if es_uretano:

             if terminacion == "mate":

                if base_color == "extra white":
                 return "ASPPA-" 
                
                elif base_color == "deep" or base_color == "ultra deep":
                 return "ASPPB-" 
                
                else:
                    return "ASPPD-"
                
             elif terminacion == "semimate":

                if base_color == "extra white":
                 return "ASPPA-" 
                
                elif base_color == "deep" or base_color == "ultra deep":
                 return "ASPPB-" 
                
                else:
                    return "ASPPD-"
                
             elif terminacion == "brillo":

                if base_color == "extra white":
                 return "ASPPA-" 
                
                elif base_color == "deep" or base_color == "ultra deep":
                 return "ASPPB-" 
                
                else:
                    return "ASPPD-"                    
             else:
                return "No Aplica"
    
        if es_tintealthinner:

            if terminacion == "claro":
                return tath 
            
            elif terminacion == "intermedio":
                return tath2
            
            elif terminacion== "especial":
                return tath3
            else:
                return "No Aplica"
            
        if es_monocapa:

            if terminacion == "mate":

                if base_color == "extra white":
                 return "ASMCA-" 
                
                elif base_color == "deep" or base_color == "ultra deep":
                 return "ASMCB-" 
                
                else:
                    return "ASMCD-" 
                
            elif terminacion == "semimate":

                if base_color == "extra white":
                 return "ASMCA-" 
                
                elif base_color == "deep" or base_color == "ultra deep":
                 return "ASMCB-" 
                
                else:
                    return "ASMCD-" 
                
            elif terminacion == "brillo":

                if base_color == "extra white":
                 return "ASMCA-" 
                
                elif base_color == "deep" or base_color == "ultra deep":
                 return "ASMCB-" 
                
                else:
                    return "ASMCD-" 
            else:
                return "No Aplica"

        if es_esmalte:

            if terminacion == "mate":
                return flat2
            elif terminacion == "satin":
                return satin3
            elif terminacion == "gloss":
                return sg4
            else:
                return "No Aplica"
            
        elif es_kempro:
                
            if terminacion == "mate":
                 return flatkp
            else:
                return "No Aplica"
            
        elif es_masterpaint:
                
            if terminacion == "mate":
                 return flatmp
            else:
                return "No Aplica"
            
        elif es_pastel:
                
            if terminacion == "mate":
                 return flatpas
            else:
                return "No Aplica"   
        elif es_emerald: 
                    
            if terminacion == "satin":
                return satinem + " K37W02751-"
            
            elif terminacion == "gloss":
                return sgem + "K38W02751-"
            else:
                return "No Aplica"
        elif es_superpaint: 
                    
            if terminacion == "mate":
                return flatsp
            
            elif terminacion == "satin":
                return satinsp    

            elif  terminacion == "gloss":
                return glossp
            else:
                return "No Aplica"           
            
        elif es_superpaintAP: 
                    
            if terminacion == "mate":
                return flatap
            elif terminacion == "satin":
                return satinap
            else:
                return "No Aplica" 
            
        elif es_sanitizing:
                
            if terminacion == "satin":
                 return satinsan
            else:
                return "No Aplica"
            
        elif es_excellopremium: 
            # Reglas especiales Excello Premium
            # 1) Ultra Deep II -> código PP4-
            es_ultra_deep_ii = any(k in base_color for k in ["ultra deep ii", "ultradeep ii", "ultra-deep ii", "ultra deep 2"])
            if es_ultra_deep_ii:
                return "PP4-"

            # 2) Ultra Deep (no II) -> solo Semisatin con código A27WDR03-
            es_ultra_deep = ("ultra deep" in base_color) and not es_ultra_deep_ii
            if es_ultra_deep:
                if terminacion == "semisatin":
                    return "A27WDR03-"
                else:
                    return "No Aplica"

            # 3) Resto de mapeos Excello Premium
            if terminacion == "mate":
                return flat
            elif terminacion == "satin":
                return satin
            elif terminacion == "semigloss":
                return sgi
            else:
                return "No Aplica"   
                                  
        else:
            return "No Aplica"
        
    except Exception as e:
        return "Error"
    
def imprimir_ficha_pintura(codigo_pintura):
    producto = producto_var.get().strip().lower()

    try:
        if producto == "excello premium":
            datos = obtener_datos_por_pintura(codigo_pintura)
            if datos:
                temp_pdf = tempfile.mktemp(".pdf")
                generar_pdf_ficha(datos, temp_pdf)
                os.startfile(temp_pdf)
                imprimir_pdf(temp_pdf)
            else:
                messagebox.showwarning("No encontrado", f"No hay fórmula disponible para el producto: {producto}")

        elif producto == "tinte al thinner":
            datos = obtener_datos_por_tinte(codigo_pintura)
            if datos:
                temp_pdf = tempfile.mktemp(".pdf")
                generar_pdf_tinte(datos, temp_pdf)
                os.startfile(temp_pdf)
                imprimir_pdf(temp_pdf)
            else:
                messagebox.showwarning("No encontrado", f"No hay fórmula disponible para el producto: {producto}")

        else:
            messagebox.showwarning("Producto no soportado", f"Producto no reconocido: {producto}")

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error al imprimir: {e}")

            

def mostrar_codigo_base():
    global codigo_base_actual

    base = base_var.get()
    producto = producto_var.get()
    terminacion = terminacion_var.get()
    presentacion = presentacion_var.get()

    if not base or not producto or not terminacion:
        aviso_var.set("Completa todos los campos")
        return

    # Obtener código base
    resultado = obtener_codigo_base(base, producto, terminacion)
    
    # Agregar sufijo de presentación si está seleccionada
    if presentacion and resultado != "No encontrado" and resultado != "No Aplica":
        sufijo_presentacion = obtener_sufijo_presentacion(presentacion)
        if sufijo_presentacion:
            resultado += sufijo_presentacion

    app.clipboard_clear()
    app.clipboard_append(resultado)
    aviso_var.set("Código facturación copiado en el portapapeles")

    # Guardamos solo para vista previa
    codigo_base_actual = resultado
    codigo_base_var.set(resultado)

    actualizar_vista()
    limpiar_mensaje_despues(3000)

def actualizar_terminaciones(*args):
    """Actualiza las terminaciones disponibles según el producto seleccionado"""
    producto = producto_var.get().lower()
    base = base_var.get().lower()
    
    # Buscar terminaciones válidas para el producto
    terminaciones_validas = []
    for key, terminaciones in TERMINACIONES_POR_PRODUCTO.items():
        if key in producto:
            terminaciones_validas = terminaciones
            break
    
    # Si no se encuentra el producto, usar todas las terminaciones
    if not terminaciones_validas:
        terminaciones_validas = ['Mate', 'Satin', 'Semigloss', 'Semimate', 'Gloss', 'Brillo', 
                                "N/A", "ESPECIAL", "CLARO", "INTERMEDIO", "MADERA", "PERLADO", "METALICO"]
    
    # Lógica especial para Excello Premium con bases Ultra Deep / Ultra Deep II
    if 'excello premium' in producto:
        es_ultra_deep_ii = any(k in base for k in ['ultra deep ii', 'ultradeep ii', 'ultra-deep ii', 'ultra deep 2'])
        es_ultra_deep = ('ultra deep' in base)
        if es_ultra_deep or es_ultra_deep_ii:
            # Solo permitir Semisatin en ambas bases
            terminaciones_validas = ['Semisatin']

    # Actualizar el combobox
    terminaciones_combobox['values'] = terminaciones_validas
    
    # Limpiar la selección actual si no es válida
    terminacion_actual = terminacion_var.get()
    if terminacion_actual and terminacion_actual not in terminaciones_validas:
        terminacion_var.set('')
    
    # Si solo hay una terminación válida, seleccionarla automáticamente
    if len(terminaciones_validas) == 1:
        terminacion_var.set(terminaciones_validas[0])
    
    # Actualizar vista previa
    actualizar_vista()    


# Agrupa los botones en un LabelFrame moderno


# Crear estilo personalizado para botones y LabelFrame
style = ttk.Style()
# Botones azul oscuro, texto/iconos blancos
style.configure("BotonGrande.TButton", font=("Segoe UI", 11, "bold"), foreground="#fff", background="#222c3c", padding=8, borderwidth=1)
# Botón Imprimir con azul menos claro
style.configure("BotonImprimir.TButton", font=("Segoe UI", 11, "bold"), foreground="#fff", background="#1976d2", padding=8, borderwidth=1)
# Botones especiales para lista (verde)
style.configure("BotonEspecial.TButton", font=("Segoe UI", 10, "bold"), foreground="#fff", background="#28a745", padding=6, borderwidth=1)
# LabelFrame y título fondo blanco, título azul oscuro
style.configure("Acciones.TLabelframe", background="#fff", borderwidth=2, relief="groove")
style.configure("Acciones.TLabelframe.Label", font=("Segoe UI", 12, "bold"), foreground="#222c3c", background="#fff")

frame_botones = ttk.LabelFrame(app, text="Acciones", padding=16, style="Acciones.TLabelframe")
frame_botones.place(x=420, y=295, width=440, height=220)  # Subido un poco más nuevamente (de 310 a 295)

btn_style = {"width": 12, "style": "BotonGrande.TButton"}  # Reducir width para 4 botones


btn_limpiar = ttk.Button(frame_botones, text="Limpiar", command=limpiar_campos, **btn_style)
btn_limpiar.grid(row=0, column=0, padx=8, pady=10, sticky='ew')

btn_codigo = ttk.Button(frame_botones, text="Código", command=mostrar_codigo_base, **btn_style)
btn_codigo.grid(row=0, column=1, padx=8, pady=10, sticky='ew')

btn_ficha = ttk.Button(frame_botones, text="Fórmula", command=on_btn_imprimir_click, **btn_style)
btn_ficha.grid(row=0, column=2, padx=8, pady=10, sticky='ew')

btn_personalizar = ttk.Button(frame_botones, text="Custom", command=abrir_ventana_personalizados, **btn_style)
btn_personalizar.grid(row=0, column=3, padx=8, pady=10, sticky='ew')

# Segunda fila de botones
btn_agregar_lista = ttk.Button(frame_botones, text="📋 Agregar a Lista", command=lambda: agregar_producto_a_lista(), style="BotonEspecial.TButton")
btn_agregar_lista.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky='ew')

btn_gestionar_lista = ttk.Button(frame_botones, text="📑 Gestionar Lista", command=lambda: abrir_gestor_lista_factura(), style="BotonEspecial.TButton")
btn_gestionar_lista.grid(row=1, column=2, columnspan=2, padx=5, pady=5, sticky='ew')

# Botón Imprimir ocupa 4 columnas y se centra debajo
btn_print = ttk.Button(frame_botones, text="� Enviar", command=imprimir_guardar, width=14, style="BotonImprimir.TButton")
btn_print.grid(row=2, column=0, columnspan=4, padx=10, pady=10, sticky='ew')

for i in range(4):  # Cambiar a 4 columnas
    frame_botones.columnconfigure(i, weight=1)

# Inicializar vista previa en blanco
actualizar_vista()

# Inicializar presentaciones
actualizar_presentaciones()

def main():
    """Función principal de la aplicación"""
    try:
        app.mainloop()
    except Exception as e:
        messagebox.showerror("Error", f"Error inesperado al ejecutar la aplicación:\n{e}")

if __name__ == "__main__":
    main()

