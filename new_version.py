import os
import re
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, Canvas
import pandas as pd
from datetime import datetime
from collections import defaultdict
import json
import sqlite3
from tkinter import filedialog
import winsound
import webbrowser
import requests # Necesitas instalar esta librería: pip install requests
import sys

# Intenta importar xlrd para archivos .xls
try:
    import xlrd
except ImportError:
    messagebox.showwarning("Dependencia Faltante", 
                           "La librería 'xlrd' no está instalada. Es necesaria para leer archivos .xls (Geometría). "
                           "Por favor, instala: pip install xlrd")
    xlrd = None

class VerificadorCablesMPO:
    def __init__(self):
        self.root = None
        self.ot_entry = None
        self.serie_entry = None
        self.resultado_text = None
        self.ruta_ilrl_label = None
        self.ruta_geo_label = None
        self.btn_ver_detalles = None
    
        self.ruta_base_ilrl = r"C:\Users\Paulo\Documents\DOCUMENTS EPCOMM\Proyectos de automatización\P-9. Verificación de datos MPO ILRL GEO\DATOS DE PRUEBA\ILRL_JWS1-1"
        self.ruta_base_geo = r"C:\Users\Paulo\Documents\DOCUMENTS EPCOMM\Proyectos de automatización\P-9. Verificación de datos MPO ILRL GEO\DATOS DE PRUEBA\Geometria_JWS1-1"
        self.ruta_base_polaridad = r"C:\Users\Paulo\Documents\DOCUMENTS EPCOMM\Proyectos de automatización\P-9. Verificación de datos MPO ILRL GEO\DATOS DE PRUEBA\Polaridad"

        self.config_file = "config.json"
        self.password = "admin123" 
        self.cable_config = {}  
        self.db_name = None
        self._init_db_path() 
    
        self.last_ilrl_analysis_data = None
        self.last_geo_analysis_data = None
        self.last_polaridad_analysis_data = None
        self.last_ilrl_file_path = None
        self.last_geo_file_path = None
        self.last_polaridad_file_path = None
    
        app_data_dir = os.path.join(os.path.expanduser('~'), 'Documents', 'VerificadorCablesData')
        self.db_name = os.path.join(app_data_dir, "cable_verifications.db")
    
        os.makedirs(app_data_dir, exist_ok=True)
    
        self._init_database()
        self._init_ot_database() 
        self.cargar_rutas()

        self.item_data_cache = {}

        # --- Variables para la actualización remota ---
        self.LOCAL_VERSION = "1.0.0"
        self.VERSION_URL = "https://raw.githubusercontent.com/ZombPool/P-9.-Verificaci-n-de-datos-MPO-ILRL-GEO/refs/heads/main/version.txt"
        self.UPDATE_URL = "https://raw.githubusercontent.com/ZombPool/P-9.-Verificaci-n-de-datos-MPO-ILRL-GEO/refs/heads/main/new_version.py"
        # --- Fin de variables de actualización ---

    def create_main_window(self):
        """Crea la ventana principal de la aplicación."""
        self.root = tk.Tk()
        self.root.title("Sistema de Verificación de Cables MPO")
        self.root.geometry("800x600")
        self.root.resizable(True, True) 
        self.root.config(bg='#F0F0F0') 

        self.verificar_actualizaciones()  # Llamada a la función de verificación al iniciar

        style = ttk.Style()
        style.theme_use('default') 
        
        BG_COLOR = "#F7F7F7" 
        ACCENT_BLUE = "#007BFF" 
        HOVER_BLUE = "#0056b3" 
        TEXT_COLOR = "#333333" 
        LIGHT_TEXT_COLOR = "#6C757D" 

        style.configure("TFrame", background=BG_COLOR)
        style.configure("TLabelframe", background=BG_COLOR, relief="flat", borderwidth=1, 
                        focusthickness=0, highlightthickness=0)
        style.configure("TLabelframe.Label", background=BG_COLOR, foreground=ACCENT_BLUE, 
                        font=("Arial", 11, "bold"))

        style.configure("TLabel", background=BG_COLOR, foreground=TEXT_COLOR, font=("Arial", 10))
        style.configure("TEntry", font=("Arial", 10), padding=5, fieldbackground="#FFFFFF", 
                        foreground=TEXT_COLOR, borderwidth=1, relief="solid")
        style.configure("TButton", font=("Arial", 10, "bold"), padding=10, 
                                  background=ACCENT_BLUE, foreground="white", 
                                  relief="flat", borderwidth=0, borderradius=5)
        style.map("TButton", 
                  background=[('active', HOVER_BLUE)], 
                  foreground=[('active', 'white')])
        style.configure("Vertical.TScrollbar", troughcolor=BG_COLOR, background=ACCENT_BLUE, 
                        gripcount=0, relief="flat", borderwidth=0)
        style.configure("Horizontal.TScrollbar", troughcolor=BG_COLOR, background=ACCENT_BLUE, 
                        gripcount=0, relief="flat", borderwidth=0)
        style.configure("Treeview", font=("Arial", 9), rowheight=25, background="#FFFFFF", 
                        fieldbackground="#FFFFFF", foreground=TEXT_COLOR)
        style.map('Treeview', background=[('selected', ACCENT_BLUE)]) 
        style.configure("Treeview.Heading", font=("Arial", 10, "bold"), background=ACCENT_BLUE, 
                        foreground="white", relief="flat", borderwidth=0)
        
        main_frame = ttk.Frame(self.root, style="TFrame")
        main_frame.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(main_frame, highlightthickness=0, bg=BG_COLOR)
        xscrollbar = ttk.Scrollbar(main_frame, orient="horizontal", command=canvas.xview, style="Horizontal.TScrollbar")
        yscrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview, style="Vertical.TScrollbar")
        
        scrollable_content_frame = ttk.Frame(canvas, style="TFrame")

        scrollable_window_id = canvas.create_window((0, 0), window=scrollable_content_frame, anchor="nw")
        
        canvas.configure(yscrollcommand=yscrollbar.set, xscrollcommand=xscrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        yscrollbar.pack(side="right", fill="y")
        xscrollbar.pack(side="bottom", fill="x")

        def _on_canvas_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas_width = event.width
            if yscrollbar.winfo_ismapped(): 
                canvas.itemconfigure(scrollable_window_id, width=canvas_width)
            else:
                canvas.itemconfigure(scrollable_window_id, width=canvas_width)
            
        canvas.bind("<Configure>", _on_canvas_configure)

        self.root.bind_all("<MouseWheel>", self._on_mouse_wheel)
        
        self.root.bind('<F11>', self.toggle_fullscreen)
        self.root.bind('<Escape>', self.exit_fullscreen)

        input_frame = ttk.LabelFrame(scrollable_content_frame, text=" Datos del Cable ", padding=(15, 10))
        input_frame.grid(row=0, column=0, columnspan=2, padx=20, pady=15, sticky="ew")
        
        input_frame.columnconfigure(1, weight=1)

        ttk.Label(input_frame, text="Número de OT (ej. 250700007):", font=("Arial", 10)).grid(row=0, column=0, sticky=tk.W, pady=5)
        self.ot_entry = ttk.Entry(input_frame, width=30)
        self.ot_entry.grid(row=0, column=1, pady=5, padx=10, sticky="ew")
        self.ot_entry.bind("<Return>", self.verificar_cable_automatico)
        
        ttk.Label(input_frame, text="Número de Serie (ej. 2507000070013):", font=("Arial", 10)).grid(row=1, column=0, sticky=tk.W, pady=5)
        self.serie_entry = ttk.Entry(input_frame, width=30)
        self.serie_entry.grid(row=1, column=1, pady=5, padx=10, sticky="ew")
        self.serie_entry.bind("<KeyRelease>", self.verificar_cable_automatico)
        self.serie_entry.bind("<Return>", self.verificar_cable_automatico)

        button_frame = ttk.Frame(scrollable_content_frame, style="TFrame")
        button_frame.grid(row=1, column=0, columnspan=2, pady=(15, 5))

        verify_button = ttk.Button(button_frame, 
                                   text="✅ Verificar Cable", 
                                   command=self.verificar_cable)
        verify_button.pack(side=tk.LEFT, padx=5, ipadx=10, ipady=5)
        
        menu_button = ttk.Menubutton(button_frame, text="⚙️ Opciones", direction="below")
        menu_button.pack(side=tk.LEFT, padx=5, ipadx=10, ipady=5)

        opciones_menu = tk.Menu(menu_button, tearoff=False)
        opciones_menu.add_command(label="📋 Ver detalles de O.T.", command=self.mostrar_detalles_ot_actual)
        opciones_menu.add_command(label="⚙️ Configurar O.T.", command=self.configurar_ot)
        opciones_menu.add_command(label="📖 Ver Registros", command=self.solicitar_contrasena_registros)
        opciones_menu.add_command(label="📁 Configurar Rutas", command=self.solicitar_contrasena)
        opciones_menu.add_command(label="📊 Verificar Ruta DB", command=self.verificar_ruta_db)
        
        menu_button.config(menu=opciones_menu)
        
        self.btn_ver_detalles = ttk.Button(button_frame, 
                                           text="🔍 Ver Detalles", 
                                           command=self.mostrar_detalles_totales,
                                           state=tk.DISABLED)
        self.btn_ver_detalles.pack(side=tk.LEFT, padx=5, ipadx=10, ipady=5)
        
        result_frame = ttk.LabelFrame(scrollable_content_frame, text=" Resultados de Verificación ", padding=(15, 10))
        result_frame.grid(row=2, column=0, columnspan=2, padx=20, pady=15, sticky="ew")
        result_frame.columnconfigure(0, weight=1)

        self.resultado_text = tk.Text(result_frame, wrap=tk.WORD, height=10, state=tk.DISABLED, 
                                    font=("Arial", 10), relief=tk.FLAT, background="#FFFFFF", foreground=TEXT_COLOR)
        self.resultado_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.resultado_text.tag_configure("normal", font=("Arial", 10), foreground=TEXT_COLOR)
        self.resultado_text.tag_configure("header", font=("Arial", 12, "bold"), foreground=ACCENT_BLUE)
        self.resultado_text.tag_configure("bold", font=("Arial", 10, "bold"), foreground=TEXT_COLOR)
        self.resultado_text.tag_configure("verde", foreground="#28A745") 
        self.resultado_text.tag_configure("rojo", foreground="#DC3545")
        self.resultado_text.tag_configure("orange", foreground="#FFC107")
        
        self.ruta_ilrl_label = ttk.Label(scrollable_content_frame, text=f"📂 Ruta ILRL: {self.ruta_base_ilrl}", font=("Arial", 9))
        self.ruta_ilrl_label.grid(row=3, column=0, columnspan=2, padx=20, pady=(5, 0), sticky="ew")

        self.ruta_geo_label = ttk.Label(scrollable_content_frame, text=f"📂 Ruta Geometría: {self.ruta_base_geo}", font=("Arial", 9))
        self.ruta_geo_label.grid(row=4, column=0, columnspan=2, padx=20, pady=(0, 15), sticky="ew")
        
        self.ruta_polaridad_label = ttk.Label(scrollable_content_frame, text=f"📂 Ruta Polaridad: {self.ruta_base_polaridad}", font=("Arial", 9))
        self.ruta_polaridad_label.grid(row=5, column=0, columnspan=2, padx=20, pady=(0, 15), sticky="ew")

        button_exit_frame = ttk.Frame(scrollable_content_frame, style="TFrame") 
        button_exit_frame.grid(row=6, column=0, columnspan=2, pady=(15, 5))
        
        exit_button = ttk.Button(button_exit_frame, 
                                 text="🚫 Salir del Programa", 
                                 command=self.root.destroy)
        exit_button.pack(pady=5, ipadx=10, ipady=5)
        
        footer_frame = ttk.Frame(scrollable_content_frame, style="TFrame") 
        footer_frame.grid(row=7, column=0, columnspan=2, pady=(10, 0))
        
        ttk.Label(footer_frame, 
                  text="Sistema de Verificación de Cables MPO v1.0", 
                  font=("Arial", 8), 
                  foreground=LIGHT_TEXT_COLOR,
                  ).pack(pady=5)

        self.root.mainloop()
        
    def verificar_actualizaciones(self):
        """
        Verifica si hay una nueva versión del programa disponible en GitHub.
        """
        try:
            print("Verificando actualizaciones...")
            response = requests.get(self.VERSION_URL)
            response.raise_for_status()  # Lanza un error si la petición falla
            version_remota = response.text.strip()

            if version_remota != self.LOCAL_VERSION:
                mensaje = f"¡Nueva versión disponible! Versión actual: {self.LOCAL_VERSION}. Nueva versión: {version_remota}. ¿Desea descargar e instalar la actualización?"
                if messagebox.askyesno("Actualización disponible", mensaje):
                    self.descargar_actualizacion()
            else:
                print("El programa está actualizado.")
        except requests.exceptions.RequestException as e:
            print(f"Error al verificar actualizaciones: {e}")
            messagebox.showwarning("Error de conexión", "No se pudo verificar si hay actualizaciones. Verifique su conexión a internet.")

    def descargar_actualizacion(self):
        """
        Descarga la nueva versión del programa y la ejecuta.
        """
        try:
            messagebox.showinfo("Descargando...", "La actualización se está descargando. Por favor, espere.")
            response = requests.get(self.UPDATE_URL)
            response.raise_for_status()
            
            # Guardar el nuevo archivo
            new_file_path = "VerificadorCables_NuevaVersion.py"
            with open(new_file_path, "wb") as f:
                f.write(response.content)

            messagebox.showinfo("Actualización completa", "La nueva versión se ha descargado. El programa se reiniciará para aplicar los cambios.")
            
            # Reiniciar el programa
            os.execv(sys.executable, ['python'] + [new_file_path])
        except requests.exceptions.RequestException as e:
            messagebox.showerror("Error de descarga", f"No se pudo descargar la actualización: {e}")
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error inesperado durante la actualización: {e}")

    def _on_mouse_wheel(self, event):
        """Maneja el evento de la rueda del ratón para el desplazamiento del canvas."""
        if self.root.winfo_exists():
            canvas_widget = event.widget
            if isinstance(canvas_widget, tk.Canvas):
                canvas_widget.yview_scroll(-1 * int((event.delta / 120)), "units")
            elif hasattr(canvas_widget, 'yview_scroll'): 
                canvas_widget.yview_scroll(-1 * int((event.delta / 120)), "units")
    
    def verificar_cable_automatico(self, event=None):
        """Método que se llama automáticamente al escribir en el campo de serie."""
        self.btn_ver_detalles.config(state=tk.DISABLED)
        self.last_ilrl_analysis_data = None
        self.last_geo_analysis_data = None
        
        serie_cable_raw = self.serie_entry.get().strip()
        ot_input_raw = self.ot_entry.get().strip()
        
        ot_part_from_serial = ""
        match_serial_ot = re.match(r'^(\d{9})\d{4}$', serie_cable_raw)
        if match_serial_ot:
            ot_part_from_serial = match_serial_ot.group(1)
        
        ot_input_cleaned_for_comp = ot_input_raw.upper().replace('JMO-', '')

        if ot_input_cleaned_for_comp and ot_part_from_serial and ot_input_cleaned_for_comp != ot_part_from_serial:
             self.resultado_text.config(state=tk.NORMAL)
             self.resultado_text.delete(1.0, tk.END)
             self.resultado_text.insert(tk.END, "⚠️ El número de OT y el de serie no coinciden. Verifique los datos.", "rojo")
             self.resultado_text.tag_unbind("ilrl_click", "<Button-1>")
             self.resultado_text.tag_unbind("geo_click", "<Button-1>")
             self.resultado_text.config(state=tk.DISABLED)
             return

        if re.match(r'^\d{13}$', serie_cable_raw):
            if not ot_input_raw or ot_input_cleaned_for_comp != ot_part_from_serial:
                self.ot_entry.delete(0, tk.END)
                self.ot_entry.insert(0, ot_part_from_serial)

            self.verificar_cable()
        elif len(serie_cable_raw) > 0: 
            self.resultado_text.config(state=tk.NORMAL)
            self.resultado_text.delete(1.0, tk.END)
            self.resultado_text.insert(tk.END, "Ingrese un número de serie de 13 dígitos para verificar (ej. 2507000070013).", "normal")
            self.resultado_text.tag_unbind("ilrl_click", "<Button-1>")
            self.resultado_text.tag_unbind("geo_click", "<Button-1>")
            self.resultado_text.config(state=tk.DISABLED)
        else: 
            self.resultado_text.config(state=tk.NORMAL)
            self.resultado_text.delete(1.0, tk.END)
            self.resultado_text.insert(tk.END, "Esperando un número de serie válido (JMO-2507000070013 o 2507000070013)...", "normal")
            self.resultado_text.tag_unbind("ilrl_click", "<Button-1>")
            self.resultado_text.tag_unbind("geo_click", "<Button-1>")
            self.resultado_text.config(state=tk.DISABLED)
    
    def mostrar_detalles_totales(self):
        """
        Muestra una ventana con los detalles de verificación para ILRL y Geometría en una sola vista.
        """
        if not self.last_ilrl_analysis_data and not self.last_geo_analysis_data and not self.last_polaridad_analysis_data:
            messagebox.showinfo("No hay datos", "No hay datos de verificación para mostrar. Realice una verificación primero.")
            return

        details_window = tk.Toplevel(self.root)
        details_window.title(f"Detalles de Verificación Completos")
        details_window.geometry("800x600")
        details_window.transient(self.root)
        details_window.grab_set()
        
        BG_COLOR = "#F7F7F7"
        ACCENT_BLUE = "#007BFF"
        TEXT_COLOR = "#333333"
        
        style = ttk.Style()
        style.configure("DetallesTotales.TFrame", background=BG_COLOR)
        style.configure("DetallesTotales.TLabel", background=BG_COLOR, foreground=TEXT_COLOR)

        main_frame = ttk.Frame(details_window, style="DetallesTotales.TFrame")
        main_frame.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(main_frame, highlightthickness=0, bg=BG_COLOR)
        yscrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        xscrollbar = ttk.Scrollbar(main_frame, orient="horizontal", command=canvas.xview)
        scrollable_frame = ttk.Frame(canvas, style="DetallesTotales.TFrame")
        
        scrollable_window_id = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        
        canvas.configure(yscrollcommand=yscrollbar.set, xscrollcommand=xscrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        yscrollbar.pack(side="right", fill="y")
        xscrollbar.pack(side="bottom", fill="x")

        def _on_details_canvas_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas_width = event.width
            canvas.itemconfigure(scrollable_window_id, width=canvas_width)

        canvas.bind("<Configure>", _on_details_canvas_configure)
        canvas.bind_all("<MouseWheel>", self._on_mouse_wheel)

        text_widget = tk.Text(scrollable_frame, wrap=tk.WORD, state=tk.DISABLED, 
                              font=("Arial", 10), relief=tk.FLAT, 
                              background="#FFFFFF", foreground=TEXT_COLOR)
        text_widget.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        text_widget.tag_configure("header", font=("Arial", 12, "bold"), foreground=ACCENT_BLUE)
        text_widget.tag_configure("subheader", font=("Arial", 11, "bold"), foreground=TEXT_COLOR)
        text_widget.tag_configure("normal", font=("Arial", 10), foreground=TEXT_COLOR)
        text_widget.tag_configure("bold", font=("Arial", 10, "bold"), foreground=TEXT_COLOR)
        text_widget.tag_configure("verde", foreground="#28A745")
        text_widget.tag_configure("rojo", foreground="#DC3545")
        text_widget.tag_configure("orange", foreground="#FFC107")

        text_widget.config(state=tk.NORMAL)
        text_widget.delete(1.0, tk.END)

        overall_status = "NO ENCONTRADO"
        
        ilrl_pass = self.last_ilrl_analysis_data and self.last_ilrl_analysis_data.get('estado') == "APROBADO"
        geo_pass = self.last_geo_analysis_data and self.last_geo_analysis_data.get('estado') == "APROBADO"
        polaridad_pass = self.last_polaridad_analysis_data and self.last_polaridad_analysis_data.get('status') == "PASS"

        if ilrl_pass and geo_pass and polaridad_pass:
            overall_status = "APROBADO"
        elif not self.last_ilrl_analysis_data or not self.last_geo_analysis_data or not self.last_polaridad_analysis_data:
            missing_tests = []
            if not self.last_ilrl_analysis_data:
                missing_tests.append("ILRL")
            if not self.last_geo_analysis_data:
                missing_tests.append("Geometría")
            if not self.last_polaridad_analysis_data:
                missing_tests.append("Polaridad")
            overall_status = f"RECHAZADO (Falta: {', '.join(missing_tests)})"
        else:
            overall_status = "RECHAZADO"
            
        color = "verde" if overall_status == "APROBADO" else "rojo" if "RECHAZADO" in overall_status else "orange"
        
        text_widget.insert(tk.END, f"🔍 Resultados para cable {self.serie_entry.get()}:\n\n", "header")
        text_widget.insert(tk.END, "🏁 ESTADO GENERAL: ", "bold")
        text_widget.insert(tk.END, f"{overall_status}\n\n", color)
        text_widget.insert(tk.END, "---", "normal")
        
        # ILRL Details
        text_widget.insert(tk.END, "\n\n📊 Detalles ILRL:\n", "subheader")
        ilrl_data = self.last_ilrl_analysis_data
        if ilrl_data:
            ilrl_status = ilrl_data.get('estado', 'N/A')
            ilrl_status_color = "verde" if ilrl_status == "APROBADO" else "rojo" if "RECHAZADO" in ilrl_status else "orange"
            text_widget.insert(tk.END, f"   Estado: ", "bold")
            text_widget.insert(tk.END, f"{ilrl_status}\n", ilrl_status_color)
            
            for lado, punta_data in ilrl_data.get('detalles_puntas', {}).items():
                text_widget.insert(tk.END, f"\n   - Lado {lado}: Estado de los conectores: {punta_data.get('estado', 'N/A')}\n", "normal")
                for conector in punta_data.get('conectores', []):
                    conector_status = conector.get('estado', 'N/A')
                    conector_color = "verde" if conector_status == "PASS" else "rojo" if "FAIL" in conector_status or "FALTAN" in conector_status else "orange"
                    text_widget.insert(tk.END, f"     • Conector {conector.get('conector')}: ", "normal")
                    text_widget.insert(tk.END, f"{conector_status}\n", conector_color)
                    
                    for fibra in conector.get('mediciones', []):
                        fibra_status = fibra.get('resultado', '').upper()
                        fibra_color = "verde" if fibra_status == "PASS" else "rojo"
                        text_widget.insert(tk.END, f"       - Fibra {fibra.get('fibra')}: ", "normal")
                        text_widget.insert(tk.END, f"{fibra_status}\n", fibra_color)
        else:
            text_widget.insert(tk.END, "   No se encontraron datos ILRL para este cable.\n", "orange")
            
        text_widget.insert(tk.END, "\n\n---", "normal")
        
        # Geometry Details
        text_widget.insert(tk.END, "\n\n📐 Detalles Geometría:\n", "subheader")
        geo_data = self.last_geo_analysis_data
        if geo_data:
            geo_status = geo_data.get('estado', 'N/A')
            geo_status_color = "verde" if geo_status == "APROBADO" else "rojo" if "RECHAZADO" in geo_status else "orange"
            text_widget.insert(tk.END, f"   Estado: ", "bold")
            text_widget.insert(tk.END, f"{geo_status}\n", geo_status_color)
            
            for lado, punta_data in geo_data.get('detalles_puntas', {}).items():
                mediciones = punta_data.get('mediciones', [])
                if mediciones:
                    text_widget.insert(tk.END, f"\n   - Lado {lado}:\n", "normal")
                    for medicion in mediciones:
                        conector_display = medicion.get('conector', 'N/A')
                        resultado_display = medicion.get('resultado', 'N/A')
                        
                        resultado_color = "verde" if resultado_display == "PASS" else "rojo" if resultado_display == "FAIL" else "orange"
                        text_widget.insert(tk.END, f"     • Conector {conector_display}: ", "normal")
                        text_widget.insert(tk.END, f"{resultado_display}\n", resultado_color)
                else:
                    text_widget.insert(tk.END, f"\n   - Lado {lado}: No se encontraron mediciones de conectores.\n", "orange")
        else:
            text_widget.insert(tk.END, "   No se encontraron datos de Geometría para este cable.\n", "orange")
            
        text_widget.insert(tk.END, "\n\n---", "normal")
        
        # Polarity Details
        text_widget.insert(tk.END, "\n\n🔀 Detalles Polaridad:\n", "subheader")
        polaridad_data = self.last_polaridad_analysis_data
        if polaridad_data:
            polaridad_status = polaridad_data.get('status', 'N/A')
            polaridad_status_color = "verde" if polaridad_status == "PASS" else "rojo" if polaridad_status == "FAIL" else "orange"
            
            text_widget.insert(tk.END, f"   Estado: ", "bold")
            text_widget.insert(tk.END, f"{polaridad_status}\n", polaridad_status_color)
            
            text_widget.insert(tk.END, f"   Fecha de Prueba: ", "bold")
            text_widget.insert(tk.END, f"{polaridad_data.get('date', 'N/A')}\n", "normal")

        else:
            text_widget.insert(tk.END, "   No se encontraron datos de Polaridad para este cable.\n", "orange")

        text_widget.config(state=tk.DISABLED)

        btn_frame = ttk.Frame(scrollable_frame, style="DetallesTotales.TFrame")
        btn_frame.pack(fill=tk.X, pady=(15, 0))
        ttk.Button(btn_frame, text="Cerrar", command=details_window.destroy).pack()
        
        details_window.mainloop()

    def toggle_fullscreen(self, event=None):
        """Alterna el modo de pantalla completa."""
        self.root.attributes("-fullscreen", not self.root.attributes("-fullscreen"))

    def exit_fullscreen(self, event=None):
        """Sale del modo de pantalla completa."""
        self.root.attributes("-fullscreen", False)

    def configurar_ot(self):
        """Muestra la ventana para configurar una OT y sus detalles de cable."""
        config_window = tk.Toplevel(self.root)
        config_window.title("Configurar O.T. y Detalles de Cable")
        config_window.geometry("800x700") 
        config_window.transient(self.root)
        config_window.grab_set()

        main_frame = ttk.Frame(config_window)
        main_frame.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(main_frame, highlightthickness=0)
        yscrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        xscrollbar = ttk.Scrollbar(main_frame, orient="horizontal", command=canvas.xview)
        scrollable_content_frame = ttk.Frame(canvas)
        
        scrollable_window_id = canvas.create_window((0, 0), window=scrollable_content_frame, anchor="nw")
        canvas.configure(yscrollcommand=yscrollbar.set, xscrollcommand=xscrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        yscrollbar.pack(side="right", fill="y")
        xscrollbar.pack(side="bottom", fill="x")

        def _on_config_canvas_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas_width = event.width
            canvas.itemconfigure(scrollable_window_id, width=canvas_width)
            
        canvas.bind("<Configure>", _on_config_canvas_configure)
        canvas.bind_all("<MouseWheel>", self._on_mouse_wheel)
        
        frame = ttk.Frame(scrollable_content_frame, padding=(20, 20), style="TFrame")
        frame.pack(fill=tk.BOTH, expand=True)

        ot_var = tk.StringVar(config_window)
        drawing_number_var = tk.StringVar(config_window) 
        link_var = tk.StringVar(config_window)
        num_conectores_a_var = tk.StringVar(config_window)
        fibras_por_conector_a_var = tk.StringVar(config_window)
        num_conectores_b_var = tk.StringVar(config_window)
        fibras_por_conector_b_var = tk.StringVar(config_window)
        
        ilrl_ot_header_var = tk.StringVar(config_window)
        ilrl_serie_header_var = tk.StringVar(config_window)
        ilrl_fecha_header_var = tk.StringVar(config_window)
        ilrl_hora_header_var = tk.StringVar(config_window)
        ilrl_estado_header_var = tk.StringVar(config_window)
        ilrl_conector_header_var = tk.StringVar(config_window)
        
        current_ot = self.ot_entry.get().strip().upper()
        if not current_ot.startswith('JMO-'):
            current_ot = f"JMO-{current_ot}"

        ot_config_data = self._cargar_ot_configuration(current_ot)
        if ot_config_data:
            drawing_number_var.set(ot_config_data.get('drawing_number', ''))
            link_var.set(ot_config_data.get('link', ''))
            num_conectores_a_var.set(str(ot_config_data.get('num_conectores_a', 1)))
            fibras_por_conector_a_var.set(str(ot_config_data.get('fibers_per_connector_a', 12)))
            num_conectores_b_var.set(str(ot_config_data.get('num_conectores_b', 1)))
            fibras_por_conector_b_var.set(str(ot_config_data.get('fibers_per_connector_b', 12)))
            ilrl_ot_header_var.set(ot_config_data.get('ilrl_ot_header', 'Work number'))
            ilrl_serie_header_var.set(ot_config_data.get('ilrl_serie_header', 'Serial number'))
            ilrl_fecha_header_var.set(ot_config_data.get('ilrl_fecha_header', 'Date'))
            ilrl_hora_header_var.set(ot_config_data.get('ilrl_hora_header', 'Time'))
            ilrl_estado_header_var.set(ot_config_data.get('ilrl_estado_header', 'Alarm Status'))
            ilrl_conector_header_var.set(ot_config_data.get('ilrl_conector_header', 'connector label'))
        else:
            drawing_number_var.set("")
            link_var.set("")
            num_conectores_a_var.set("1")
            fibras_por_conector_a_var.set("12")
            num_conectores_b_var.set("1")
            fibras_por_conector_b_var.set("12")
            ilrl_ot_header_var.set('Work number')
            ilrl_serie_header_var.set('Serial number')
            ilrl_fecha_header_var.set('Date')
            ilrl_hora_header_var.set('Time')
            ilrl_estado_header_var.set('Alarm Status')
            ilrl_conector_header_var.set('connector label')
        
        ot_var.set(current_ot)

        input_section_frame = ttk.LabelFrame(frame, text=" Configuración de O.T. ", padding=(10, 10))
        input_section_frame.pack(fill=tk.X, pady=10)
        input_section_frame.columnconfigure(1, weight=1)

        ttk.Label(input_section_frame, text="Número de OT:", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky=tk.W, pady=5)
        ot_entry_config = ttk.Entry(input_section_frame, textvariable=ot_var, width=30)
        ot_entry_config.grid(row=0, column=1, sticky=tk.EW, pady=5, padx=10)

        ttk.Label(input_section_frame, text="Número de Dibujo:", font=("Arial", 10)).grid(row=1, column=0, sticky=tk.W, pady=5)
        drawing_number_entry = ttk.Entry(input_section_frame, textvariable=drawing_number_var, width=50)
        drawing_number_entry.grid(row=1, column=1, sticky=tk.EW, pady=5, padx=10)

        ttk.Label(input_section_frame, text="Link a Dibujo Técnico:", font=("Arial", 10)).grid(row=2, column=0, sticky=tk.W, pady=5)
        link_entry = ttk.Entry(input_section_frame, textvariable=link_var, width=50)
        link_entry.grid(row=2, column=1, sticky=tk.EW, pady=5, padx=10)

        ttk.Separator(input_section_frame, orient='horizontal').grid(row=3, columnspan=2, sticky='ew', pady=10)

        mpo_config_frame = ttk.LabelFrame(frame, text=" Configuración del Cable MPO ", padding=(10, 10))
        mpo_config_frame.pack(fill=tk.X, pady=10)
        mpo_config_frame.columnconfigure(1, weight=1)

        ttk.Label(mpo_config_frame, text="Lado A - Conectores:", font=("Arial", 10)).grid(row=0, column=0, sticky=tk.W, pady=5)
        num_conectores_a_entry = ttk.Entry(mpo_config_frame, textvariable=num_conectores_a_var, width=10)
        num_conectores_a_entry.grid(row=0, column=1, sticky=tk.W, pady=5, padx=10)

        ttk.Label(mpo_config_frame, text="Lado A - Fibras/Conector:", font=("Arial", 10)).grid(row=1, column=0, sticky=tk.W, pady=5)
        fibras_por_conector_a_entry = ttk.Entry(mpo_config_frame, textvariable=fibras_por_conector_a_var, width=10)
        fibras_por_conector_a_entry.grid(row=1, column=1, sticky=tk.W, pady=5, padx=10)

        ttk.Label(mpo_config_frame, text="Lado B - Conectores:", font=("Arial", 10)).grid(row=2, column=0, sticky=tk.W, pady=5)
        num_conectores_b_entry = ttk.Entry(mpo_config_frame, textvariable=num_conectores_b_var, width=10)
        num_conectores_b_entry.grid(row=2, column=1, sticky=tk.W, pady=5, padx=10)

        ttk.Label(mpo_config_frame, text="Lado B - Fibras/Conector:", font=("Arial", 10)).grid(row=3, column=0, sticky=tk.W, pady=5)
        fibras_por_conector_b_entry = ttk.Entry(mpo_config_frame, textvariable=fibras_por_conector_b_var, width=10)
        fibras_por_conector_b_entry.grid(row=3, column=1, sticky=tk.W, pady=5, padx=10)
        
        ilrl_headers_frame = ttk.LabelFrame(frame, text=" Encabezados de Excel ILRL ", padding=(10, 10))
        ilrl_headers_frame.pack(fill=tk.X, pady=10)
        ilrl_headers_frame.columnconfigure(1, weight=1)

        ttk.Label(ilrl_headers_frame, text="Encabezado de O.T.:", font=("Arial", 10)).grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(ilrl_headers_frame, textvariable=ilrl_ot_header_var, width=30).grid(row=0, column=1, sticky=tk.EW, padx=10)
        
        ttk.Label(ilrl_headers_frame, text="Encabezado de Nro. de Serie:", font=("Arial", 10)).grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(ilrl_headers_frame, textvariable=ilrl_serie_header_var, width=30).grid(row=1, column=1, sticky=tk.EW, padx=10)

        ttk.Label(ilrl_headers_frame, text="Encabezado de Fecha:", font=("Arial", 10)).grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(ilrl_headers_frame, textvariable=ilrl_fecha_header_var, width=30).grid(row=2, column=1, sticky=tk.EW, padx=10)

        ttk.Label(ilrl_headers_frame, text="Encabezado de Hora:", font=("Arial", 10)).grid(row=3, column=0, sticky=tk.W, pady=5)
        ttk.Entry(ilrl_headers_frame, textvariable=ilrl_hora_header_var, width=30).grid(row=3, column=1, sticky=tk.EW, padx=10)
        
        ttk.Label(ilrl_headers_frame, text="Encabezado de Estado:", font=("Arial", 10)).grid(row=4, column=0, sticky=tk.W, pady=5)
        ttk.Entry(ilrl_headers_frame, textvariable=ilrl_estado_header_var, width=30).grid(row=4, column=1, sticky=tk.EW, padx=10)
        
        ttk.Label(ilrl_headers_frame, text="Encabezado de Conector:", font=("Arial", 10)).grid(row=5, column=0, sticky=tk.W, pady=5)
        ttk.Entry(ilrl_headers_frame, textvariable=ilrl_conector_header_var, width=30).grid(row=5, column=1, sticky=tk.EW, padx=10)
        
        canvas_frame = ttk.LabelFrame(frame, text=" Vista Previa del Cable MPO ", padding=(10, 10))
        canvas_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        cable_canvas = Canvas(canvas_frame, bg="white", highlightthickness=1, highlightbackground="#DDDDDD")
        cable_canvas.pack(fill=tk.BOTH, expand=True)

        def draw_mpo_cable_config():
            cable_canvas.delete("all")
            
            canvas_width = cable_canvas.winfo_width() if cable_canvas.winfo_width() > 1 else 600
            canvas_height = cable_canvas.winfo_height() if cable_canvas.winfo_height() > 1 else 200

            center_y = canvas_height / 2
            cable_length = canvas_width - 150 
            cable_start_x = 75
            cable_end_x = cable_start_x + cable_length
            
            cable_canvas.create_line(cable_start_x, center_y, cable_end_x, center_y, 
                                     width=10, fill="#607D8B", capstyle=tk.ROUND)

            try:
                num_ca = int(num_conectores_a_var.get())
                fibras_ca = int(fibras_por_conector_a_var.get())
                num_cb = int(num_conectores_b_var.get())
                fibras_cb = int(fibras_por_conector_b_var.get())
            except ValueError:
                cable_canvas.create_text(canvas_width/2, canvas_height/2, text="Valores de entrada inválidos", fill="red", font=("Arial", 12))
                return

            connector_width = 30
            connector_height_unit = 20
            spacing_between_conectores = 10

            total_height_a = num_ca * connector_height_unit + (num_ca - 1) * spacing_between_conectores
            start_y_a = center_y - total_height_a / 2

            for i in range(num_ca):
                conn_y = start_y_a + i * (connector_height_unit + spacing_between_conectores)
                cable_canvas.create_rectangle(cable_start_x - connector_width, conn_y, 
                                            cable_start_x, conn_y + connector_height_unit, 
                                            fill="#A2A2A2", outline="#333333", width=1)
                cable_canvas.create_text(cable_start_x - connector_width/2, conn_y + connector_height_unit/2, 
                                        text=f"A{i+1}", fill="white", font=("Arial", 8, "bold"))

                fiber_radius = 2
                fiber_spacing = (connector_height_unit - (2 * fiber_radius)) / (fibras_ca + 1)
                for f in range(fibras_ca):
                    fiber_y = conn_y + fiber_radius + (f + 1) * fiber_spacing
                    cable_canvas.create_oval(cable_start_x - connector_width + 5, fiber_y - fiber_radius,
                                            cable_start_x - 5, fiber_y + fiber_radius,
                                            fill="#4CAF50", outline="#388E3C")

            total_height_b = num_cb * connector_height_unit + (num_cb - 1) * spacing_between_conectores
            start_y_b = center_y - total_height_b / 2

            for i in range(num_cb):
                conn_y = start_y_b + i * (connector_height_unit + spacing_between_conectores)
                cable_canvas.create_rectangle(cable_end_x, conn_y, 
                                            cable_end_x + connector_width, conn_y + connector_height_unit, 
                                            fill="#A2A2A2", outline="#333333", width=1)
                cable_canvas.create_text(cable_end_x + connector_width/2, conn_y + connector_height_unit/2, 
                                        text=f"B{i+1}", fill="white", font=("Arial", 8, "bold"))
                
                fiber_spacing = (connector_height_unit - (2 * fiber_radius)) / (fibras_cb + 1)
                for f in range(fibras_cb):
                    fiber_y = conn_y + fiber_radius + (f + 1) * fiber_spacing
                    cable_canvas.create_oval(cable_end_x + 5, fiber_y - fiber_radius,
                                            cable_end_x + connector_width - 5, fiber_y + fiber_radius,
                                            fill="#FFC107", outline="#FFA000")

            total_fibers_a = num_ca * fibras_ca
            total_fibers_b = num_cb * fibras_cb
            cable_canvas.create_text(cable_start_x, center_y + total_height_a/2 + 30, 
                                    text=f"Total Fibras Lado A: {total_fibers_a}", 
                                    anchor="n", font=("Arial", 9, "bold"), fill="blue")
            cable_canvas.create_text(cable_end_x, center_y + total_height_b/2 + 30, 
                                    text=f"Total Fibras Lado B: {total_fibers_b}", 
                                    anchor="n", font=("Arial", 9, "bold"), fill="blue")

        num_conectores_a_var.trace_add("write", lambda *args: draw_mpo_cable_config())
        fibras_por_conector_a_var.trace_add("write", lambda *args: draw_mpo_cable_config())
        num_conectores_b_var.trace_add("write", lambda *args: draw_mpo_cable_config())
        fibras_por_conector_b_var.trace_add("write", lambda *args: draw_mpo_cable_config())

        cable_canvas.bind("<Configure>", lambda event: draw_mpo_cable_config())
        config_window.after(100, draw_mpo_cable_config)

        def guardar_configuracion_ot():
            ot = ot_var.get().strip().upper()
            drawing_number = drawing_number_var.get().strip()
            link = link_var.get().strip()
            
            ilrl_ot_header = ilrl_ot_header_var.get().strip()
            ilrl_serie_header = ilrl_serie_header_var.get().strip()
            ilrl_fecha_header = ilrl_fecha_header_var.get().strip()
            ilrl_hora_header = ilrl_hora_header_var.get().strip()
            ilrl_estado_header = ilrl_estado_header_var.get().strip()
            ilrl_conector_header = ilrl_conector_header_var.get().strip()

            try:
                num_ca = int(num_conectores_a_var.get())
                fibras_ca = int(fibras_por_conector_a_var.get())
                num_cb = int(num_conectores_b_var.get())
                fibras_cb = int(fibras_por_conector_b_var.get())

                if not all(x > 0 for x in [num_ca, fibras_ca, num_cb, fibras_cb]):
                    messagebox.showerror("Error de Entrada", "Todos los valores de conectores y fibras deben ser números enteros positivos.")
                    return
                
                if link and not (link.startswith("http://") or link.startswith("https://")):
                    messagebox.showwarning("Formato de Link Inválido", "El enlace al dibujo técnico debe empezar con 'http://' o 'https://'.")
                    return

                ot_data = {
                    'ot_number': ot,
                    'drawing_number': drawing_number,
                    'link': link,
                    'num_conectores_a': num_ca,
                    'fibers_per_connector_a': fibras_ca,
                    'num_conectores_b': num_cb,
                    'fibers_per_connector_b': fibras_cb,
                    'ilrl_ot_header': ilrl_ot_header,
                    'ilrl_serie_header': ilrl_serie_header,
                    'ilrl_fecha_header': ilrl_fecha_header,
                    'ilrl_hora_header': ilrl_hora_header,
                    'ilrl_estado_header': ilrl_estado_header,
                    'ilrl_conector_header': ilrl_conector_header
                }

                self._guardar_ot_configuration(ot_data)
                messagebox.showinfo("Éxito", f"Configuración para OT '{ot}' guardada correctamente.")
                config_window.destroy()

            except ValueError:
                messagebox.showerror("Error de Entrada", "Por favor, ingrese números válidos en los campos de conectores y fibras.")
            except Exception as e:
                messagebox.showerror("Error", f"Ocurrió un error al guardar la configuración: {e}")

        btn_guardar = ttk.Button(frame, text="💾 Guardar Configuración", command=guardar_configuracion_ot)
        btn_guardar.pack(pady=10)

        config_window.mainloop()

    def mostrar_detalles_ot_actual(self):
        """Muestra los detalles de la OT ingresada en la pantalla principal."""
        ot_input = self.ot_entry.get().strip().upper()
        if not ot_input.startswith('JMO-'):
            ot_input = f"JMO-{ot_input}"
        
        if not ot_input:
            messagebox.showwarning("Falta OT", "Por favor, ingrese un número de O.T. para ver sus detalles.")
            return

        ot_data = self._cargar_ot_configuration(ot_input)
        
        if not ot_data:
            messagebox.showinfo("No Encontrado", f"No se encontró ninguna configuración para la O.T.: {ot_input}")
            return
            
        details_window = tk.Toplevel(self.root)
        details_window.title(f"Detalles de la O.T. {ot_data['ot_number']}")
        details_window.geometry("600x450")
        details_window.transient(self.root)
        details_window.grab_set()

        main_frame = ttk.Frame(details_window)
        main_frame.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(main_frame, highlightthickness=0)
        yscrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        xscrollbar = ttk.Scrollbar(main_frame, orient="horizontal", command=canvas.xview)
        scrollable_content_frame = ttk.Frame(canvas)

        scrollable_window_id = canvas.create_window((0, 0), window=scrollable_content_frame, anchor="nw")
        canvas.configure(yscrollcommand=yscrollbar.set, xscrollcommand=xscrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        yscrollbar.pack(side="right", fill="y")
        xscrollbar.pack(side="bottom", fill="x")

        def _on_ot_details_canvas_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas_width = event.width
            canvas.itemconfigure(scrollable_window_id, width=canvas_width)

        canvas.bind("<Configure>", _on_ot_details_canvas_configure)
        canvas.bind_all("<MouseWheel>", self._on_mouse_wheel)
        
        frame = ttk.Frame(scrollable_content_frame, padding=(20, 20), style="TFrame")
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text="📋 Detalles de la Orden de Trabajo", font=("Arial", 12, "bold")).pack(anchor="w", pady=(0, 10))
        
        ttk.Label(frame, text=f"• Número de O.T.: {ot_data['ot_number']}", font=("Arial", 10)).pack(anchor="w", pady=2)
        ttk.Label(frame, text=f"• Número de Dibujo: {ot_data['drawing_number']}", font=("Arial", 10)).pack(anchor="w", pady=2)
        
        link_text = ot_data['link'] if ot_data['link'] and (ot_data['link'].startswith("http://") or ot_data['link'].startswith("https://")) else "N/A"
        link_label = ttk.Label(frame, text=f"• Link al Dibujo: {link_text}", font=("Arial", 10), foreground="blue", cursor="hand2")
        link_label.pack(anchor="w", pady=2)
        if link_text != "N/A":
            link_label.bind("<Button-1>", lambda e: webbrowser.open_new(ot_data['link']))
        else:
            link_label.config(foreground="gray", cursor="")

        ttk.Separator(frame, orient='horizontal').pack(fill='x', pady=10)
        
        ttk.Label(frame, text="Especificaciones del Cable MPO:", font=("Arial", 10, "bold")).pack(anchor="w", pady=(0, 5))
        
        fibers_a_display = ot_data['fibers_per_connector_a']
        fibers_a_text = str(fibers_a_display) if fibers_a_display is not None and fibers_a_display > 0 else "N/A"
        ttk.Label(frame, text=f"• Lado A: {ot_data['num_conectores_a']} Conectores, {fibers_a_text} Fibras/Conector", font=("Arial", 10)).pack(anchor="w", pady=2)
        
        fibers_b_display = ot_data['fibers_per_connector_b']
        fibers_b_text = str(fibers_b_display) if fibers_b_display is not None and fibers_b_display > 0 else "N/A"
        ttk.Label(frame, text=f"• Lado B: {ot_data['num_conectores_b']} Conectores, {fibers_b_text} Fibras/Conector", font=("Arial", 10)).pack(anchor="w", pady=2)
        
        ttk.Separator(frame, orient='horizontal').pack(fill='x', pady=10)
        ttk.Label(frame, text="Encabezados de Excel ILRL configurados:", font=("Arial", 10, "bold")).pack(anchor="w", pady=(0, 5))
        
        ttk.Label(frame, text=f"• O.T.: {ot_data.get('ilrl_ot_header', 'N/A')}", font=("Arial", 10)).pack(anchor="w", pady=2)
        ttk.Label(frame, text=f"• Nro. de Serie: {ot_data.get('ilrl_serie_header', 'N/A')}", font=("Arial", 10)).pack(anchor="w", pady=2)
        ttk.Label(frame, text=f"• Fecha: {ot_data.get('ilrl_fecha_header', 'N/A')}", font=("Arial", 10)).pack(anchor="w", pady=2)
        ttk.Label(frame, text=f"• Hora: {ot_data.get('ilrl_hora_header', 'N/A')}", font=("Arial", 10)).pack(anchor="w", pady=2)
        ttk.Label(frame, text=f"• Estado: {ot_data.get('ilrl_estado_header', 'N/A')}", font=("Arial", 10)).pack(anchor="w", pady=2)
        ttk.Label(frame, text=f"• Conector: {ot_data.get('ilrl_conector_header', 'N/A')}", font=("Arial", 10)).pack(anchor="w", pady=2)

        ttk.Button(frame, text="Cerrar", command=details_window.destroy).pack(pady=20)

        details_window.mainloop()

    def _init_ot_database(self):
        """Inicializa la tabla de configuraciones de OTs si no existe y asegura que todas las columnas estén presentes."""
        conn = None
        try:
            conn = sqlite3.connect(self.db_name, check_same_thread=False)
            cursor = conn.cursor()
            
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS ot_configurations (
                    ot_number TEXT PRIMARY KEY,
                    drawing_number TEXT,
                    link TEXT,
                    num_conectores_a INTEGER,
                    fibers_per_connector_a INTEGER,
                    num_conectores_b INTEGER,
                    fibers_per_connector_b INTEGER,
                    ilrl_ot_header TEXT,
                    ilrl_serie_header TEXT,
                    ilrl_fecha_header TEXT,
                    ilrl_hora_header TEXT,
                    ilrl_estado_header TEXT,
                    ilrl_conector_header TEXT,
                    last_modified TEXT
                )
            """)
            conn.commit()

            expected_columns = {
                'drawing_number': 'TEXT', 'link': 'TEXT', 'num_conectores_a': 'INTEGER',
                'fibers_per_connector_a': 'INTEGER', 'num_conectores_b': 'INTEGER',
                'fibers_per_connector_b': 'INTEGER', 'last_modified': 'TEXT',
                'ilrl_ot_header': 'TEXT', 'ilrl_serie_header': 'TEXT', 'ilrl_fecha_header': 'TEXT',
                'ilrl_hora_header': 'TEXT', 'ilrl_estado_header': 'TEXT', 'ilrl_conector_header': 'TEXT'
            }

            cursor.execute("PRAGMA table_info(ot_configurations)")
            existing_columns = [info[1] for info in cursor.fetchall()]

            for col_name, col_type in expected_columns.items():
                if col_name not in existing_columns:
                    default_value = "'N/A'" if col_type == 'TEXT' else '0'
                    cursor.execute(f"ALTER TABLE ot_configurations ADD COLUMN {col_name} {col_type} DEFAULT {default_value}")
                    conn.commit()
                    print(f"DEBUG: Columna '{col_name}' añadida a la tabla 'ot_configurations'.")
                
        except sqlite3.Error as e:
            messagebox.showerror("Error de Base de Datos", 
                            f"No se pudo inicializar la tabla de configuraciones de OTs o añadir columnas: {e}")
        finally:
            if conn:
                conn.close()

    def _guardar_ot_configuration(self, ot_data):
        """Guarda o actualiza la configuración de una OT en la base de datos."""
        conn = None
        try:
            conn = sqlite3.connect(self.db_name, check_same_thread=False)
            cursor = conn.cursor()
            last_modified = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            cursor.execute("""
                INSERT INTO ot_configurations (
                    ot_number, drawing_number, link, num_conectores_a,
                    fibers_per_connector_a, num_conectores_b, fibers_per_connector_b,
                    ilrl_ot_header, ilrl_serie_header, ilrl_fecha_header,
                    ilrl_hora_header, ilrl_estado_header, ilrl_conector_header, last_modified
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(ot_number) DO UPDATE SET
                    drawing_number = excluded.drawing_number,
                    link = excluded.link,
                    num_conectores_a = excluded.num_conectores_a,
                    fibers_per_connector_a = excluded.fibers_per_connector_a,
                    num_conectores_b = excluded.num_conectores_b,
                    fibers_per_connector_b = excluded.fibers_per_connector_b,
                    ilrl_ot_header = excluded.ilrl_ot_header,
                    ilrl_serie_header = excluded.ilrl_serie_header,
                    ilrl_fecha_header = excluded.ilrl_fecha_header,
                    ilrl_hora_header = excluded.ilrl_hora_header,
                    ilrl_estado_header = excluded.ilrl_estado_header,
                    ilrl_conector_header = excluded.ilrl_conector_header,
                    last_modified = excluded.last_modified
            """, (
                ot_data['ot_number'], ot_data['drawing_number'], ot_data['link'],
                ot_data['num_conectores_a'], ot_data['fibers_per_connector_a'],
                ot_data['num_conectores_b'], ot_data['fibers_per_connector_b'],
                ot_data['ilrl_ot_header'], ot_data['ilrl_serie_header'],
                ot_data['ilrl_fecha_header'], ot_data['ilrl_hora_header'],
                ot_data['ilrl_estado_header'], ot_data['ilrl_conector_header'], last_modified
            ))
            conn.commit()
        except sqlite3.Error as e:
            messagebox.showerror("Error de Base de Datos", f"No se pudo guardar la configuración de la OT: {e}")
        finally:
            if conn:
                conn.close()

    def _cargar_ot_configuration(self, ot_number):
        """Carga la configuración de una OT desde la base de datos."""
        conn = None
        try:
            conn = sqlite3.connect(self.db_name, check_same_thread=False)
            cursor = conn.cursor()
            
            cursor.execute("PRAGMA table_info(ot_configurations)")
            column_info = cursor.fetchall()
            column_names = [info[1] for info in column_info]
            
            cursor.execute("SELECT * FROM ot_configurations WHERE ot_number = ?", (ot_number,))
            row = cursor.fetchone()
            
            if row:
                ot_data = {column_names[i]: row[i] for i in range(len(column_names))}
                
                return {
                    'ot_number': ot_data.get('ot_number'),
                    'drawing_number': ot_data.get('drawing_number', ''),
                    'link': ot_data.get('link', ''),
                    'num_conectores_a': ot_data.get('num_conectores_a', 0),
                    'fibers_per_connector_a': ot_data.get('fibers_per_connector_a', 0),
                    'num_conectores_b': ot_data.get('num_conectores_b', 0),
                    'fibers_per_connector_b': ot_data.get('fibers_per_connector_b', 0),
                    'ilrl_ot_header': ot_data.get('ilrl_ot_header', 'Work number'),
                    'ilrl_serie_header': ot_data.get('ilrl_serie_header', 'Serial number'),
                    'ilrl_fecha_header': ot_data.get('ilrl_fecha_header', 'Date'),
                    'ilrl_hora_header': ot_data.get('ilrl_hora_header', 'Time'),
                    'ilrl_estado_header': ot_data.get('ilrl_estado_header', 'Alarm Status'),
                    'ilrl_conector_header': ot_data.get('ilrl_conector_header', 'connector label'),
                    'last_modified': ot_data.get('last_modified', '')
                }
            return None
        except sqlite3.Error as e:
            messagebox.showerror("Error de Base de Datos", f"No se pudo cargar la configuración de la OT: {e}")
            return None
        finally:
            if conn:
                conn.close()

    def _init_db_path(self):
        """Inicializa la ruta de la base de datos cargando desde config.json o usando la predeterminada"""
        app_data_dir = os.path.join(os.path.expanduser('~'), 'Documents', 'VerificadorCablesData')
        default_db_path = os.path.join(app_data_dir, "cable_verifications.db")
    
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r') as f:
                    config = json.load(f)
                    self.ruta_base_ilrl = config.get('ruta_ilrl', self.ruta_base_ilrl)
                    self.ruta_base_geo = config.get('ruta_geo', self.ruta_base_geo)
                    self.ruta_base_polaridad = config.get('ruta_polaridad', self.ruta_base_polaridad)
                    self.db_name = config.get('ruta_db', default_db_path)
            except Exception as e:
                messagebox.showerror("Error de Configuración", 
                                f"No se pudo cargar la configuración de DB: {e}. Usando ruta por defecto.")
                self.db_name = default_db_path
                self.guardar_rutas()
        else:
            self.db_name = default_db_path
            self.guardar_rutas()
    
        os.makedirs(os.path.dirname(self.db_name), exist_ok=True)
        self._init_database()

    def _init_database(self):
        """Inicializa la base de datos SQLite y crea la tabla si no existe."""
        conn = None
        try:
            conn = sqlite3.connect(self.db_name, check_same_thread=False)
            cursor = conn.cursor()
        
            cursor.execute("""
                SELECT count(name) FROM sqlite_master 
                WHERE type='table' AND name='cable_verifications'
            """)
        
            if cursor.fetchone()[0] == 0:
                cursor.execute("""
                    CREATE TABLE cable_verifications (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        entry_date TEXT NOT NULL,
                        serial_number TEXT NOT NULL,
                        ot_number TEXT NOT NULL,
                        overall_status TEXT NOT NULL,
                        ilrl_status TEXT,
                        ilrl_date TEXT,
                        geo_status TEXT,
                        geo_date TEXT,
                        polaridad_status TEXT,
                        polaridad_date TEXT,
                        ilrl_details_json TEXT,
                        geo_details_json TEXT,
                        polaridad_details_json TEXT
                    )
                """)
                conn.commit()
            else:
                existing_cols = [col[1] for col in cursor.execute("PRAGMA table_info(cable_verifications)").fetchall()]
                if 'polaridad_status' not in existing_cols:
                    cursor.execute("ALTER TABLE cable_verifications ADD COLUMN polaridad_status TEXT")
                if 'polaridad_date' not in existing_cols:
                    cursor.execute("ALTER TABLE cable_verifications ADD COLUMN polaridad_date TEXT")
                if 'polaridad_details_json' not in existing_cols:
                    cursor.execute("ALTER TABLE cable_verifications ADD COLUMN polaridad_details_json TEXT")
                conn.commit()
            
        except sqlite3.Error as e:
            messagebox.showerror("Error de Base de Datos", 
                            f"No se pudo inicializar la base de datos: {e}")
            if not os.path.exists(self.db_name):
                try:
                    open(self.db_name, 'w').close()
                    messagebox.showinfo("Base de Datos", "Archivo de base de datos creado. Intente reiniciar la aplicación.")
                except Exception as e:
                    messagebox.showerror("Error Crítico", 
                                    f"No se pudo crear el archivo de base de datos: {e}")
        finally:
            if conn:
                conn.close()

    def _log_verification_result(self, serial_number, ot_number, overall_status, 
                           ilrl_status, ilrl_date, ilrl_details, 
                           geo_status, geo_date, geo_details,
                           polaridad_status, polaridad_date, polaridad_details):
        """Registra el resultado de la verificación de un cable en la base de datos."""
        conn = None
        try:
            conn = sqlite3.connect(self.db_name, check_same_thread=False)
            cursor = conn.cursor()
            entry_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            # Asegurar que el objeto datetime se convierta a string antes de serializar a JSON
            if polaridad_details and 'date_dt' in polaridad_details:
                polaridad_details['date_dt'] = polaridad_details['date_dt'].strftime("%Y-%m-%d %H:%M:%S")

            ilrl_details_json = json.dumps(ilrl_details, ensure_ascii=False) if ilrl_details else None
            geo_details_json = json.dumps(geo_details, ensure_ascii=False) if geo_details else None
            polaridad_details_json = json.dumps(polaridad_details, ensure_ascii=False) if polaridad_details else None

            cursor.execute("""
                INSERT INTO cable_verifications (
                    entry_date, serial_number, ot_number, overall_status,
                    ilrl_status, ilrl_date, ilrl_details_json,
                    geo_status, geo_date, geo_details_json,
                    polaridad_status, polaridad_date, polaridad_details_json
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                entry_date, serial_number, ot_number, overall_status,
                ilrl_status, ilrl_date, ilrl_details_json,
                geo_status, geo_date, geo_details_json,
                polaridad_status, polaridad_date, polaridad_details_json
            ))
        
            conn.commit()
        
        except sqlite3.Error as e:
            messagebox.showerror("Error de Base de Datos", 
                            f"No se pudo registrar el resultado: {e}\n"
                            f"Base de datos: {os.path.abspath(self.db_name)}")
        finally:
            if conn:
                conn.close()

    def verificar_ruta_db(self):
        """Muestra la ruta real de la base de datos para diagnóstico."""
        ruta_absoluta = os.path.abspath(self.db_name)
        messagebox.showinfo(
            "Ubicación de la Base de Datos",
            f"La base de datos se está guardando en:\n\n{ruta_absoluta}\n\n"
            f"Tamaño del archivo: {os.path.getsize(self.db_name) if os.path.exists(self.db_name) else 0} bytes"
        )

    def cargar_rutas(self):
        """Carga las rutas de los archivos desde un archivo de configuración JSON"""
        app_data_dir = os.path.join(os.path.expanduser('~'), 'Documents', 'VerificadorCablesData')
        default_db_path = os.path.join(app_data_dir, "cable_verifications.db")
    
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r') as f:
                    config = json.load(f)
                    self.ruta_base_ilrl = config.get('ruta_ilrl', self.ruta_base_ilrl)
                    self.ruta_base_geo = config.get('ruta_geo', self.ruta_base_geo)
                    self.ruta_base_polaridad = config.get('ruta_polaridad', self.ruta_base_polaridad)
                    self.db_name = config.get('ruta_db', default_db_path)
            except Exception as e:
                messagebox.showerror("Error de Configuración", 
                                f"No se pudo cargar la configuración de DB: {e}. Usando ruta por defecto.")
                self.db_name = default_db_path
                self.guardar_rutas()
        else:
            self.db_name = default_db_path
            self.guardar_rutas()
    
        os.makedirs(os.path.dirname(self.db_name), exist_ok=True)

    def guardar_rutas(self):
        """Guarda las rutas actuales en un archivo de configuración JSON"""
        config = {
            'ruta_ilrl': self.ruta_base_ilrl,
            'ruta_geo': self.ruta_base_geo,
            'ruta_polaridad': self.ruta_base_polaridad,
            'ruta_db': self.db_name
        }
        try:
            with open(self.config_file, 'w') as f:
                json.dump(config, f, indent=4)
            messagebox.showinfo("Configuración Guardada", "Las rutas se han guardado correctamente.")
        except Exception as e:
            messagebox.showerror("Error al Guardar", f"No se pudieron guardar las rutas: {e}")

    def leer_resultado_ilrl(self, ruta, cable_config_ot):
        """
        Lee el archivo de resultados ILRL, extrayendo los datos de cada fibra y organizándolos
        por cable.
        """
        try:
            if not os.path.exists(ruta):
                print(f"ILRL Debug: Archivo no encontrado en leer_resultado_ilrl: {ruta}")
                return None
            
            num_conectores_a = cable_config_ot.get('num_conectores_a', 1)
            fibras_por_conector_a = cable_config_ot.get('fibers_per_connector_a', 12)
            num_conectores_b = cable_config_ot.get('num_conectores_b', 1)
            fibras_por_conector_b = cable_config_ot.get('fibers_per_connector_b', 12)

            ilrl_ot_header = cable_config_ot.get('ilrl_ot_header', 'Work number')
            ilrl_serie_header = cable_config_ot.get('ilrl_serie_header', 'Serial number')
            ilrl_fecha_header = cable_config_ot.get('ilrl_fecha_header', 'Date')
            ilrl_hora_header = cable_config_ot.get('ilrl_hora_header', 'Time')
            ilrl_estado_header = cable_config_ot.get('ilrl_estado_header', 'Alarm Status')
            ilrl_conector_header = cable_config_ot.get('ilrl_conector_header', 'connector label')

            df = pd.read_excel(ruta, sheet_name="Results")
            
            required_headers = [ilrl_ot_header, ilrl_serie_header, ilrl_fecha_header, ilrl_hora_header, ilrl_estado_header, ilrl_conector_header]
            if not all(h in df.columns for h in required_headers):
                missing_headers = [h for h in required_headers if h not in df.columns]
                messagebox.showerror("Error de Encabezado ILRL", 
                                     f"No se encontraron los siguientes encabezados en el archivo Excel ILRL: {missing_headers}. Verifique la configuración de O.T.")
                return None
            
            datos = []
            for idx, row in df.iterrows():
                try:
                    ot_parte_raw = row[ilrl_ot_header] if pd.notna(row[ilrl_ot_header]) else None
                    consecutivo_parte_raw = row[ilrl_serie_header] if pd.notna(row[ilrl_serie_header]) else None
                    punta_raw = row[ilrl_conector_header] if pd.notna(row[ilrl_conector_header]) else None
                    resultado_raw = row[ilrl_estado_header] if pd.notna(row[ilrl_estado_header]) else None
                    fecha_raw = row[ilrl_fecha_header] if pd.notna(row[ilrl_fecha_header]) else None
                    hora_raw = row[ilrl_hora_header] if pd.notna(row[ilrl_hora_header]) else None

                    ot_parte = str(ot_parte_raw).strip().upper() if ot_parte_raw is not None else None
                    consecutivo_parte = str(consecutivo_parte_raw).strip().zfill(4) if consecutivo_parte_raw is not None else None
                    punta = str(punta_raw).strip().upper() if punta_raw is not None else None
                    
                    resultado = str(resultado_raw).strip().upper() if resultado_raw is not None else '' 
                    if resultado not in ['PASS', 'FAIL']:
                        resultado = 'INVALID_RESULT' 
                    
                    if not all([ot_parte, consecutivo_parte, punta, fecha_raw, hora_raw]):
                        continue
                    
                    ot_parte_numerica = ot_parte.replace('JMO-', '')
                    serie_busqueda_ilrl = f"JMO-{ot_parte_numerica}-{consecutivo_parte}"
                    
                    fecha_str = ""
                    hora_str = ""
                    fecha_completa_dt = None

                    if isinstance(fecha_raw, datetime):
                        fecha_str = fecha_raw.strftime("%d/%m/%Y")
                    elif isinstance(fecha_raw, pd.Timestamp): 
                        fecha_str = fecha_raw.strftime("%d/%m/%Y")
                    else:
                        try:
                            if isinstance(fecha_raw, str) and '-' in fecha_raw and ':' in fecha_raw:
                                temp_dt = datetime.strptime(fecha_raw, "%Y-%m-%d %H:%M:%S")
                                fecha_str = temp_dt.strftime("%d/%m/%Y")
                            else: 
                                fecha_str = str(fecha_raw).split(' ')[0] 
                                if '-' in fecha_str: 
                                    parts = fecha_str.split('-')
                                    if len(parts) == 3:
                                        fecha_str = f"{parts[2]}/{parts[1]}/{parts[0]}"
                        except Exception:
                            fecha_str = "N/A"
                    
                    if isinstance(hora_raw, datetime):
                        hora_str = hora_raw.strftime("%H:%M:%S")
                    elif isinstance(hora_raw, pd.Timestamp): 
                        hora_str = hora_raw.strftime("%H:%M:%S")
                    else:
                        try:
                            time_match = re.search(r'(\d{1,2}:\d{2}:\d{2})', str(hora_raw))
                            if time_match:
                                hora_str = time_match.group(1)
                            else:
                                hora_str = str(hora_raw).split(' ')[0]
                        except Exception:
                            hora_str = "N/A"

                    try:
                        fecha_completa_dt = datetime.strptime(f"{fecha_str} {hora_str}", "%d/%m/%Y %H:%M:%S")
                    except ValueError:
                        try:
                            fecha_completa_dt = datetime.strptime(f"{fecha_str} {hora_str}", "%d/%m/%Y %H:%M")
                        except ValueError:
                            try:
                                fecha_completa_dt = datetime.strptime(fecha_str, "%d/%m/%Y")
                            except ValueError:
                                fecha_completa_dt = None

                    datos.append({
                        'serie_busqueda_ilrl': serie_busqueda_ilrl,
                        'ot': ot_parte,
                        'consecutivo': consecutivo_parte,
                        'punta': punta,
                        'resultado': resultado,
                        'fecha_str': fecha_str,
                        'hora_str': hora_str,
                        'fecha_completa_dt': fecha_completa_dt, 
                        'fecha_completa_display': f"{fecha_str} {hora_str}".strip(), 
                        'fibra_idx': idx  
                    })
                except Exception as e:
                    print(f"ILRL Debug - Fila {idx}: Error CRÍTICO procesando fila: {e}. Saltando.")
                    continue
            
            if not datos:
                return None
            
            all_cables_data = {}
            resultados_temp = defaultdict(lambda: defaultdict(lambda: {'mediciones': [], 'detalles': []}))
            latest_date_per_cable = defaultdict(lambda: datetime.min)

            for dato in datos:
                serie_key = dato['serie_busqueda_ilrl']
                punta_full = dato['punta']
                
                match = re.match(r'([AB])(\d*)', punta_full)
                if not match:
                    print(f"ILRL Debug: Formato de punta inválido '{punta_full}' para serie {serie_key}. Saltando.")
                    continue
                lado = match.group(1)
                conector_index = int(match.group(2)) if match.group(2) else 1
                
                resultados_temp[serie_key][(lado, conector_index)]['mediciones'].append(dato['resultado'])
                resultados_temp[serie_key][(lado, conector_index)]['detalles'].append({
                    'fibra': f"Fibra {len(resultados_temp[serie_key][(lado, conector_index)]['detalles']) + 1}",
                    'resultado': dato['resultado'],
                    'fecha': dato['fecha_str'],
                    'hora': dato['hora_str']
                })
                
                if dato['fecha_completa_dt'] and dato['fecha_completa_dt'] > latest_date_per_cable[serie_key]:
                    latest_date_per_cable[serie_key] = dato['fecha_completa_dt']

            for serie, conectores_data in resultados_temp.items():
                estado_cable_general = "APROBADO"
                detalles_puntas_para_display = defaultdict(lambda: {'estado': 'PASS', 'conectores': []})
                
                conectores_por_lado = defaultdict(list)
                for (lado, conector_index), data in conectores_data.items():
                    conectores_por_lado[lado].append({'index': conector_index, 'data': data})
                
                for lado in ['A', 'B']:
                    conectores_lado = sorted(conectores_por_lado[lado], key=lambda x: x['index'])
                    
                    if lado == 'A':
                        conectores_esperados = num_conectores_a
                        fibras_esperadas_por_conector = fibras_por_conector_a
                    else:
                        conectores_esperados = num_conectores_b
                        fibras_esperadas_por_conector = fibras_por_conector_b
                    
                    if len(conectores_lado) != conectores_esperados:
                        estado_cable_general = "RECHAZADO"
                        detalles_puntas_para_display[lado]['estado'] = f"RECHAZADO (FALTAN CONECTORES - esperados: {conectores_esperados}, encontrados: {len(conectores_lado)})"
                        
                    for conector in conectores_lado:
                        mediciones = conector['data']['mediciones']
                        detalles = conector['data']['detalles']
                        
                        estado_conector = "PASS"
                        if len(mediciones) != fibras_esperadas_por_conector:
                            estado_conector = f"RECHAZADO (FALTAN FIBRAS - {len(mediciones)}/{fibras_esperadas_por_conector})"
                            estado_cable_general = "RECHAZADO"
                        elif any(m != 'PASS' for m in mediciones):
                            estado_conector = "FAIL"
                            estado_cable_general = "RECHAZADO"
                        
                        detalles_puntas_para_display[lado]['conectores'].append({
                            'conector': f"{lado}{conector['index']}",
                            'estado': estado_conector,
                            'mediciones': detalles
                        })
                    
                    if detalles_puntas_para_display[lado]['estado'] == 'PASS' and any(c['estado'] != 'PASS' for c in detalles_puntas_para_display[lado]['conectores']):
                        detalles_puntas_para_display[lado]['estado'] = 'FAIL'
                        estado_cable_general = "RECHAZADO"

                latest_date_display = latest_date_per_cable[serie].strftime("%d/%m/%Y %H:%M:%S") if latest_date_per_cable[serie] != datetime.min else "N/A"
                
                all_cables_data[serie] = {
                    'status': estado_cable_general,
                    'date': latest_date_display,
                    'details': {
                        'estado': estado_cable_general, 
                        'detalles_puntas': detalles_puntas_para_display,
                    }
                }
            
            return all_cables_data
            
        except Exception as e:
            if isinstance(e, FileNotFoundError):
                messagebox.showerror("Error de Archivo", 
                                     f"El archivo de ILRL no se encontró en la ruta esperada: {ruta}. "
                                     f"Por favor, verifique la ruta configurada o si el archivo existe.")
            elif isinstance(e, PermissionError):
                messagebox.showerror("Error de Permisos", 
                                     f"Permiso denegado al intentar leer el archivo de ILRL: {ruta}. "
                                     f"Asegúrese de que el archivo no esté abierto en otro programa y tenga los permisos adecuados.")
            else:
                messagebox.showerror("Error de Lectura ILRL", f"No se pudo leer el archivo ILRL {os.path.basename(ruta)}: {e}")
            print(f"Error general leyendo archivo ILRL {os.path.basename(ruta)}: {e}")
            return None

    def leer_resultado_geo(self, ruta, cable_config_ot):
        """Método mejorado para leer resultados de geometría de cables MPO
        Retorna un diccionario con datos por cada cable encontrado en el archivo.
        Recibe la configuración específica del cable para esa OT.
        """
        try:
            if not os.path.exists(ruta):
                print(f"Geo Debug: Archivo no encontrado en leer_resultado_geo: {ruta}")
                return None
            
            if ruta.lower().endswith('.xls') and xlrd is None:
                messagebox.showerror("Error de Dependencia", 
                                     "La librería 'xlrd' es necesaria para leer archivos .xls. Por favor, instálala.")
                return None

            num_conectores_a = cable_config_ot.get('num_conectores_a', 1)
            num_conectores_b = cable_config_ot.get('num_conectores_b', 1)

            full_df = pd.read_excel(ruta, sheet_name="MT12", header=None)
            
            header_row_index = -1
            for idx, row_val in full_df[0].items(): 
                if isinstance(row_val, str) and row_val.strip().lower() == 'name':
                    header_row_index = idx
                    break
                if idx > 50:
                    break
            
            if header_row_index == -1:
                messagebox.showerror("Error de Lectura Geometría", 
                                     "No se encontró la fila con encabezado 'Name' en la columna A de la hoja 'MT12'. "
                                     "Verifique el formato del archivo de Geometría.")
                return None

            df = full_df.iloc[header_row_index:].copy()
            df.columns = df.iloc[0]
            df = df[1:].reset_index(drop=True)

            column_mapping = {
                'name': 'name',
                'pass/fail': 'pass/fail',
                'date & time': 'date_and_time',
                'date and time': 'date_and_time'
            }
            df.columns = [
                column_mapping.get(str(col).strip().lower(), str(col).strip().lower().replace(' ', '_').replace('&', 'and'))
                for col in df.columns
            ]
            
            required_columns = ['name', 'pass/fail', 'date_and_time'] 
            if not all(col in df.columns for col in required_columns):
                missing_cols = [col for col in required_columns if col not in df.columns]
                messagebox.showerror("Error de Columnas", 
                                     f"Las columnas esperadas ({', '.join(required_columns)}) no se encontraron o no se renombraron correctamente. "
                                     f"Faltan: {', '.join(missing_cols)}. Columnas detectadas: {df.columns.tolist()}"
                                     "Verifique el formato del archivo de Geometría (hoja 'MT12').")
                return None

            datos = []
            
            for idx, row in df.iterrows():
                try:
                    serie_completo_raw = row['name']
                    resultado_raw = row['pass/fail']
                    fecha_raw = row['date_and_time']
                    
                    if pd.isna(serie_completo_raw) or pd.isna(resultado_raw) or pd.isna(fecha_raw):
                        continue

                    serie_completo = str(serie_completo_raw).strip().upper()
                    resultado = str(resultado_raw).strip().upper()
                    
                    if resultado not in ['PASS', 'FAIL']:
                        continue

                    match = re.match(r'^(JMO\d{9})-(\d+)-([12])$', serie_completo)
                    if not match:
                        match = re.match(r'^(JMO\d{9})(\d{4})-([AB]|\d[AB])(?:-R)?$', serie_completo)
                        if not match:
                            continue
                        
                        ot_part_from_geo_with_jmo = match.group(1) 
                        consecutive_part_from_geo = match.group(2)
                        punta_raw_from_regex = match.group(3)
                        
                        if punta_raw_from_regex.endswith('A'):
                            punta_key = 'A'
                        elif punta_raw_from_regex.endswith('B'):
                            punta_key = 'B'
                        else:
                            continue
                            
                        serie_busqueda_geo = f"JMO{ot_part_from_geo_with_jmo.replace('JMO', '')}{consecutive_part_from_geo}"
                        connector_index = '1'
                    
                    else:
                        ot_part_from_geo_with_jmo = match.group(1) 
                        consecutive_part_from_geo = match.group(2)
                        side_number = match.group(3)
                        
                        punta_key = 'A' if side_number == '1' else 'B'
                        
                        serie_busqueda_geo = f"JMO{ot_part_from_geo_with_jmo.replace('JMO', '')}{consecutive_part_from_geo.zfill(4)}"
                        connector_index = '1'
                    
                    if not punta_key:
                        continue

                    fecha_str = ""
                    fecha_completa_dt = None
                    if isinstance(fecha_raw, datetime):
                        fecha_str = fecha_raw.strftime("%d/%m/%Y %H:%M:%S")
                    elif isinstance(fecha_raw, pd.Timestamp): 
                        fecha_str = fecha_raw.strftime("%d/%m/%Y %H:%M:%S")
                    else:
                        fecha_str = str(fecha_raw).strip() 

                    try:
                        fecha_completa_dt = datetime.strptime(fecha_str, "%d/%m/%Y %H:%M:%S")
                    except ValueError:
                        try:
                            fecha_completa_dt = datetime.strptime(fecha_str, "%d/%m/%Y %H:%M")
                        except ValueError:
                            try:
                                fecha_completa_dt = datetime.strptime(fecha_str.split(' ')[0], "%d/%m/%Y")
                            except ValueError:
                                fecha_completa_dt = None

                    datos.append({
                        'serie_busqueda_geo': serie_busqueda_geo,
                        'serie_completo_raw': serie_completo,
                        'punta_key': punta_key,
                        'resultado': resultado,
                        'fecha_completa_dt': fecha_completa_dt, 
                        'fecha_completa_display': fecha_str, 
                        'connector_index': connector_index
                    })
                except Exception as e:
                    print(f"Geo Debug - Fila {idx}: Error CRÍTICO procesando fila: {e}. Saltando.")
                    continue
            
            if not datos:
                return None
            
            all_cables_data = {}
            resultados_temp = defaultdict(lambda: {'A': defaultdict(lambda: None), 'B': defaultdict(lambda: None)}) 
            detalles_temp = defaultdict(lambda: {'A': [], 'B': []}) 
            latest_date_per_cable = defaultdict(lambda: datetime.min)

            for dato in datos:
                serie_key = dato['serie_busqueda_geo']
                current_date = dato['fecha_completa_dt']
                
                punta_key = dato['punta_key'] 
                connector_index = dato['connector_index']

                current_connector_entry_key = (punta_key, connector_index)
                
                if resultados_temp[serie_key][punta_key][connector_index] is None or \
                   (current_date and resultados_temp[serie_key][punta_key][connector_index].get('fecha_completa_dt') and \
                    current_date > resultados_temp[serie_key][punta_key][connector_index]['fecha_completa_dt']):
                    resultados_temp[serie_key][punta_key][connector_index] = dato
                    resultados_temp[serie_key][punta_key][connector_index]['conector_display_name'] = f"{punta_key}{connector_index}"

                detalles_temp[serie_key][punta_key].append({
                    'conector': f"{punta_key}{connector_index}",
                    'serie_completo': dato['serie_completo_raw'],
                    'resultado': dato['resultado'],
                    'fecha': dato['fecha_completa_display']
                })

                if current_date and current_date > latest_date_per_cable[serie_key]:
                    latest_date_per_cable[serie_key] = current_date


            for serie, puntas_data in resultados_temp.items():
                estado_cable_general = "APROBADO" 
                detalles_puntas_para_display = {}
                
                for lado in ['A', 'B']:
                    conectores_medidos = puntas_data.get(lado, {})
                    
                    num_conectores_esperados = num_conectores_a if lado == 'A' else num_conectores_b

                    if len(conectores_medidos) != num_conectores_esperados:
                        estado_cable_general = "RECHAZADO"
                        detalles_puntas_para_display[lado] = {
                            'estado': f"RECHAZADO (FALTAN CONECTORES - esperados: {num_conectores_esperados}, encontrados: {len(conectores_medidos)})",
                            'conectores_medidos': [entry['conector_display_name'] for entry in conectores_medidos.values() if entry],
                            'mediciones': detalles_temp[serie].get(lado, [])
                        }
                    else:
                        estado_punta_global = "PASS"
                        for conector_index, entry in conectores_medidos.items():
                            if entry is None or entry['resultado'] != 'PASS':
                                estado_punta_global = "FAIL"
                                break
                        
                        if estado_punta_global == "FAIL":
                            estado_cable_general = "RECHAZADO"
                            detalles_puntas_para_display[lado] = {
                                'estado': "FAIL",
                                'mediciones': detalles_temp[serie].get(lado, [])
                            }
                        else:
                            detalles_puntas_para_display[lado] = {
                                'estado': "PASS",
                                'mediciones': detalles_temp[serie].get(lado, [])
                            }
                
                latest_date_display = latest_date_per_cable[serie].strftime("%d/%m/%Y %H:%M:%S") if latest_date_per_cable[serie] != datetime.min else "N/A"

                all_cables_data[serie] = {
                    'status': estado_cable_general,
                    'date': latest_date_display,
                    'details': {
                        'estado': estado_cable_general, 
                        'detalles_puntas': detalles_puntas_para_display,
                        'num_conectores_a': num_conectores_a,
                        'num_conectores_b': num_conectores_b,
                        'conectores_encontrados': {k: len(v) for k, v in puntas_data.items()}
                    }
                }
            
            return all_cables_data
            
        except Exception as e:
            if isinstance(e, FileNotFoundError):
                messagebox.showerror("Error de Archivo", 
                                     f"El archivo de Geometría no se encontró en la ruta esperada: {ruta}. "
                                     f"Por favor, verifique la ruta configurada o si el archivo existe.")
            elif isinstance(e, PermissionError):
                messagebox.showerror("Error de Permisos", 
                                     f"Permiso denegado al intentar leer el archivo de Geometría: {ruta}. "
                                     f"Asegúrese de que el archivo no esté abierto en otro programa y tenga los permisos adecuados.")
            else:
                messagebox.showerror("Error de Lectura ILRL", f"No se pudo leer el archivo ILRL {os.path.basename(ruta)}: {e}")
            print(f"Error general leyendo archivo ILRL {os.path.basename(ruta)}: {e}")
            return None

    def buscar_archivos_ilrl(self, ot_para_buscar):
        """
        Busca archivos ILRL para la OT especificada de manera optimizada.
        ot_para_buscar debe ser en formato 'JMO-XXXXXXXXX'.
        """
        archivos_encontrados = []
        ruta_ot = os.path.join(self.ruta_base_ilrl, ot_para_buscar)
        
        if os.path.isdir(ruta_ot):
            for f in os.listdir(ruta_ot):
                if f.lower().endswith(('.xlsx', '.xls')) and not f.startswith('~$'):
                    if f.lower().startswith(ot_para_buscar.lower()):
                        archivos_encontrados.append(os.path.join(ruta_ot, f))
        
        return archivos_encontrados

    def buscar_archivos_geo(self, ot_para_buscar):
        """
        Busca archivos de Geometría para la OT especificada de manera optimizada.
        ot_para_buscar debe ser en formato 'JMO-XXXXXXXXX'.
        """
        archivos = []
        ot_limpia_numerica = ot_para_buscar.replace('JMO-', '').upper() 
    
        if os.path.exists(self.ruta_base_geo):
            for f in os.listdir(self.ruta_base_geo):
                nombre_archivo_base_sin_extension = os.path.splitext(f)[0].upper()
                nombre_archivo_limpia_numerica = nombre_archivo_base_sin_extension.replace('JMO-', '').replace('JMO', '')
                
                print(f"DEBUG_GEO_SEARCH: Comparando '{ot_limpia_numerica}' con '{nombre_archivo_limpia_numerica}' para archivo '{f}'")

                if f.lower().endswith(('.xlsx', '.xls')) and not f.startswith('~$') and ot_limpia_numerica == nombre_archivo_limpia_numerica:
                    archivos.append(os.path.join(self.ruta_base_geo, f))
        return archivos
        
    def leer_resultado_polaridad(self, ruta_archivo):
        """Lee un archivo de polaridad (.xlsx) y extrae los datos relevantes del formato específico."""
        try:
            if not os.path.exists(ruta_archivo):
                print(f"Polaridad Debug: Archivo no encontrado en leer_resultado_polaridad: {ruta_archivo}")
                return None
            
            df = pd.read_excel(ruta_archivo, sheet_name=0, header=None)
            
            ot_number_cell = df.iloc[2, 1] if df.shape[0] > 2 and df.shape[1] > 1 else None
            serial_number_part_cell = df.iloc[3, 1] if df.shape[0] > 3 and df.shape[1] > 1 else None
            status_cell = df.iloc[12, 1] if df.shape[0] > 12 and df.shape[1] > 1 else None
            date_cell = df.iloc[13, 1] if df.shape[0] > 13 and df.shape[1] > 1 else None
            full_serial_from_excel_cell = df.iloc[1, 1] if df.shape[0] > 1 and df.shape[1] > 1 else None
            
            ot_number = str(ot_number_cell).strip().replace('JMO','').zfill(9) if pd.notna(ot_number_cell) else None
            serial_number_part = str(serial_number_part_cell).strip().zfill(4) if pd.notna(serial_number_part_cell) else None
            status = str(status_cell).strip().upper() if pd.notna(status_cell) else None
            date_raw = str(date_cell).strip() if pd.notna(date_cell) else None
            
            if not all([ot_number, serial_number_part, status, date_raw]):
                print("Polaridad Debug: Datos clave faltantes en el archivo.")
                return None
            
            date_dt = None
            try:
                date_dt = datetime.strptime(date_raw, "%Y-%m-%d %H:%M:%S")
                date_display = date_dt.strftime("%d/%m/%Y %H:%M:%S")
            except ValueError:
                date_display = date_raw
                date_dt = datetime.now()

            file_name = os.path.basename(ruta_archivo)
            match_file_name = re.search(r'JMO(\d{9})##(\d{4})', file_name)
            if match_file_name:
                ot_part_from_file = match_file_name.group(1)
                consecutive_part_from_file = match_file_name.group(2)
                full_serial = f"JMO{ot_part_from_file}{consecutive_part_from_file}"
            else:
                full_serial = None
            
            result = {
                'ot_number': ot_number,
                'serial_number_part': serial_number_part,
                'full_serial': full_serial,
                'status': status,
                'date': date_display,
                'date_dt': date_dt,
                'file_name': os.path.basename(ruta_archivo)
            }
            return result
        
        except Exception as e:
            messagebox.showerror("Error de Lectura Polaridad", 
                                 f"No se pudo leer el archivo de polaridad {os.path.basename(ruta_archivo)}: {e}")
            print(f"Polaridad Debug: Error leyendo archivo: {e}")
            return None

    def buscar_archivos_polaridad(self, ot_para_buscar):
        """
        Busca archivos de polaridad para la OT especificada, buscando en las subcarpetas
        'PASS' y 'FAIL'.
        """
        archivos = []
        ot_limpia_numerica = ot_para_buscar.replace('JMO-', '').upper() 
        serial_completo = self.serie_entry.get().strip() 
        serial_consecutivo = serial_completo[9:]

        # Se espera una estructura de carpetas: ruta_base_polaridad/OT/PASS o FAIL
        ruta_ot_base = os.path.join(self.ruta_base_polaridad, f"JMO{ot_limpia_numerica}")
        
        if os.path.exists(ruta_ot_base):
            for root, dirs, files in os.walk(ruta_ot_base):
                for f in files:
                    # El formato del archivo es JMOXXXXXXXXX##YYYY##...xlsx
                    if f.lower().endswith('.xlsx') and not f.startswith('~$'):
                        match = re.search(r'JMO(\d{9})##(\d{4})', f)
                        if match:
                            ot_del_archivo = match.group(1)
                            consecutivo_del_archivo = match.group(2)
                            
                            if ot_del_archivo == ot_limpia_numerica and consecutivo_del_archivo == serial_consecutivo:
                                archivos.append(os.path.join(root, f))
                                print(f"Polaridad Debug: Archivo encontrado: {f}")
        return archivos


    def verificar_cable(self):
        ot_input_raw = self.ot_entry.get().strip()
        serial_input_raw = self.serie_entry.get().strip()
        
        self.last_ilrl_analysis_data = None
        self.last_geo_analysis_data = None
        self.last_polaridad_analysis_data = None
        self.btn_ver_detalles.config(state=tk.DISABLED)
        
        self.ruta_ilrl_label.config(text=f"📂 Ruta ILRL: {self.ruta_base_ilrl}")
        self.ruta_geo_label.config(text=f"📂 Ruta Geometría: {self.ruta_base_geo}")
        self.ruta_polaridad_label.config(text=f"📂 Ruta Polaridad: {self.ruta_base_polaridad}")
        
        if not ot_input_raw or not serial_input_raw:
            self.resultado_text.config(state=tk.NORMAL)
            self.resultado_text.delete(1.0, tk.END)
            self.resultado_text.insert(tk.END, "Por favor, ingrese OT y Número de Serie para verificar.", "normal")
            self.resultado_text.tag_unbind("ilrl_click", "<Button-1>")
            self.resultado_text.tag_unbind("geo_click", "<Button-1>")
            self.resultado_text.config(state=tk.DISABLED)
            return

        ot_input_cleaned_for_comp = ot_input_raw.upper().replace('JMO-', '')
        
        if not re.match(r'^\d{13}$', serial_input_raw):
            mensaje_error_formato = (
                "⚠️ Formato de número de serie incorrecto. "
                "El número de serie debe ser de 13 dígitos numéricos (ej. 2507000070013)."
            )
            messagebox.showwarning("Formato Incorrecto", mensaje_error_formato)
            self.resultado_text.config(state=tk.NORMAL)
            self.resultado_text.delete(1.0, tk.END)
            self.resultado_text.insert(tk.END, mensaje_error_formato + "\n\nVerificación no realizada.", "rojo")
            self.resultado_text.config(state=tk.DISABLED)
            return

        ot_part_from_serial = serial_input_raw[:9] 
        consecutive_part = serial_input_raw[9:].zfill(4) 

        print(f"DEBUG_MATCH: ot_input_cleaned_for_comp = '{ot_input_cleaned_for_comp}' (len: {len(ot_input_cleaned_for_comp)})")
        print(f"DEBUG_MATCH: ot_part_from_serial = '{ot_part_from_serial}' (len: {len(ot_part_from_serial)})")
        print(f"DEBUG_MATCH: Comparison result: {ot_input_cleaned_for_comp != ot_part_from_serial}")

        if ot_input_cleaned_for_comp != ot_part_from_serial:
            messagebox.showwarning(
                "Error de Coincidencia",
                f"La Orden de Trabajo ingresada ('{ot_input_raw}') no coincide con la parte inicial "
                f"del Número de Serie ('{ot_part_from_serial}').\n\n"
                "Verifique que los datos sean correctos. No se realizará la verificación ni el registro."
            )
            self.resultado_text.config(state=tk.NORMAL)
            self.resultado_text.delete(1.0, tk.END)
            self.resultado_text.insert(tk.END, "⚠️ ERROR: La OT y el Número de Serie no coinciden.\n"
                                         "Por favor, verifique los datos.", "rojo")
            self.resultado_text.config(state=tk.DISABLED)
            return
        
        ot_input_cleaned = f"JMO-{ot_input_cleaned_for_comp}"
        ot_for_file_search = ot_input_cleaned

        cable_config_ot = self._cargar_ot_configuration(ot_input_cleaned)
        if not cable_config_ot:
            messagebox.showwarning(
                "Configuración Requerida",
                f"Primero debe configurar el cable para la OT {ot_input_cleaned}.\n"
                "Especifique la cantidad de conectores, fibras por conector, y detalles del dibujo."
            )
            self.configurar_ot()
            cable_config_ot = self._cargar_ot_configuration(ot_input_cleaned)
            if not cable_config_ot:
                return

        ilrl_lookup_key = f"{ot_input_cleaned}-{consecutive_part}" 
        geo_lookup_key = f"JMO{ot_part_from_serial}{consecutive_part}"

        print(f"OT para búsqueda de archivos: {ot_for_file_search}")
        print(f"Clave de búsqueda ILRL: {ilrl_lookup_key}")
        print(f"Clave de búsqueda Geometría: {geo_lookup_key}")
        print(f"Configuración del cable MPO para OT {ot_input_cleaned}: {cable_config_ot}")

        archivos_ilrl = self.buscar_archivos_ilrl(ot_for_file_search)
        ilrl_data_all_cables_from_files = {}
        self.last_ilrl_file_path = None 

        if archivos_ilrl:
            for archivo_ilrl_path in archivos_ilrl:
                current_ilrl_file_data = self.leer_resultado_ilrl(archivo_ilrl_path, cable_config_ot)
                if current_ilrl_file_data:
                    for cable_key, cable_info in current_ilrl_file_data.items():
                        if cable_key not in ilrl_data_all_cables_from_files:
                            ilrl_data_all_cables_from_files[cable_key] = cable_info
                            ilrl_data_all_cables_from_files[cable_key]['_file_path'] = archivo_ilrl_path 
                        else:
                            existing_date_str = ilrl_data_all_cables_from_files[cable_key]['date']
                            current_date_str = cable_info['date']
                            try:
                                existing_date = datetime.strptime(existing_date_str, "%d/%m/%Y %H:%M:%S")
                                current_date = datetime.strptime(current_date_str, "%d/%m/%Y %H:%M:%S")
                                if current_date > existing_date:
                                    ilrl_data_all_cables_from_files[cable_key] = cable_info
                                    ilrl_data_all_cables_from_files[cable_key]['_file_path'] = archivo_ilrl_path
                            except ValueError:
                                pass
        
        ilrl_cable_result = ilrl_data_all_cables_from_files.get(ilrl_lookup_key)
        
        resultado_ilrl = "NO ENCONTRADO"
        fecha_ilrl = None
        ilrl_detalles_para_db = None
        
        if ilrl_cable_result:
            resultado_ilrl = ilrl_cable_result['status']
            fecha_ilrl = ilrl_cable_result['date']
            ilrl_detalles_para_db = ilrl_cable_result['details']
            self.last_ilrl_file_path = ilrl_cable_result.get('_file_path') 
        
        self.last_ilrl_analysis_data = ilrl_detalles_para_db

        archivos_geo = self.buscar_archivos_geo(ot_for_file_search)
        geo_data_all_cables_from_files = {}
        self.last_geo_file_path = None 

        if archivos_geo:
            for archivo_geo_path in archivos_geo:
                current_geo_file_data = self.leer_resultado_geo(archivo_geo_path, cable_config_ot)
                if current_geo_file_data:
                    for cable_key, cable_info in current_geo_file_data.items():
                        if cable_key not in geo_data_all_cables_from_files:
                            geo_data_all_cables_from_files[cable_key] = cable_info
                            geo_data_all_cables_from_files[cable_key]['_file_path'] = archivo_geo_path 
                        else:
                            existing_date_str = geo_data_all_cables_from_files[cable_key]['date']
                            current_date_str = cable_info['date']
                            try:
                                existing_date = datetime.strptime(existing_date_str, "%d/%m/%Y %H:%M:%S")
                                current_date = datetime.strptime(current_date_str, "%d/%m/%Y %H:%M:%S")
                                if current_date > existing_date:
                                    geo_data_all_cables_from_files[cable_key] = cable_info
                                    geo_data_all_cables_from_files[cable_key]['_file_path'] = archivo_ilrl_path
                            except ValueError:
                                pass
        
        geo_cable_result = geo_data_all_cables_from_files.get(geo_lookup_key)
        
        resultado_geo = "NO ENCONTRADO"
        fecha_geo = None
        geo_detalles_para_db = None
        
        if geo_cable_result:
            resultado_geo = geo_cable_result['status']
            fecha_geo = geo_cable_result['date']
            geo_detalles_para_db = geo_cable_result['details']
            self.last_geo_file_path = geo_cable_result.get('_file_path') 

        self.last_geo_analysis_data = geo_detalles_para_db

        # --- Lógica de Polaridad ---
        archivos_polaridad = self.buscar_archivos_polaridad(ot_for_file_search)
        polaridad_data_all_cables = {}
        
        for archivo_pol_path in archivos_polaridad:
            polaridad_result = self.leer_resultado_polaridad(archivo_pol_path)
            if polaridad_result and 'full_serial' in polaridad_result:
                polaridad_lookup_key = polaridad_result['full_serial']
                current_date = polaridad_result.get('date_dt')
                
                if polaridad_lookup_key not in polaridad_data_all_cables or (current_date and current_date > polaridad_data_all_cables[polaridad_lookup_key].get('date_dt')):
                    polaridad_data_all_cables[polaridad_lookup_key] = polaridad_result
        
        polaridad_lookup_key = geo_lookup_key
        polaridad_cable_result = polaridad_data_all_cables.get(polaridad_lookup_key)

        resultado_polaridad = "NO ENCONTRADO"
        fecha_polaridad = None
        polaridad_detalles_para_db = None
        if polaridad_cable_result:
            resultado_polaridad = polaridad_cable_result['status']
            fecha_polaridad = polaridad_cable_result['date']
            polaridad_detalles_para_db = polaridad_cable_result
            self.last_polaridad_file_path = polaridad_cable_result['file_name']
        self.last_polaridad_analysis_data = polaridad_detalles_para_db

        # --- Determinar el estado final ---
        overall_status_db = "NO ENCONTRADO"
        
        ilrl_pass = resultado_ilrl == "APROBADO"
        geo_pass = resultado_geo == "APROBADO"
        polaridad_pass = resultado_polaridad == "PASS"

        if ilrl_pass and geo_pass and polaridad_pass:
            overall_status_db = "APROBADO"
        elif resultado_ilrl != "NO ENCONTRADO" or resultado_geo != "NO ENCONTRADO" or resultado_polaridad != "NO ENCONTRADO":
             if ilrl_pass and geo_pass and not polaridad_pass:
                overall_status_db = "RECHAZADO (Falla Polaridad)"
             elif ilrl_pass and not geo_pass and polaridad_pass:
                overall_status_db = "RECHAZADO (Falla Geometría)"
             elif not ilrl_pass and geo_pass and polaridad_pass:
                overall_status_db = "RECHAZADO (Falla ILRL)"
             else:
                overall_status_db = "RECHAZADO"

        
        # Registrar el resultado de la verificación
        self._log_verification_result(
            serial_number=serial_input_raw, 
            ot_number=ot_input_cleaned,
            overall_status=overall_status_db,
            ilrl_status=resultado_ilrl,
            ilrl_date=fecha_ilrl,
            ilrl_details=ilrl_detalles_para_db,
            geo_status=resultado_geo,
            geo_date=fecha_geo,
            geo_details=geo_detalles_para_db,
            polaridad_status=resultado_polaridad,
            polaridad_date=fecha_polaridad,
            polaridad_details=polaridad_detalles_para_db
        )

        # Mostrar resultados en la interfaz
        self.resultado_text.config(state=tk.NORMAL)
        self.resultado_text.delete(1.0, tk.END)
        
        self.resultado_text.tag_unbind("ilrl_click", "<Button-1>")
        self.resultado_text.tag_unbind("geo_click", "<Button-1>")
        self.resultado_text.tag_unbind("polaridad_click", "<Button-1>")

        self.resultado_text.insert(tk.END, f"🔍 Resultados para cable {serial_input_raw} en OT {ot_input_cleaned}:\n\n", "header")
        
        self.resultado_text.insert(tk.END, "📊 ILRL: ", "bold")
        if resultado_ilrl != "NO ENCONTRADO":
            color_tag = "verde" if ilrl_pass else "rojo"
            self.resultado_text.insert(tk.END, f"{resultado_ilrl}", (color_tag, "ilrl_click"))
            if fecha_ilrl:
                self.resultado_text.insert(tk.END, f" (📅 {fecha_ilrl})", "normal")
            self.resultado_text.tag_bind("ilrl_click", "<Button-1>", lambda e: self.mostrar_detalles_ilrl(self.last_ilrl_analysis_data))
            self.resultado_text.tag_config("ilrl_click", underline=1)
        else:
            self.resultado_text.insert(tk.END, "NO ENCONTRADO", "orange")
        self.resultado_text.insert(tk.END, "\n")
        
        self.resultado_text.insert(tk.END, "📐 Geometría: ", "bold")
        if resultado_geo != "NO ENCONTRADO":
            color_tag = "verde" if geo_pass else "rojo"
            self.resultado_text.insert(tk.END, f"{resultado_geo}", (color_tag, "geo_click"))
            if fecha_geo:
                self.resultado_text.insert(tk.END, f" (📅 {fecha_geo})", "normal")
            self.resultado_text.tag_bind("geo_click", "<Button-1>", lambda e: self.mostrar_detalles_geo(self.last_geo_analysis_data))
            self.resultado_text.tag_config("geo_click", underline=1)
        else:
            self.resultado_text.insert(tk.END, "NO ENCONTRADA", "orange")
        self.resultado_text.insert(tk.END, "\n")
        
        self.resultado_text.insert(tk.END, "🔀 Polaridad: ", "bold")
        if resultado_polaridad != "NO ENCONTRADO":
            color_tag = "verde" if polaridad_pass else "rojo"
            self.resultado_text.insert(tk.END, f"{resultado_polaridad}", (color_tag, "polaridad_click"))
            if fecha_polaridad:
                self.resultado_text.insert(tk.END, f" (📅 {fecha_polaridad})", "normal")
            self.resultado_text.tag_bind("polaridad_click", "<Button-1>", lambda e: self.mostrar_detalles_polaridad(self.last_polaridad_analysis_data))
            self.resultado_text.tag_config("polaridad_click", underline=1)
        else:
            self.resultado_text.insert(tk.END, "NO ENCONTRADA", "orange")
        self.resultado_text.insert(tk.END, "\n\n")

        self.resultado_text.insert(tk.END, "🏁 ESTADO FINAL: ", "bold")
        color = "verde" if overall_status_db == "APROBADO" else "red" if "RECHAZADO" in overall_status_db else "orange"
        self.resultado_text.insert(tk.END, f"{overall_status_db}\n", color)
            
        if overall_status_db == "APROBADO":
            self.resultado_text.insert(tk.END, "✅ ¡El cable cumple con todos los requisitos!\n", "verde")
        elif "RECHAZADO" in overall_status_db:
            self.resultado_text.insert(tk.END, "❌ El cable no cumple con los requisitos\n", "rojo")
            try:
                winsound.PlaySound("SystemExclamation", winsound.SND_ALIAS)
            except Exception as e:
                print(f"Error al reproducir sonido: {e}")
        else:
            self.resultado_text.insert(tk.END, "⚠️ No se pudo verificar completamente el cable.\n", "orange")

        self.resultado_text.config(state=tk.DISABLED)
        
        if self.last_ilrl_analysis_data or self.last_geo_analysis_data or self.last_polaridad_analysis_data:
            self.btn_ver_detalles.config(state=tk.NORMAL)

    def mostrar_detalles_polaridad(self, data=None):
        """Muestra una ventana con los detalles de la prueba de polaridad."""
        details_to_show = data if data else self.last_polaridad_analysis_data
        
        if not details_to_show:
            messagebox.showinfo("Detalles Polaridad", "No hay datos de Polaridad para mostrar. Realice una verificación primero.")
            return

        detalles_window = tk.Toplevel(self.root)
        detalles_window.title("Detalles de Verificación Polaridad")
        detalles_window.geometry("500x300")
        detalles_window.transient(self.root)
        detalles_window.grab_set()
        
        main_frame = ttk.Frame(detalles_window)
        main_frame.pack(fill=tk.BOTH, expand=True)
        canvas = tk.Canvas(main_frame, highlightthickness=0)
        yscrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=yscrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        yscrollbar.pack(side="right", fill="y")
        
        def _on_pol_details_canvas_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        canvas.bind("<Configure>", _on_pol_details_canvas_configure)
        canvas.bind_all("<MouseWheel>", self._on_mouse_wheel)

        frame = ttk.Frame(scrollable_frame, padding=(20, 20))
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="📋 Detalles de la Prueba de Polaridad", font=("Arial", 12, "bold")).pack(anchor="w", pady=(0, 10))
        
        pol_status = details_to_show.get('status', 'N/A')
        pol_status_color = "green" if pol_status == "PASS" else "red" if pol_status == "FAIL" else "orange"
        
        ttk.Label(frame, text=f"• Estado: {pol_status}", font=("Arial", 10, "bold"), foreground=pol_status_color).pack(anchor="w", pady=2)
        ttk.Label(frame, text=f"• O.T. del Reporte: {details_to_show.get('ot_number', 'N/A')}", font=("Arial", 10)).pack(anchor="w", pady=2)
        ttk.Label(frame, text=f"• Número de Serie: {details_to_show.get('serial_number_part', 'N/A')}", font=("Arial", 10)).pack(anchor="w", pady=2)
        ttk.Label(frame, text=f"• Número de Serie Completo: {details_to_show.get('full_serial', 'N/A')}", font=("Arial", 10)).pack(anchor="w", pady=2)
        ttk.Label(frame, text=f"• Fecha y Hora de Prueba: {details_to_show.get('date', 'N/A')}", font=("Arial", 10)).pack(anchor="w", pady=2)

        ttk.Button(frame, text="Cerrar", command=detalles_window.destroy).pack(pady=20)
        
        detalles_window.mainloop()

    def mostrar_detalles_ilrl(self, data=None):
        """Muestra una ventana con los detalles completos del análisis ILRL para MPO.
        Ahora muestra el desglose por conector.
        """
        details_to_show = data if data else self.last_ilrl_analysis_data

        if not details_to_show:
            messagebox.showinfo("Detalles ILRL", "No hay datos de ILRL para mostrar detalles. Realice una verificación primero.")
            return

        detalles_window = tk.Toplevel(self.root)
        detalles_window.title("Detalles de Verificación ILRL (MPO)")
        detalles_window.geometry("900x700")
        detalles_window.transient(self.root)
        detalles_window.grab_set()

        main_frame = ttk.Frame(detalles_window)
        main_frame.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(main_frame, highlightthickness=0)
        yscrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        xscrollbar = ttk.Scrollbar(main_frame, orient="horizontal", command=canvas.xview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_window_id = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=yscrollbar.set, xscrollcommand=xscrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        yscrollbar.pack(side="right", fill="y")
        xscrollbar.pack(side="bottom", fill="x")

        def _on_ilrl_details_canvas_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas_width = event.width
            canvas.itemconfigure(scrollable_window_id, width=canvas_width)

        canvas.bind("<Configure>", _on_ilrl_details_canvas_configure)
        canvas.bind_all("<MouseWheel>", self._on_mouse_wheel)

        content_frame = ttk.Frame(scrollable_frame, style="Detalles.TFrame", padding=(20, 20))
        content_frame.pack(fill=tk.BOTH, expand=True)

        style = ttk.Style()
        style.theme_use('default') 
        BG_COLOR = "#F7F7F7"
        ACCENT_BLUE = "#007BFF"
        TEXT_COLOR = "#333333"
        style.configure("Detalles.TFrame", background=BG_COLOR)
        style.configure("Detalles.TLabel", background=BG_COLOR, foreground=TEXT_COLOR)
        style.configure("Detalles.Treeview", background="#FFFFFF", fieldbackground="#FFFFFF", foreground=TEXT_COLOR)
        style.configure("Detalles.Treeview.Heading", background=ACCENT_BLUE, foreground="white", font=('Arial', 9, 'bold'))
        style.map("Detalles.Treeview", background=[('selected', ACCENT_BLUE)])
        
        ttk.Label(content_frame, 
                text="📊 Detalles Completos de Verificación ILRL (MPO)", 
                font=("Arial", 12, "bold"), 
                style="Detalles.TLabel").pack(anchor="w", pady=(0, 15))

        ttk.Label(content_frame, 
                text=f"🔹 Estado General ILRL: {details_to_show.get('estado', 'N/A')}", 
                font=("Arial", 10), 
                style="Detalles.TLabel").pack(anchor="w", pady=(0, 5))
        
        ot_data_config = self._cargar_ot_configuration(self.ot_entry.get().strip().upper())
        if ot_data_config:
             ttk.Label(content_frame, 
                  text=f"🔹 Lado A: {ot_data_config.get('num_conectores_a', 'N/A')} Conectores, {ot_data_config.get('fibers_per_connector_a', 'N/A')} Fibras/Conector (Total: {ot_data_config.get('num_conectores_a', 0) * ot_data_config.get('fibers_per_connector_a', 0)} Fibras)", 
                  font=("Arial", 10), 
                  style="Detalles.TLabel").pack(anchor="w", pady=(0, 5))
             ttk.Label(content_frame, 
                  text=f"🔹 Lado B: {ot_data_config.get('num_conectores_b', 'N/A')} Conectores, {ot_data_config.get('fibers_per_connector_b', 'N/A')} Fibras/Conector (Total: {ot_data_config.get('num_conectores_b', 0) * ot_data_config.get('fibers_per_connector_b', 0)} Fibras)", 
                  font=("Arial", 10), 
                  style="Detalles.TLabel").pack(anchor="w", pady=(0, 5))


        for punta in ['A', 'B']:
            punta_data = details_to_show.get('detalles_puntas', {}).get(punta, {})
            estado_punta_display = punta_data.get('estado', 'NO ENCONTRADO')
            color = "green" if estado_punta_display == "PASS" else "red" if "RECHAZADO" in estado_punta_display or "FAIL" in estado_punta_display or "FALTAN" in estado_punta_display else "orange"
            
            ttk.Label(content_frame, 
                    text=f"• Lado {punta}: {estado_punta_display}", 
                    font=("Arial", 10, "bold"), 
                    foreground=color,
                    style="Detalles.TLabel").pack(anchor="w", pady=(5, 5))

            for conector in punta_data.get('conectores', []):
                tree_frame = ttk.Frame(content_frame, style="Detalles.TFrame")
                tree_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 5))

                estado_conector_display = conector['estado']
                conector_color = "green" if estado_conector_display == "PASS" else "red" if "FAIL" in estado_conector_display or "FALTAN" in estado_conector_display else "orange"
                ttk.Label(tree_frame, text=f"  - Conector {conector['conector']}: {estado_conector_display}", font=("Arial", 10, "bold"), foreground=conector_color, style="Detalles.TLabel").pack(anchor="w")

                tree = ttk.Treeview(
                    tree_frame,
                    columns=("Fibra", "Resultado", "Fecha", "Hora"),
                    show="headings",
                    height=min(6, len(conector['mediciones'])),
                    style="Detalles.Treeview"
                )

                tree.heading("Fibra", text="Fibra", anchor=tk.W)
                tree.heading("Resultado", text="Resultado", anchor=tk.W)
                tree.heading("Fecha", text="Fecha", anchor=tk.W)
                tree.heading("Hora", text="Hora", anchor=tk.W)

                tree.column("Fibra", width=50, stretch=tk.NO, anchor=tk.W)
                tree.column("Resultado", width=80, stretch=tk.NO, anchor=tk.W)
                tree.column("Fecha", width=100, stretch=tk.NO, anchor=tk.W)
                tree.column("Hora", width=80, stretch=tk.NO, anchor=tk.W)
                
                tree_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview, style="Vertical.TScrollbar")
                tree.configure(yscrollcommand=tree_scroll.set)
                tree.pack(side="left", fill=tk.BOTH, expand=True)
                tree_scroll.pack(side="right", fill="y")
                
                tree.tag_configure('PASS', foreground='green')
                tree.tag_configure('FAIL', foreground='red')
                tree.tag_configure('INVALID_RESULT', foreground='orange')

                for fibra in conector['mediciones']:
                    result_tag = fibra.get('resultado', '').upper()
                    if result_tag not in ['PASS', 'FAIL']:
                        result_tag = 'INVALID_RESULT'
                    
                    tree.insert(
                        "", 
                        tk.END, 
                        values=(
                            fibra.get('fibra', 'N/A'),
                            fibra.get('resultado', 'N/A'),
                            fibra.get('fecha', 'N/A'),
                            fibra.get('hora', 'N/A')
                        ), 
                        tags=(result_tag,)
                    )

        btn_frame = ttk.Frame(content_frame, style="Detalles.TFrame")
        btn_frame.pack(fill=tk.X, pady=(15, 0))

        ttk.Button(btn_frame, 
                text="Cerrar", 
                command=detalles_window.destroy).pack(pady=10)

        detalles_window.mainloop()

    def mostrar_detalles_geo(self, data=None):
        """Muestra una ventana con los detalles completos del análisis de Geometría para MPO"""
        details_to_show = data if data else self.last_geo_analysis_data
        
        if not details_to_show:
            messagebox.showinfo("Detalles Geometría", "No hay datos de Geometría para mostrar detalles. Realice una verificación primero.")
            return

        detalles_window = tk.Toplevel(self.root)
        detalles_window.title("Detalles de Verificación Geometría (MPO)")
        detalles_window.geometry("600x400")
        detalles_window.transient(self.root)
        detalles_window.grab_set()

        main_frame = ttk.Frame(detalles_window)
        main_frame.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(main_frame, highlightthickness=0)
        yscrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        xscrollbar = ttk.Scrollbar(main_frame, orient="horizontal", command=canvas.xview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_window_id = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=yscrollbar.set, xscrollcommand=xscrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        yscrollbar.pack(side="right", fill="y")
        xscrollbar.pack(side="bottom", fill="x")

        def _on_geo_details_canvas_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas_width = event.width
            canvas.itemconfigure(scrollable_window_id, width=canvas_width)

        canvas.bind("<Configure>", _on_geo_details_canvas_configure)
        canvas.bind_all("<MouseWheel>", self._on_mouse_wheel)
        
        frame = ttk.Frame(scrollable_frame, padding=(20, 20), style="GeoDetalles.TFrame")
        frame.pack(fill=tk.BOTH, expand=True)

        BG_COLOR = "#F7F7F7"
        ACCENT_BLUE = "#007BFF"
        TEXT_COLOR = "#333333"

        style = ttk.Style()
        style.configure("GeoDetalles.TFrame", background=BG_COLOR)
        style.configure("GeoDetalles.TLabel", background=BG_COLOR, foreground=TEXT_COLOR)
        style.configure("GeoDetalles.Treeview", background="#FFFFFF", fieldbackground="#FFFFFF", foreground=TEXT_COLOR)
        style.map("GeoDetalles.Treeview", background=[('selected', ACCENT_BLUE)])
        style.configure("GeoDetalles.Treeview.Heading", background=ACCENT_BLUE, foreground="white", font=('Arial', 9, 'bold'))

        ttk.Label(frame, text="📋 Detalles de Verificación Geometría (MPO)", font=("Arial", 12, "bold"), style="GeoDetalles.TLabel").pack(anchor="w", pady=(0, 15))
        
        ttk.Label(frame, 
                  text=f"🔹 Estado General Geometría: {details_to_show.get('estado', 'N/A')}", 
                  font=("Arial", 10), 
                  style="GeoDetalles.TLabel").pack(anchor="w", pady=(0, 5))

        ot_data_config = self._cargar_ot_configuration(self.ot_entry.get().strip().upper())
        if ot_data_config:
            ttk.Label(frame, 
                      text=f"🔹 Lado A: {ot_data_config.get('num_conectores_a', 'N/A')} Conectores esperados", 
                      font=("Arial", 10), 
                      style="GeoDetalles.TLabel").pack(anchor="w", pady=(0, 5))
            ttk.Label(frame, 
                      text=f"🔹 Lado B: {ot_data_config.get('num_conectores_b', 'N/A')} Conectores esperados", 
                      font=("Arial", 10), 
                      style="GeoDetalles.TLabel").pack(anchor="w", pady=(0, 5))
        
        tree = ttk.Treeview(frame, columns=("Punta", "Conector", "Resultado", "Fecha"), show="headings", height=5, style="GeoDetalles.TReevie_w")
        tree.heading("Punta", text="Punta", anchor=tk.W)
        tree.heading("Conector", text="Conector", anchor=tk.W)
        tree.heading("Resultado", text="Resultado", anchor=tk.W)
        tree.heading("Fecha", text="Fecha", anchor=tk.W)

        tree.column("Punta", width=70, stretch=tk.NO)
        tree.column("Conector", width=80, stretch=tk.NO)
        tree.column("Resultado", width=100, stretch=tk.NO)
        tree.column("Fecha", width=150, stretch=tk.NO)

        tree.tag_configure('PASS', foreground='green')
        tree.tag_configure('FAIL', foreground='red')
        tree.tag_configure('NO ENCONTRADO', foreground='orange')

        for lado_key in ['A', 'B']:
            punta_data = details_to_show.get('detalles_puntas', {}).get(lado_key, {})
            mediciones = punta_data.get('mediciones', [])
            
            mediciones_ordenadas = sorted(mediciones, key=lambda x: (x.get('conector', 'Z'), x.get('fecha', '')))

            for m in mediciones_ordenadas:
                conector_display = m.get('conector', 'N/A')
                resultado_display = m.get('resultado', 'N/A')
                fecha_display = m.get('fecha', 'N/A')
                
                tags = []
                if 'PASS' in resultado_display:
                    tags = ['PASS']
                elif 'FAIL' in resultado_display:
                    tags = ['FAIL']
                
                tree.insert("", tk.END, values=(
                    lado_key,
                    conector_display,
                    resultado_display,
                    fecha_display
                ), tags=tags)


        tree.pack(fill=tk.BOTH, expand=True, pady=(10, 0))

        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=tree.yview, style="Vertical.TScrollbar")
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        tree.configure(yscrollcommand=scrollbar.set)

        btn_frame = ttk.Frame(frame, style="GeoDetalles.TFrame")
        btn_frame.pack(fill=tk.X, pady=(15, 0))

        ttk.Button(btn_frame, 
                 text="Cerrar", 
                 command=detalles_window.destroy).pack()

        detalles_window.mainloop()

    def seleccionar_nueva_db(self):
        """Permite al usuario seleccionar un nuevo archivo de base de datos"""
        password_ingresada = simpledialog.askstring("Contraseña Requerida", 
                                                "Ingrese la contraseña para cambiar la base de datos:", 
                                                show='*')
        if password_ingresada != self.password:
            messagebox.showerror("Acceso Denegado", "Contraseña incorrecta.")
            return
    
        nuevo_db_path = filedialog.askopenfilename(
            title="Seleccionar archivo de base de datos",
            filetypes=[("Archivos SQLite", "*.db"), ("Todos los archivos", "*.*")],
            initialdir=os.path.dirname(self.db_name) if self.db_name else None
        )
    
        if not nuevo_db_path:
            return
    
        try:
            conn = sqlite3.connect(nuevo_db_path)
            cursor = conn.cursor()
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='cable_verifications'")
            if not cursor.fetchone():
                raise ValueError("La base de datos no contiene la tabla requerida")
            
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='ot_configurations'")
            if not cursor.fetchone():
                raise ValueError("La base de datos no contiene la tabla de configuraciones de OT requerida")

            conn.close()
        except Exception as e:
            messagebox.showerror("Error", 
                            f"El archivo seleccionado no es una base de datos válida: {e}")
            return
    
        self.db_name = nuevo_db_path
        self.guardar_rutas()
        messagebox.showinfo("Éxito", f"Base de datos cambiada a:\n{nuevo_db_path}")

    def solicitar_contrasena(self):
        """Solicita la contraseña para acceder a la configuración de rutas."""
        password_ingresada = simpledialog.askstring("Contraseña Requerida", "Ingrese la contraseña para cambiar las rutas:", show='*')
        if password_ingresada == self.password:
            self.mostrar_ventana_configuracion_rutas()
        else:
            messagebox.showerror("Acceso Denegado", "Contraseña incorrecta.")

    def solicitar_contrasena_registros(self):
        """Solicita la contraseña para acceder a la vista de registros."""
        password_ingresada = simpledialog.askstring("Contraseña Requerida", "Ingrese la contraseña para ver los registros:", show='*')
        if password_ingresada == self.password:
            self.mostrar_vista_registros()
        else:
            messagebox.showerror("Acceso Denegado", "Contraseña incorrecta.")

    def mostrar_ventana_configuracion_rutas(self):
        """Muestra la ventana para configurar las rutas de ILRL y Geometría."""
        config_window = tk.Toplevel(self.root)
        config_window.title("Configurar Rutas de Archivos")
        config_window.geometry("700x300")
        config_window.transient(self.root)
        config_window.grab_set()

        main_frame = ttk.Frame(config_window)
        main_frame.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(main_frame, highlightthickness=0)
        yscrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        xscrollbar = ttk.Scrollbar(main_frame, orient="horizontal", command=canvas.xview)
        scrollable_content_frame = ttk.Frame(canvas)
        
        scrollable_window_id = canvas.create_window((0, 0), window=scrollable_content_frame, anchor="nw")
        canvas.configure(yscrollcommand=yscrollbar.set, xscrollcommand=xscrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        yscrollbar.pack(side="right", fill="y")
        xscrollbar.pack(side="bottom", fill="x")

        def _on_rutas_canvas_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas_width = event.width
            canvas.itemconfigure(scrollable_window_id, width=canvas_width)

        canvas.bind("<Configure>", _on_rutas_canvas_configure)
        canvas.bind_all("<MouseWheel>", self._on_mouse_wheel)

        frame = ttk.Frame(scrollable_content_frame, padding=(20, 20), style="ConfigRutas.TFrame")
        frame.pack(fill=tk.BOTH, expand=True)

        BG_COLOR = "#F7F7F7"
        ACCENT_BLUE = "#007BFF"
        TEXT_COLOR = "#333333"

        style = ttk.Style()
        style.configure("ConfigRutas.TFrame", background=BG_COLOR)
        style.configure("ConfigRutas.TLabel", background=BG_COLOR, foreground=TEXT_COLOR)
        style.configure("ConfigRutas.TEntry", fieldbackground="#FFFFFF", foreground=TEXT_COLOR)
        style.configure("ConfigRutas.TButton", background=ACCENT_BLUE, foreground="white", relief="flat")
        style.map("ConfigRutas.TButton", background=[('active', "#0056b3")])

        ttk.Label(frame, text="Ruta Base ILRL:", font=("Arial", 10, "bold"), style="ConfigRutas.TLabel").grid(row=0, column=0, sticky=tk.W, pady=5)
        ilrl_entry = ttk.Entry(frame, width=60, style="ConfigRutas.TEntry")
        ilrl_entry.insert(0, self.ruta_base_ilrl)
        ilrl_entry.grid(row=0, column=1, pady=5, padx=10, sticky="ew")

        ttk.Label(frame, text="Ruta Base Geometría:", font=("Arial", 10, "bold"), style="ConfigRutas.TLabel").grid(row=1, column=0, sticky=tk.W, pady=5)
        geo_entry = ttk.Entry(frame, width=60, style="ConfigRutas.TEntry")
        geo_entry.insert(0, self.ruta_base_geo)
        geo_entry.grid(row=1, column=1, pady=5, padx=10, sticky="ew")
        
        ttk.Label(frame, text="Ruta Base Polaridad:", font=("Arial", 10, "bold"), style="ConfigRutas.TLabel").grid(row=2, column=0, sticky=tk.W, pady=5)
        polaridad_entry = ttk.Entry(frame, width=60, style="ConfigRutas.TEntry")
        polaridad_entry.insert(0, self.ruta_base_polaridad)
        polaridad_entry.grid(row=2, column=1, pady=5, padx=10, sticky="ew")

        def guardar_nuevas_rutas():
            nueva_ilrl = ilrl_entry.get().strip()
            nueva_geo = geo_entry.get().strip()
            nueva_polaridad = polaridad_entry.get().strip()

            if not os.path.isdir(nueva_ilrl):
                messagebox.showwarning("Ruta Inválida", "La ruta de ILRL no es un directorio válido.")
                return
            if not os.path.isdir(nueva_geo):
                messagebox.showwarning("Ruta Inválida", "La ruta de Geometría no es un directorio válido.")
                return
            if not os.path.isdir(nueva_polaridad):
                messagebox.showwarning("Ruta Inválida", "La ruta de Polaridad no es un directorio válido.")
                return

            self.ruta_base_ilrl = nueva_ilrl
            self.ruta_base_geo = nueva_geo
            self.ruta_base_polaridad = nueva_polaridad
            self.guardar_rutas()
            self.ruta_ilrl_label.config(text=f"📂 Ruta ILRL: {self.ruta_base_ilrl}")
            self.ruta_geo_label.config(text=f"📂 Ruta Geometría: {self.ruta_base_geo}")
            self.ruta_polaridad_label.config(text=f"📂 Ruta Polaridad: {self.ruta_base_polaridad}")
            config_window.destroy()

        save_button = ttk.Button(frame, text="Guardar Rutas", command=guardar_nuevas_rutas, style="ConfigRutas.TButton")
        save_button.grid(row=3, column=0, columnspan=2, pady=20)

        config_window.mainloop()
    
    def _borrar_todos_los_registros(self):
        """Borra todos los registros de la tabla cable_verifications."""
        if not messagebox.askyesno("Confirmar Eliminación", 
                                   "¿Está seguro de que desea borrar TODOS los registros de la base de datos?\n\n"
                                   "Esta acción es irreversible."):
            return

        conn = None
        try:
            conn = sqlite3.connect(self.db_name)
            cursor = conn.cursor()
            cursor.execute("DELETE FROM cable_verifications")
            cursor.execute("DELETE FROM ot_configurations")
            conn.commit()
            messagebox.showinfo("Éxito", "Todos los registros han sido eliminados correctamente.")
            if hasattr(self, 'tree_registros'):
                self.cargar_registros()
        except sqlite3.Error as e:
            messagebox.showerror("Error de Base de Datos", f"No se pudieron borrar los registros: {e}")
        finally:
            if conn:
                conn.close()

    def solicitar_contrasena_borrar_datos(self):
        """Solicita la contraseña para borrar todos los datos de la base de datos."""
        password_ingresada = simpledialog.askstring("Contraseña Requerida", 
                                                     "Ingrese la contraseña para borrar TODOS los datos:", 
                                                     show='*')
        if password_ingresada == self.password:
            self._borrar_todos_los_registros()
        else:
            messagebox.showerror("Acceso Denegado", "Contraseña incorrecta.")

    def mostrar_vista_registros(self):
        """Muestra la ventana para que un ingeniero visualice los registros de cables."""
        registros_window = tk.Toplevel(self.root)
        registros_window.title("Vista de Registros de Cables MPO")
        registros_window.geometry("1000x700")
        registros_window.transient(self.root)
        registros_window.grab_set()

        main_frame = ttk.Frame(registros_window)
        main_frame.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(main_frame, highlightthickness=0)
        yscrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        xscrollbar = ttk.Scrollbar(main_frame, orient="horizontal", command=canvas.xview)
        scrollable_content_frame = ttk.Frame(canvas)
        
        scrollable_window_id = canvas.create_window((0, 0), window=scrollable_content_frame, anchor="nw")
        canvas.configure(yscrollcommand=yscrollbar.set, xscrollcommand=xscrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        yscrollbar.pack(side="right", fill="y")
        xscrollbar.pack(side="bottom", fill="x")

        def _on_registros_canvas_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas_width = event.width
            canvas.itemconfigure(scrollable_window_id, width=canvas_width)

        canvas.bind("<Configure>", _on_registros_canvas_configure)
        canvas.bind_all("<MouseWheel>", self._on_mouse_wheel)
        
        main_frame = ttk.Frame(scrollable_content_frame, padding=(20, 20), style="Registros.TFrame")
        main_frame.pack(fill=tk.BOTH, expand=True)

        BG_COLOR = "#F7F7F7"
        ACCENT_BLUE = "#007BFF"
        TEXT_COLOR = "#333333"

        style = ttk.Style()
        style.configure("Registros.TFrame", background=BG_COLOR)
        style.configure("Registros.TLabel", background=BG_COLOR, foreground=TEXT_COLOR)
        style.configure("Registros.TEntry", fieldbackground="#FFFFFF", foreground=TEXT_COLOR)
        style.configure("Registros.TButton", background=ACCENT_BLUE, foreground="white", relief="flat")
        style.map("Registros.TButton", background=[('active', "#0056b3")])
        style.configure("Registros.Treeview", background="#FFFFFF", fieldbackground="#FFFFFF", foreground=TEXT_COLOR)
        style.map('Registros.Treeview', background=[('selected', ACCENT_BLUE)])
        style.configure("Registros.Treeview.Heading", background=ACCENT_BLUE, foreground="white", font=('Arial', 10, 'bold'))


        filter_frame = ttk.Frame(main_frame, style="Registros.TFrame")
        filter_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(filter_frame, text="Filtrar por OT o Serie:", font=("Arial", 10, "bold"), style="Registros.TLabel").pack(side=tk.LEFT, padx=(0, 5))
        self.filtro_entry = ttk.Entry(filter_frame, width=30, style="Registros.TEntry")
        self.filtro_entry.pack(side=tk.LEFT, padx=(0, 10))
        self.filtro_entry.bind("<KeyRelease>", self.aplicar_filtro_registros)

        btn_aplicar_filtro = ttk.Button(filter_frame, text="Aplicar Filtro", command=self.aplicar_filtro_registros, style="Registros.TButton")
        btn_aplicar_filtro.pack(side=tk.LEFT, padx=(0, 10))

        btn_limpiar_filtro = ttk.Button(filter_frame, text="Limpiar Filtro", command=self.limpiar_filtro_registros, style="Registros.TButton")
        btn_limpiar_filtro.pack(side=tk.LEFT, padx=(0, 20))

        btn_borrar_todos = ttk.Button(filter_frame, text="🗑️ Borrar Todos los Registros", 
                                      command=self.solicitar_contrasena_borrar_datos, style="Registros.TButton")
        btn_borrar_todos.pack(side=tk.RIGHT)

        columns = ("ID", "Fecha Entrada", "Número Serie", "Número OT", "Estado General", 
                   "ILRL Estatus", "ILRL Fecha", "Geo Estatus", "Geo Fecha", "Pol. Estatus", "Pol. Fecha")
        self.tree_registros = ttk.Treeview(main_frame, columns=columns, show="headings", style="Registros.TReevie_w")
        
        for col in columns:
            self.tree_registros.heading(col, text=col, anchor=tk.W)
            self.tree_registros.column(col, width=100, anchor=tk.W)

        self.tree_registros.column("ID", width=50, stretch=tk.NO)
        self.tree_registros.column("Fecha Entrada", width=140, stretch=tk.NO)
        self.tree_registros.column("Número Serie", width=120, stretch=tk.NO)
        self.tree_registros.column("Número OT", width=120, stretch=tk.NO)
        self.tree_registros.column("Estado General", width=100, stretch=tk.NO)
        self.tree_registros.column("ILRL Estatus", width=90, stretch=tk.NO)
        self.tree_registros.column("ILRL Fecha", width=120, stretch=tk.NO)
        self.tree_registros.column("Geo Estatus", width=90, stretch=tk.NO)
        self.tree_registros.column("Geo Fecha", width=120, stretch=tk.NO)
        self.tree_registros.column("Pol. Estatus", width=90, stretch=tk.NO)
        self.tree_registros.column("Pol. Fecha", width=120, stretch=tk.NO)

        self.tree_registros.pack(fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=self.tree_registros.yview, style="Vertical.TScrollbar")
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree_registros.configure(yscrollcommand=scrollbar.set)

        self.tree_registros.tag_configure('APROBADO', foreground='green')
        self.tree_registros.tag_configure('RECHAZADO', foreground='red')
        self.tree_registros.tag_configure('RECHAZADO (Falta Geometría)', foreground='red') 
        self.tree_registros.tag_configure('RECHAZADO (Falta ILRL)', foreground='red') 
        self.tree_registros.tag_configure('RECHAZADO (Falta Polaridad)', foreground='red') 
        self.tree_registros.tag_configure('NO ENCONTRADO', foreground='orange')


        self.tree_registros.bind("<Double-1>", self.mostrar_detalles_registro_bd)

        self.cargar_registros()
        registros_window.mainloop()
        
    def cargar_registros(self):
        """Carga los registros de la base de datos en el Treeview."""
        for item in self.tree_registros.get_children():
            self.tree_registros.delete(item)
        
        self.item_data_cache = {}

        conn = None
        try:
            conn = sqlite3.connect(self.db_name)
            cursor = conn.cursor()
            
            cursor.execute("SELECT * FROM cable_verifications ORDER BY entry_date DESC")
            registros = cursor.fetchall()
            
            for i, row in enumerate(registros):
                ilrl_details = json.loads(row[11]) if len(row) > 11 and row[11] else None
                geo_details = json.loads(row[12]) if len(row) > 12 and row[12] else None
                polaridad_details = json.loads(row[13]) if len(row) > 13 and row[13] else None
                
                self.item_data_cache[row[0]] = {
                    "id": row[0],
                    "entry_date": row[1],
                    "serial_number": row[2],
                    "ot_number": row[3],
                    "overall_status": row[4],
                    "ilrl_status": row[5],
                    "ilrl_date": row[6],
                    "geo_status": row[7],
                    "geo_date": row[8],
                    "polaridad_status": row[9] if len(row) > 9 else 'N/A',
                    "polaridad_date": row[10] if len(row) > 10 else 'N/A',
                    "ilrl_details": ilrl_details,
                    "geo_details": geo_details,
                    "polaridad_details": polaridad_details
                }

                self.tree_registros.insert("", tk.END, iid=row[0], values=(
                    row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8],
                    self.item_data_cache[row[0]]['polaridad_status'],
                    self.item_data_cache[row[0]]['polaridad_date']
                ), tags=(row[4],)) 
        except sqlite3.Error as e:
            messagebox.showerror("Error de Base de Datos", f"No se pudieron cargar los registros: {e}")
        finally:
            if conn:
                conn.close()

    def aplicar_filtro_registros(self, event=None):
        """Aplica un filtro a los registros mostrados en el Treeview."""
        filtro = self.filtro_entry.get().strip().upper()
        
        for item in self.tree_registros.get_children():
            self.tree_registros.delete(item)
            
        conn = None
        try:
            conn = sqlite3.connect(self.db_name)
            cursor = conn.cursor()
            
            if filtro:
                cursor.execute("""
                    SELECT * FROM cable_verifications 
                    WHERE UPPER(ot_number) LIKE ? OR serial_number LIKE ?
                    ORDER BY entry_date DESC
                """, (f"%{filtro}%", f"%{filtro}%"))
            else:
                cursor.execute("SELECT * FROM cable_verifications ORDER BY entry_date DESC")
                
            registros = cursor.fetchall()
            
            for i, row in enumerate(registros):
                ilrl_details = json.loads(row[11]) if len(row) > 11 and row[11] else None
                geo_details = json.loads(row[12]) if len(row) > 12 and row[12] else None
                polaridad_details = json.loads(row[13]) if len(row) > 13 and row[13] else None
                
                self.item_data_cache[row[0]] = {
                    "id": row[0],
                    "entry_date": row[1],
                    "serial_number": row[2],
                    "ot_number": row[3],
                    "overall_status": row[4],
                    "ilrl_status": row[5],
                    "ilrl_date": row[6],
                    "geo_status": row[7],
                    "geo_date": row[8],
                    "polaridad_status": row[9] if len(row) > 9 else 'N/A',
                    "polaridad_date": row[10] if len(row) > 10 else 'N/A',
                    "ilrl_details": ilrl_details,
                    "geo_details": geo_details,
                    "polaridad_details": polaridad_details
                }
                
                self.tree_registros.insert("", tk.END, iid=row[0], values=(
                    row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8],
                    self.item_data_cache[row[0]]['polaridad_status'],
                    self.item_data_cache[row[0]]['polaridad_date']
                ), tags=(row[4],))
        except sqlite3.Error as e:
            messagebox.showerror("Error de Base de Datos", f"No se pudieron cargar los registros filtrados: {e}")
        finally:
            if conn:
                conn.close()

    def limpiar_filtro_registros(self):
        """Limpia el campo de filtro y recarga todos los registros."""
        self.filtro_entry.delete(0, tk.END)
        self.cargar_registros()

    def mostrar_detalles_registro_bd(self, event):
        """Muestra una ventana de detalles para el registro seleccionado en la base de datos."""
        selected_item_id = self.tree_registros.focus()
        if not selected_item_id:
            return

        record_id = int(selected_item_id)
        record_data = self.item_data_cache.get(record_id)

        if not record_data:
            messagebox.showerror("Error", "No se encontraron los detalles del registro.")
            return

        detalles_window = tk.Toplevel(self.root)
        detalles_window.title(f"Detalles del Registro #{record_data['id']}")
        detalles_window.geometry("800x600")
        detalles_window.transient(self.root)
        detalles_window.grab_set()

        main_frame = ttk.Frame(detalles_window)
        main_frame.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(main_frame, highlightthickness=0)
        yscrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        xscrollbar = ttk.Scrollbar(main_frame, orient="horizontal", command=canvas.xview)
        scrollable_content_frame = ttk.Frame(canvas)
        
        scrollable_window_id = canvas.create_window((0, 0), window=scrollable_content_frame, anchor="nw")
        canvas.configure(yscrollcommand=yscrollbar.set, xscrollcommand=xscrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        yscrollbar.pack(side="right", fill="y")
        xscrollbar.pack(side="bottom", fill="x")

        def _on_registro_details_canvas_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas_width = event.width
            canvas.itemconfigure(scrollable_window_id, width=canvas_width)

        canvas.bind("<Configure>", _on_registro_details_canvas_configure)
        canvas.bind_all("<MouseWheel>", self._on_mouse_wheel)
        
        frame = ttk.Frame(scrollable_content_frame, padding=(20, 20), style="DetallesRegistro.TFrame")
        frame.pack(fill=tk.BOTH, expand=True)

        BG_COLOR = "#F7F7F7"
        ACCENT_BLUE = "#007BFF"
        TEXT_COLOR = "#333333"

        style = ttk.Style()
        style.configure("DetallesRegistro.TFrame", background=BG_COLOR)
        style.configure("DetallesRegistro.TLabel", background=BG_COLOR, foreground=TEXT_COLOR)
        style.configure("DetallesRegistro.TButton", background=ACCENT_BLUE, foreground="white", relief="flat")
        style.map("DetallesRegistro.TButton", background=[('active', "#0056b3")])

        ttk.Label(frame, text="📋 Información General:", font=("Arial", 12, "bold"), style="DetallesRegistro.TLabel").pack(anchor="w", pady=(0, 10))
        
        info_general_text = (
            f"   • ID de Registro: {record_data['id']}\n"
            f"   • Fecha de Entrada: {record_data['entry_date']}\n"
            f"   • Número de Serie: {record_data['serial_number']}\n"
            f"   • Número de OT: {record_data['ot_number']}\n"
        )
        ttk.Label(frame, text=info_general_text, justify=tk.LEFT, font=("Arial", 10), style="DetallesRegistro.TLabel").pack(anchor="w")

        ttk.Label(frame, text="🏁 Estado General:", font=("Arial", 12, "bold"), style="DetallesRegistro.TLabel").pack(anchor="w", pady=(10, 5))
        overall_status_color = "green" if record_data['overall_status'] == "APROBADO" else "red" if "RECHAZADO" in record_data['overall_status'] else "orange"
        ttk.Label(frame, text=f"   • {record_data['overall_status']}", font=("Arial", 10, "bold"), foreground=overall_status_color, style="DetallesRegistro.TLabel").pack(anchor="w")

        ttk.Label(frame, text="📊 Detalles ILRL:", font=("Arial", 12, "bold"), style="DetallesRegistro.TLabel").pack(anchor="w", pady=(10, 5))
        ilrl_status_color = "green" if record_data['ilrl_status'] == "APROBADO" else "red" if "RECHAZADO" in record_data['ilrl_status'] else "orange"
        ttk.Label(frame, text=f"   • Estado: {record_data['ilrl_status']}", font=("Arial", 10, "bold"), foreground=ilrl_status_color, style="DetallesRegistro.TLabel").pack(anchor="w")
        ttk.Label(frame, text=f"   • Fecha: {record_data['ilrl_date'] if record_data['ilrl_date'] else 'N/A'}", font=("Arial", 10), style="DetallesRegistro.TLabel").pack(anchor="w")
        
        ilrl_details_from_db = record_data['ilrl_details']
        if ilrl_details_from_db:
            btn_ver_detalles_ilrl = ttk.Button(frame, text="Ver Detalles ILRL (Ventana Completa)", 
                                            command=lambda: self.mostrar_detalles_ilrl(ilrl_details_from_db), style="DetallesRegistro.TButton")
            btn_ver_detalles_ilrl.pack(anchor="w", pady=(5, 5))
        else:
            ttk.Label(frame, text="   • No hay detalles ILRL disponibles.", font=("Arial", 10), foreground="#999999", style="DetallesRegistro.TLabel").pack(anchor="w")

        ttk.Label(frame, text="📐 Detalles Geometría:", font=("Arial", 12, "bold"), style="DetallesRegistro.TLabel").pack(anchor="w", pady=(10, 5))
        geo_status_color = "green" if record_data['geo_status'] == "APROBADO" else "red" if "RECHAZADO" in record_data['geo_status'] else "orange"
        ttk.Label(frame, text=f"   • Estado: {record_data['geo_status']}", font=("Arial", 10, "bold"), foreground=geo_status_color, style="DetallesRegistro.TLabel").pack(anchor="w")
        geo_date_str = record_data['geo_date'] if record_data['geo_date'] else 'N/A'
        ttk.Label(frame, text=f"   • Fecha: {geo_date_str}", font=("Arial", 10), style="DetallesRegistro.TLabel").pack(anchor="w")

        if record_data['geo_details']:
            btn_ver_detalles_geo = ttk.Button(frame, text="Ver Detalles Geometría (Ventana Completa)", 
                                              command=lambda: self.mostrar_detalles_geo(record_data['geo_details']), style="DetallesRegistro.TButton")
            btn_ver_detalles_geo.pack(anchor="w", pady=(5, 5))
        else:
            ttk.Label(frame, text="   • No hay detalles de Geometría disponibles.", font=("Arial", 10), foreground="#999999", style="DetallesRegistro.TLabel").pack(anchor="w")

        ttk.Label(frame, text="🔀 Detalles Polaridad:", font=("Arial", 12, "bold"), style="DetallesRegistro.TLabel").pack(anchor="w", pady=(10, 5))
        pol_status_color = "green" if record_data['polaridad_status'] == "PASS" else "red" if "RECHAZADO" in record_data['polaridad_status'] else "orange"
        ttk.Label(frame, text=f"   • Estado: {record_data['polaridad_status']}", font=("Arial", 10, "bold"), foreground=pol_status_color, style="DetallesRegistro.TLabel").pack(anchor="w")
        pol_date_str = record_data['polaridad_date'] if record_data['polaridad_date'] else 'N/A'
        ttk.Label(frame, text=f"   • Fecha: {pol_date_str}", font=("Arial", 10), style="DetallesRegistro.TLabel").pack(anchor="w")
        
        if record_data['polaridad_details']:
            btn_ver_detalles_pol = ttk.Button(frame, text="Ver Detalles Polaridad (Ventana Completa)",
                                              command=lambda: self.mostrar_detalles_polaridad(record_data['polaridad_details']), style="DetallesRegistro.TButton")
            btn_ver_detalles_pol.pack(anchor="w", pady=(5, 5))
        else:
            ttk.Label(frame, text="   • No hay detalles de Polaridad disponibles.", font=("Arial", 10), foreground="#999999", style="DetallesRegistro.TLabel").pack(anchor="w")

        detalles_window.mainloop()


if __name__ == "__main__":
    app = VerificadorCablesMPO()
    app.create_main_window()
