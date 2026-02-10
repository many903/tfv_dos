#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
OCR TO EXCEL APP v3.0 - Aplicación completa 100% offline
Autor: OCR Tool Team
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from PIL import Image, ImageTk, ImageEnhance, ImageFilter
import pytesseract
import cv2
import numpy as np
import pandas as pd
from datetime import datetime
import threading
import json
import traceback
import webbrowser
from pathlib import Path

# ============================================================================
# CONFIGURACIÓN Y UTILIDADES
# ============================================================================

class ConfigManager:
    """Gestor de configuración de la aplicación"""
    
    DEFAULT_CONFIG = {
        "app": {
            "version": "3.0",
            "language": "es",
            "theme": "light",
            "auto_save": True,
            "auto_export": False
        },
        "paths": {
            "tesseract": "",
            "last_folder": "",
            "export_folder": "exportados",
            "tessdata": "tessdata"
        },
        "ocr": {
            "language": "eng",
            "psm": "6",
            "oem": "3",
            "dpi": "300"
        },
        "preprocessing": {
            "grayscale": True,
            "denoise": True,
            "contrast": 1.5,
            "brightness": 1.0,
            "threshold": "adaptive",
            "deskew": True
        },
        "ui": {
            "font_size": 10,
            "font_family": "Segoe UI",
            "show_grid": True,
            "alternate_colors": True
        }
    }
    
    def __init__(self):
        self.config_file = "config/settings.json"
        self.load_config()
        
    def load_config(self):
        """Cargar configuración desde archivo"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    self.config = json.load(f)
                # Actualizar con valores por defecto si faltan
                for section, values in self.DEFAULT_CONFIG.items():
                    if section not in self.config:
                        self.config[section] = values
                    else:
                        for key, value in values.items():
                            if key not in self.config[section]:
                                self.config[section][key] = value
            else:
                self.config = self.DEFAULT_CONFIG.copy()
                self.save_config()
        except Exception as e:
            print(f"Error cargando configuración: {e}")
            self.config = self.DEFAULT_CONFIG.copy()
            
    def save_config(self):
        """Guardar configuración en archivo"""
        try:
            os.makedirs(os.path.dirname(self.config_file), exist_ok=True)
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4, ensure_ascii=False)
        except Exception as e:
            print(f"Error guardando configuración: {e}")
            
    def get(self, key, default=None):
        """Obtener valor de configuración"""
        try:
            keys = key.split('.')
            value = self.config
            for k in keys:
                value = value[k]
            return value
        except:
            return default
            
    def set(self, key, value):
        """Establecer valor de configuración"""
        try:
            keys = key.split('.')
            config = self.config
            for k in keys[:-1]:
                if k not in config:
                    config[k] = {}
                config = config[k]
            config[keys[-1]] = value
            self.save_config()
        except Exception as e:
            print(f"Error guardando configuración {key}: {e}")

class ImageProcessor:
    """Procesador de imágenes para OCR"""
    
    @staticmethod
    def preprocess_image(image, config):
        """Preprocesar imagen para mejorar OCR"""
        try:
            # Convertir PIL a OpenCV
            if image.mode != 'RGB':
                img_cv = cv2.cvtColor(np.array(image), cv2.COLOR_GRAY2BGR)
            else:
                img_cv = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
            
            # Convertir a escala de grises
            gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)
            
            # Aplicar configuración de preprocesamiento
            if config.get('preprocessing.grayscale', True):
                pass  # Ya está en escala de grises
            
            if config.get('preprocessing.denoise', True):
                gray = cv2.medianBlur(gray, 3)
            
            # Ajustar brillo y contraste
            alpha = config.get('preprocessing.contrast', 1.5)
            beta = config.get('preprocessing.brightness', 1.0) * 50 - 50
            gray = cv2.convertScaleAbs(gray, alpha=alpha, beta=beta)
            
            # Umbralización
            threshold_type = config.get('preprocessing.threshold', 'adaptive')
            if threshold_type == 'adaptive':
                thresh = cv2.adaptiveThreshold(
                    gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                    cv2.THRESH_BINARY, 11, 2
                )
            elif threshold_type == 'otsu':
                _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            else:
                _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY)
            
            # Enderezar imagen si está configurado
            if config.get('preprocessing.deskew', True):
                thresh = ImageProcessor.deskew_image(thresh)
            
            return thresh
            
        except Exception as e:
            print(f"Error en preprocesamiento: {e}")
            return np.array(image.convert('L'))
    
    @staticmethod
    def deskew_image(image):
        """Enderezar imagen inclinada"""
        try:
            # Encontrar contornos de texto
            coords = np.column_stack(np.where(image > 0))
            
            if len(coords) < 10:
                return image
            
            # Calcular ángulo de inclinación
            angle = cv2.minAreaRect(coords)[-1]
            
            if angle < -45:
                angle = 90 + angle
            elif angle > 45:
                angle = angle - 90
            
            # Rotar imagen para corregir
            if abs(angle) > 1.0:
                (h, w) = image.shape[:2]
                center = (w // 2, h // 2)
                M = cv2.getRotationMatrix2D(center, angle, 1.0)
                image = cv2.warpAffine(image, M, (w, h),
                                     flags=cv2.INTER_CUBIC,
                                     borderMode=cv2.BORDER_REPLICATE)
            
            return image
            
        except:
            return image
    
    @staticmethod
    def resize_for_display(image, max_width=800, max_height=600):
        """Redimensionar imagen para visualización"""
        width, height = image.size
        
        if width > max_width or height > max_height:
            ratio = min(max_width/width, max_height/height)
            new_width = int(width * ratio)
            new_height = int(height * ratio)
            return image.resize((new_width, new_height), Image.Resampling.LANCZOS)
        
        return image

# ============================================================================
# INTERFAZ GRÁFICA PRINCIPAL
# ============================================================================

class OCRApp:
    """Aplicación principal OCR to Excel"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("OCR to Excel v3.0 - 100% Offline")
        self.root.geometry("1300x800")
        
        # Configuración
        self.config = ConfigManager()
        
        # Variables de estado
        self.image_path = None
        self.original_image = None
        self.preview_image = None
        self.ocr_text = ""
        self.ocr_data = []
        self.headers = []
        self.processing = False
        
        # Configurar Tesseract
        self.setup_tesseract()
        
        # Configurar interfaz
        self.setup_styles()
        self.create_widgets()
        
        # Centrar ventana
        self.center_window()
        
        # Cargar última imagen si existe
        self.load_last_image()
        
    def setup_tesseract(self):
        """Configurar Tesseract OCR"""
        try:
            # Intentar encontrar Tesseract automáticamente
            tesseract_path = self.config.get("paths.tesseract", "")
            
            if tesseract_path and os.path.exists(tesseract_path):
                pytesseract.pytesseract.tesseract_cmd = tesseract_path
                return True
            
            # Rutas comunes
            common_paths = [
                r"C:\Program Files\Tesseract-OCR\tesseract.exe",
                r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
                "/usr/bin/tesseract",
                "/usr/local/bin/tesseract"
            ]
            
            for path in common_paths:
                if os.path.exists(path):
                    pytesseract.pytesseract.tesseract_cmd = path
                    self.config.set("paths.tesseract", path)
                    return True
            
            # Si no se encuentra, intentar usar el del PATH
            import shutil
            tesseract_cmd = shutil.which("tesseract")
            if tesseract_cmd:
                pytesseract.pytesseract.tesseract_cmd = tesseract_cmd
                self.config.set("paths.tesseract", tesseract_cmd)
                return True
            
            messagebox.showwarning(
                "Tesseract no encontrado",
                "Tesseract OCR no está instalado o configurado.\n\n"
                "La aplicación intentará usar OCR básico.\n"
                "Para mejor precisión, instala Tesseract desde:\n"
                "https://github.com/UB-Mannheim/tesseract/wiki\n\n"
                "Luego configura la ruta en: Configuración → Tesseract Path"
            )
            return False
            
        except Exception as e:
            print(f"Error configurando Tesseract: {e}")
            return False
    
    def setup_styles(self):
        """Configurar estilos de la aplicación"""
        style = ttk.Style()
        
        # Configurar tema según configuración
        theme = self.config.get("app.theme", "light")
        
        if theme == "dark":
            bg_color = "#2b2b2b"
            fg_color = "#ffffff"
            entry_bg = "#3c3c3c"
        else:
            bg_color = "#f5f5f5"
            fg_color = "#000000"
            entry_bg = "#ffffff"
        
        self.root.configure(bg=bg_color)
        
        # Configurar estilos
        style.theme_use('clam')
        
        # Colores personalizados
        self.colors = {
            'primary': '#2563eb',
            'secondary': '#10b981',
            'danger': '#ef4444',
            'warning': '#f59e0b',
            'bg': bg_color,
            'fg': fg_color,
            'entry_bg': entry_bg
        }
        
        # Configurar estilos de widgets
        style.configure('Title.TLabel', 
                       font=('Segoe UI', 16, 'bold'),
                       foreground=self.colors['primary'])
        
        style.configure('Subtitle.TLabel',
                       font=('Segoe UI', 10),
                       foreground='#666666')
        
        style.configure('Primary.TButton',
                       font=('Segoe UI', 10, 'bold'),
                       padding=10)
        
        style.configure('Secondary.TButton',
                       font=('Segoe UI', 10),
                       padding=8)
        
        style.configure('Accent.TButton',
                       background=self.colors['primary'],
                       foreground='white',
                       font=('Segoe UI', 10, 'bold'),
                       padding=10)
        
        style.map('Accent.TButton',
                 background=[('active', '#1d4ed8')])
    
    def create_widgets(self):
        """Crear todos los widgets de la interfaz"""
        # Frame principal
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # ========== PANEL SUPERIOR ==========
        top_frame = ttk.Frame(main_frame)
        top_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Título
        title_label = ttk.Label(top_frame, 
                               text="🖼️ OCR to Excel v3.0",
                               style='Title.TLabel')
        title_label.pack(side=tk.LEFT)
        
        # Subtítulo
        subtitle_label = ttk.Label(top_frame,
                                  text="Extrae texto de imágenes y exporta a Excel - 100% Offline",
                                  style='Subtitle.TLabel')
        subtitle_label.pack(side=tk.LEFT, padx=(10, 0))
        
        # ========== PANEL IZQUIERDO (Imagen) ==========
        left_panel = ttk.LabelFrame(main_frame, text="Imagen", padding=10)
        left_panel.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        
        # Botones de imagen
        btn_frame = ttk.Frame(left_panel)
        btn_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Button(btn_frame, text="📂 Abrir Imagen",
                  command=self.load_image,
                  style='Primary.TButton').pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(btn_frame, text="🗑️ Limpiar",
                  command=self.clear_all,
                  style='Secondary.TButton').pack(side=tk.LEFT)
        
        # Ruta del archivo
        self.file_label = ttk.Label(left_panel, text="Ninguna imagen seleccionada")
        self.file_label.pack(anchor=tk.W, pady=(0, 10))
        
        # Canvas para vista previa
        self.preview_canvas = tk.Canvas(left_panel, bg='black',
                                       highlightthickness=1,
                                       highlightbackground='#cccccc')
        self.preview_canvas.pack(fill=tk.BOTH, expand=True)
        
        # Label para cuando no hay imagen
        self.preview_label = ttk.Label(self.preview_canvas,
                                      text="Vista previa de la imagen\n\nArrastra una imagen aquí\no haz clic en 'Abrir Imagen'",
                                      foreground='white',
                                      background='black',
                                      font=('Segoe UI', 12))
        self.preview_label.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
        
        # Configurar drag & drop
        self.setup_drag_drop()
        
        # ========== PANEL DERECHO (Resultados) ==========
        right_panel = ttk.Frame(main_frame)
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # Notebook para pestañas
        self.notebook = ttk.Notebook(right_panel)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Pestaña 1: Texto extraído
        text_tab = ttk.Frame(self.notebook)
        self.notebook.add(text_tab, text="📝 Texto Extraído")
        
        self.text_area = scrolledtext.ScrolledText(text_tab,
                                                  wrap=tk.WORD,
                                                  font=('Consolas', 10),
                                                  undo=True)
        self.text_area.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Pestaña 2: Tabla de datos
        table_tab = ttk.Frame(self.notebook)
        self.notebook.add(table_tab, text="📊 Tabla de Datos")
        
        # Frame para controles de tabla
        table_controls = ttk.Frame(table_tab)
        table_controls.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(table_controls, text="📥 Exportar Excel",
                  command=self.export_to_excel,
                  style='Secondary.TButton').pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(table_controls, text="📋 Copiar Tabla",
                  command=self.copy_table,
                  style='Secondary.TButton').pack(side=tk.LEFT)
        
        # Treeview para tabla
        self.create_table_widget(table_tab)
        
        # ========== PANEL INFERIOR ==========
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill=tk.X, pady=(10, 0))
        
        # Botón de procesamiento
        self.process_btn = ttk.Button(bottom_frame,
                                     text="🔍 Procesar OCR",
                                     command=self.process_ocr,
                                     style='Accent.TButton',
                                     state='disabled')
        self.process_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Barra de progreso
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(bottom_frame,
                                          variable=self.progress_var,
                                          maximum=100,
                                          mode='determinate')
        self.progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        # Label de estado
        self.status_label = ttk.Label(bottom_frame, text="Listo")
        self.status_label.pack(side=tk.LEFT)
        
        # ========== MENÚ ==========
        self.create_menu()
        
        # Configurar eventos
        self.setup_events()
    
    def create_table_widget(self, parent):
        """Crear widget de tabla para mostrar resultados"""
        # Frame para la tabla con scrollbars
        table_frame = ttk.Frame(parent)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Scrollbars
        v_scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL)
        h_scrollbar = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL)
        
        # Treeview
        self.tree = ttk.Treeview(table_frame,
                                yscrollcommand=v_scrollbar.set,
                                xscrollcommand=h_scrollbar.set,
                                selectmode='extended',
                                show='headings')
        
        v_scrollbar.config(command=self.tree.yview)
        h_scrollbar.config(command=self.tree.xview)
        
        # Layout
        self.tree.grid(row=0, column=0, sticky=tk.NSEW)
        v_scrollbar.grid(row=0, column=1, sticky=tk.NS)
        h_scrollbar.grid(row=1, column=0, sticky=tk.EW)
        
        # Configurar grid weights
        table_frame.grid_columnconfigure(0, weight=1)
        table_frame.grid_rowconfigure(0, weight=1)
        
        # Configurar doble clic para editar
        self.tree.bind('<Double-1>', self.edit_cell)
    
    def create_menu(self):
        """Crear menú de la aplicación"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # Menú Archivo
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Archivo", menu=file_menu)
        file_menu.add_command(label="Abrir Imagen...", 
                             command=self.load_image, 
                             accelerator="Ctrl+O")
        file_menu.add_separator()
        file_menu.add_command(label="Exportar a Excel...",
                             command=self.export_to_excel,
                             accelerator="Ctrl+E")
        file_menu.add_command(label="Guardar Texto...",
                             command=self.save_text)
        file_menu.add_separator()
        file_menu.add_command(label="Salir",
                             command=self.root.quit,
                             accelerator="Alt+F4")
        
        # Menú Edición
        edit_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Edición", menu=edit_menu)
        edit_menu.add_command(label="Copiar Texto",
                             command=self.copy_text)
        edit_menu.add_command(label="Pegar Texto",
                             command=self.paste_text)
        edit_menu.add_separator()
        edit_menu.add_command(label="Limpiar Todo",
                             command=self.clear_all)
        
        # Menú Herramientas
        tools_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Herramientas", menu=tools_menu)
        tools_menu.add_command(label="Configurar Tesseract...",
                             command=self.configure_tesseract)
        tools_menu.add_command(label="Preprocesar Imagen",
                             command=self.show_preprocessing_dialog)
        tools_menu.add_separator()
        tools_menu.add_command(label="Reconocer Tablas",
                             command=self.detect_tables)
        
        # Menú Ayuda
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Ayuda", menu=help_menu)
        help_menu.add_command(label="Manual de Usuario",
                             command=self.show_help)
        help_menu.add_command(label="Acerca de...",
                             command=self.show_about)
        
        # Atajos de teclado
        self.root.bind('<Control-o>', lambda e: self.load_image())
        self.root.bind('<Control-e>', lambda e: self.export_to_excel())
        self.root.bind('<Control-s>', lambda e: self.save_text())
        self.root.bind('<F1>', lambda e: self.show_help())
    
    def setup_drag_drop(self):
        """Configurar drag & drop para imágenes"""
        def handle_drop(event):
            files = self.root.tk.splitlist(event.data)
            if files:
                self.load_dropped_image(files[0])
        
        def drag_enter(event):
            self.preview_canvas.config(highlightbackground='#2563eb', highlightthickness=2)
        
        def drag_leave(event):
            self.preview_canvas.config(highlightbackground='#cccccc', highlightthickness=1)
        
        # Configurar eventos de drag & drop
        self.preview_canvas.drop_target_register('DND_Files')
        self.preview_canvas.dnd_bind('<<Drop>>', handle_drop)
        self.preview_canvas.dnd_bind('<<DragEnter>>', drag_enter)
        self.preview_canvas.dnd_bind('<<DragLeave>>', drag_leave)
    
    def setup_events(self):
        """Configurar eventos adicionales"""
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def center_window(self):
        """Centrar ventana en pantalla"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def load_last_image(self):
        """Cargar última imagen usada"""
        last_folder = self.config.get("paths.last_folder", "")
        if last_folder and os.path.exists(last_folder):
            # Solo actualizar el folder inicial
            pass
    
    # ============================================================================
    # FUNCIONALIDADES PRINCIPALES
    # ============================================================================
    
    def load_image(self):
        """Cargar imagen desde archivo"""
        filetypes = [
            ('Imágenes', '*.png *.jpg *.jpeg *.bmp *.tiff'),
            ('Todos los archivos', '*.*')
        ]
        
        last_folder = self.config.get("paths.last_folder", "")
        
        filename = filedialog.askopenfilename(
            title='Seleccionar imagen',
            filetypes=filetypes,
            initialdir=last_folder
        )
        
        if filename:
            self.load_image_file(filename)
    
    def load_dropped_image(self, filename):
        """Cargar imagen arrastrada"""
        self.load_image_file(filename)
    
    def load_image_file(self, filename):
        """Cargar archivo de imagen"""
        try:
            self.image_path = filename
            self.original_image = Image.open(filename)
            
            # Actualizar configuración
            self.config.set("paths.last_folder", os.path.dirname(filename))
            
            # Mostrar vista previa
            self.update_preview()
            
            # Actualizar interfaz
            self.file_label.config(text=os.path.basename(filename))
            self.process_btn.config(state='normal')
            self.status_label.config(text=f"Imagen cargada: {os.path.basename(filename)}")
            
            # Limpiar resultados anteriores
            self.clear_results()
            
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar la imagen:\n{str(e)}")
    
    def update_preview(self):
        """Actualizar vista previa de la imagen"""
        if not self.original_image:
            return
        
        try:
            # Redimensionar para vista previa
            preview_image = ImageProcessor.resize_for_display(self.original_image, 400, 300)
            
            # Convertir para tkinter
            self.tk_preview = ImageTk.PhotoImage(preview_image)
            
            # Actualizar canvas
            self.preview_canvas.delete("all")
            canvas_width = self.preview_canvas.winfo_width()
            canvas_height = self.preview_canvas.winfo_height()
            
            if canvas_width > 1 and canvas_height > 1:
                x = (canvas_width - preview_image.width) // 2
                y = (canvas_height - preview_image.height) // 2
                self.preview_canvas.create_image(x, y, anchor=tk.NW, image=self.tk_preview)
            
            # Ocultar label de placeholder
            self.preview_label.place_forget()
            
        except Exception as e:
            print(f"Error actualizando vista previa: {e}")
    
    def process_ocr(self):
        """Iniciar procesamiento OCR"""
        if self.processing or not self.original_image:
            return
        
        # Actualizar estado
        self.processing = True
        self.process_btn.config(state='disabled')
        self.progress_var.set(0)
        self.status_label.config(text="Procesando OCR...")
        
        # Limpiar resultados anteriores
        self.clear_results()
        
        # Ejecutar en hilo separado
        thread = threading.Thread(target=self._process_ocr_thread, daemon=True)
        thread.start()
    
    def _process_ocr_thread(self):
        """Procesar OCR en hilo separado"""
        try:
            # Paso 1: Preprocesar imagen
            self._update_progress(10, "Preprocesando imagen...")
            processed_image = ImageProcessor.preprocess_image(self.original_image, self.config)
            
            # Paso 2: Configurar OCR
            self._update_progress(30, "Configurando OCR...")
            lang = self.config.get("ocr.language", "eng")
            psm = self.config.get("ocr.psm", "6")
            oem = self.config.get("ocr.oem", "3")
            
            custom_config = f'--psm {psm} --oem {oem}'
            
            # Paso 3: Ejecutar OCR
            self._update_progress(50, "Ejecutando reconocimiento OCR...")
            text = pytesseract.image_to_string(processed_image, lang=lang, config=custom_config)
            
            # Paso 4: Procesar resultados
            self._update_progress(80, "Procesando resultados...")
            
            # Guardar texto completo
            self.ocr_text = text
            
            # Procesar como tabla
            self._process_text_to_table(text)
            
            # Paso 5: Completar
            self._update_progress(100, "Completado")
            self.root.after(0, self._ocr_completed)
            
        except Exception as e:
            error_msg = f"Error en OCR: {str(e)}"
            print(traceback.format_exc())
            self.root.after(0, lambda: self._ocr_failed(error_msg))
    
    def _update_progress(self, value, message):
        """Actualizar progreso desde hilo"""
        self.root.after(0, lambda: self.progress_var.set(value))
        self.root.after(0, lambda: self.status_label.config(text=message))
    
    def _process_text_to_table(self, text):
        """Convertir texto OCR a tabla"""
        # Dividir en líneas
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        
        if not lines:
            return
        
        # Detectar separadores
        first_line = lines[0]
        separators = ['\t', '  ', '|', ',', ';']
        
        detected_separator = None
        for sep in separators:
            if sep in first_line:
                parts = first_line.split(sep)
                if len(parts) > 1:
                    detected_separator = sep
                    break
        
        # Procesar líneas
        if detected_separator:
            # Procesar como tabla con separador
            self.headers = [h.strip() for h in first_line.split(detected_separator) if h.strip()]
            data_lines = lines[1:] if len(self.headers) > 1 else lines
            
            for line in data_lines:
                cells = [c.strip() for c in line.split(detected_separator) if c.strip()]
                if cells:
                    self.ocr_data.append(cells)
        else:
            # Procesar como texto simple
            self.headers = ["Texto Extraído"]
            for line in lines:
                self.ocr_data.append([line])
    
    def _ocr_completed(self):
        """Llamado cuando OCR se completa exitosamente"""
        self.processing = False
        self.process_btn.config(state='normal')
        
        # Mostrar texto en área de texto
        self.text_area.delete('1.0', tk.END)
        self.text_area.insert('1.0', self.ocr_text)
        
        # Mostrar datos en tabla
        self.display_table()
        
        # Cambiar a pestaña de texto
        self.notebook.select(0)
        
        # Actualizar estado
        rows = len(self.ocr_data)
        self.status_label.config(text=f"OCR completado: {rows} filas detectadas")
        
        # Mostrar notificación
        if rows > 0:
            messagebox.showinfo("OCR Completado", 
                              f"Se extrajeron {rows} filas de datos.\n"
                              f"Revisa y edita los datos antes de exportar.")
    
    def _ocr_failed(self, error_msg):
        """Llamado cuando OCR falla"""
        self.processing = False
        self.process_btn.config(state='normal')
        self.progress_var.set(0)
        self.status_label.config(text="Error en OCR")
        
        messagebox.showerror("Error OCR", error_msg)
    
    def display_table(self):
        """Mostrar datos en la tabla Treeview"""
        # Limpiar tabla existente
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Limpiar columnas existentes
        self.tree['columns'] = []
        
        # Configurar columnas si hay datos
        if self.headers and self.ocr_data:
            self.tree['columns'] = self.headers
            
            # Configurar encabezados
            for header in self.headers:
                self.tree.heading(header, text=header)
                self.tree.column(header, width=150, anchor='w', stretch=True)
            
            # Insertar datos
            for row_idx, row in enumerate(self.ocr_data):
                # Asegurar que la fila tenga todas las columnas
                values = row + [''] * (len(self.headers) - len(row))
                self.tree.insert('', 'end', values=values)
    
    def edit_cell(self, event):
        """Editar celda al hacer doble clic"""
        # Identificar ítem y columna
        row_id = self.tree.identify_row(event.y)
        column = self.tree.identify_column(event.x)
        
        if not row_id or not column:
            return
        
        # Obtener índices
        col_idx = int(column[1:]) - 1
        row_idx = self.tree.index(row_id)
        
        # Obtener valor actual
        current_value = self.tree.item(row_id, 'values')[col_idx]
        
        # Crear ventana de edición
        self.create_edit_dialog(row_idx, col_idx, current_value, row_id)
    
    def create_edit_dialog(self, row_idx, col_idx, current_value, row_id):
        """Crear diálogo para editar celda"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Editar Celda")
        dialog.geometry("400x200")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Frame principal
        frame = ttk.Frame(dialog, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Etiqueta
        col_name = self.headers[col_idx] if col_idx < len(self.headers) else f"Columna {col_idx+1}"
        ttk.Label(frame, 
                 text=f"Fila {row_idx+1}, {col_name}",
                 font=('Segoe UI', 12, 'bold')).pack(anchor=tk.W, pady=(0, 10))
        
        # Campo de texto
        text_frame = ttk.Frame(frame)
        text_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
        
        text_widget = tk.Text(text_frame, height=5, font=('Segoe UI', 10))
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        text_widget.insert('1.0', current_value)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(text_frame, command=text_widget.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        text_widget.config(yscrollcommand=scrollbar.set)
        
        # Botones
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X)
        
        def save_changes():
            new_value = text_widget.get('1.0', 'end-1c').strip()
            
            # Actualizar datos
            if row_idx < len(self.ocr_data):
                if col_idx >= len(self.ocr_data[row_idx]):
                    # Extender la fila si es necesario
                    self.ocr_data[row_idx].extend([''] * (col_idx - len(self.ocr_data[row_idx]) + 1))
                self.ocr_data[row_idx][col_idx] = new_value
            
            # Actualizar Treeview
            values = self.ocr_data[row_idx] + [''] * (len(self.headers) - len(self.ocr_data[row_idx]))
            self.tree.item(row_id, values=values)
            
            dialog.destroy()
        
        ttk.Button(btn_frame, text="💾 Guardar",
                  command=save_changes,
                  style='Accent.TButton').pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(btn_frame, text="❌ Cancelar",
                  command=dialog.destroy).pack(side=tk.LEFT)
        
        # Centrar diálogo
        dialog.update_idletasks()
        width = dialog.winfo_width