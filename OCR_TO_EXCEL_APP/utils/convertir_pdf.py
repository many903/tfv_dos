#!/usr/bin/env python3
"""
Convertidor de PDF a imágenes para OCR
"""

import os
import sys
from pdf2image import convert_from_path
from PIL import Image
import tempfile

def pdf_a_imagenes(ruta_pdf, dpi=300, formato='PNG', primera_pagina=None, ultima_pagina=None):
    """
    Convertir PDF a imágenes
    
    Args:
        ruta_pdf: Ruta del archivo PDF
        dpi: Resolución DPI (recomendado 300 para OCR)
        formato: Formato de salida ('PNG', 'JPEG', 'TIFF')
        primera_pagina: Primera página a convertir (1-indexed)
        ultima_pagina: Última página a convertir (1-indexed)
    
    Returns:
        Lista de rutas de imágenes generadas
    """
    
    try:
        # Verificar que el PDF existe
        if not os.path.exists(ruta_pdf):
            raise FileNotFoundError(f"PDF no encontrado: {ruta_pdf}")
        
        print(f"Convirtiendo PDF: {ruta_pdf}")
        print(f"DPI: {dpi}, Formato: {formato}")
        
        # Crear carpeta temporal para imágenes
        temp_dir = tempfile.mkdtemp(prefix='pdf_ocr_')
        print(f"Directorio temporal: {temp_dir}")
        
        # Convertir PDF a imágenes
        imagenes = convert_from_path(
            ruta_pdf,
            dpi=dpi,
            first_page=primera_pagina,
            last_page=ultima_pagina,
            output_folder=temp_dir,
            fmt=formato.lower(),
            thread_count=4
        )
        
        rutas_imagenes = []
        
        # Guardar imágenes
        for i, imagen in enumerate(imagenes):
            numero_pagina = i + (primera_pagina or 1)
            nombre_archivo = f"pagina_{numero_pagina:03d}.{formato.lower()}"
            ruta_completa = os.path.join(temp_dir, nombre_archivo)
            
            imagen.save(ruta_completa, formato)
            rutas_imagenes.append(ruta_completa)
            
            print(f"  Página {numero_pagina}: {ruta_completa}")
        
        print(f"\nConversión completada: {len(rutas_imagenes)} páginas convertidas")
        
        return rutas_imagenes, temp_dir
        
    except Exception as e:
        print(f"Error convirtiendo PDF: {e}")
        return [], None

def pdf_a_imagen_unica(ruta_pdf, dpi=300, unir_vertical=True):
    """
    Convertir PDF completo a una sola imagen
    
    Args:
        ruta_pdf: Ruta del archivo PDF
        dpi: Resolución DPI
        unir_vertical: True para unir verticalmente, False para horizontalmente
    
    Returns:
        PIL.Image: Imagen única con todas las páginas
    """
    
    try:
        # Convertir todas las páginas
        imagenes = convert_from_path(ruta_pdf, dpi=dpi)
        
        if not imagenes:
            raise ValueError("No se pudieron extraer páginas del PDF")
        
        # Determinar dimensiones de la imagen final
        if unir_vertical:
            # Unir verticalmente
            ancho_max = max(img.width for img in imagenes)
            alto_total = sum(img.height for img in imagenes)
            
            # Crear imagen en blanco
            imagen_final = Image.new('RGB', (ancho_max, alto_total), (255, 255, 255))
            
            # Pegar cada imagen
            y_offset = 0
            for i, img in enumerate(imagenes):
                # Centrar horizontalmente si es más pequeña
                x_offset = (ancho_max - img.width) // 2
                imagen_final.paste(img, (x_offset, y_offset))
                y_offset += img.height
        else:
            # Unir horizontalmente
            alto_max = max(img.height for img in imagenes)
            ancho_total = sum(img.width for img in imagenes)
            
            # Crear imagen en blanco
            imagen_final = Image.new('RGB', (ancho_total, alto_max), (255, 255, 255))
            
            # Pegar cada imagen
            x_offset = 0
            for i, img in enumerate(imagenes):
                # Centrar verticalmente si es más pequeña
                y_offset = (alto_max - img.height) // 2
                imagen_final.paste(img, (x_offset, y_offset))
                x_offset += img.width
        
        return imagen_final
        
    except Exception as e:
        print(f"Error convirtiendo PDF a imagen única: {e}")
        return None

def extraer_texto_pdf(ruta_pdf, idioma='eng'):
    """
    Extraer texto de PDF directamente usando OCR en cada página
    
    Args:
        ruta_pdf: Ruta del archivo PDF
        idioma: Idioma para OCR
    
    Returns:
        str: Texto extraído
    """
    try:
        import pytesseract
        
        # Convertir PDF a imágenes
        imagenes = convert_from_path(ruta_pdf, dpi=300)
        
        textos = []
        
        for i, imagen in enumerate(imagenes):
            print(f"Procesando página {i+1} de {len(imagenes)}...")
            
            # Convertir PIL a formato compatible con Tesseract
            texto = pytesseract.image_to_string(imagen, lang=idioma)
            textos.append(texto)
        
        # Unir todos los textos
        texto_completo = '\n\n--- Página {} ---\n\n'.join(
            [f'{i+1}' for i in range(len(textos))]
        ).format(*textos)
        
        return texto_completo
        
    except ImportError:
        print("Error: pytesseract no está instalado")
        return ""
    except Exception as e:
        print(f"Error extrayendo texto: {e}")
        return ""

def procesar_pdf_lote(carpeta_pdfs, dpi=300, formato='PNG'):
    """
    Procesar múltiples PDFs en una carpeta
    
    Args:
        carpeta_pdfs: Carpeta con archivos PDF
        dpi: Resolución DPI
        formato: Formato de salida
    
    Returns:
        Dict: {nombre_pdf: [rutas_imagenes]}
    """
    
    resultados = {}
    
    # Buscar archivos PDF en la carpeta
    for archivo in os.listdir(carpeta_pdfs):
        if archivo.lower().endswith('.pdf'):
            ruta_pdf = os.path.join(carpeta_pdfs, archivo)
            
            print(f"\nProcesando: {archivo}")
            
            # Convertir PDF
            imagenes, temp_dir = pdf_a_imagenes(ruta_pdf, dpi, formato)
            
            if imagenes:
                resultados[archivo] = {
                    'imagenes': imagenes,
                    'temp_dir': temp_dir,
                    'paginas': len(imagenes)
                }
    
    return resultados

if __name__ == "__main__":
    # Ejemplo de uso desde línea de comandos
    if len(sys.argv) < 2:
        print("""
Uso:
  python convertir_pdf.py <ruta_pdf> [dpi] [formato]
  
Ejemplos:
  python convertir_pdf.py documento.pdf
  python convertir_pdf.py documento.pdf 300 PNG
  python convertir_pdf.py documento.pdf 150 JPEG
  
Formato soportados: PNG, JPEG, TIFF
        """)
        sys.exit(1)
    
    ruta_pdf = sys.argv[1]
    dpi = int(sys.argv[2]) if len(sys.argv) > 2 else 300
    formato = sys.argv[3] if len(sys.argv) > 3 else 'PNG'
    
    if not os.path.exists(ruta_pdf):
        print(f"Error: Archivo no encontrado: {ruta_pdf}")
        sys.exit(1)
    
    # Convertir PDF
    imagenes, temp_dir = pdf_a_imagenes(ruta_pdf, dpi, formato)
    
    if imagenes:
        print(f"\nPDF convertido exitosamente.")
        print(f"Imágenes guardadas en: {temp_dir}")
        
        # Preguntar si quiere procesar con OCR
        respuesta = input("\n¿Deseas procesar las imágenes con OCR? (s/n): ")
        
        if respuesta.lower() == 's':
            try:
                import pytesseract
                
                for ruta_imagen in imagenes:
                    print(f"\nProcesando: {os.path.basename(ruta_imagen)}")
                    texto = pytesseract.image_to_string(ruta_imagen, lang='spa')
                    
                    # Guardar texto extraído
                    ruta_txt = ruta_imagen.replace(f'.{formato.lower()}', '.txt')
                    with open(ruta_txt, 'w', encoding='utf-8') as f:
                        f.write(texto)
                    
                    print(f"Texto guardado en: {ruta_txt}")
                    
            except ImportError:
                print("pytesseract no instalado. Instala con: pip install pytesseract")
    else:
        print("Error al convertir PDF")