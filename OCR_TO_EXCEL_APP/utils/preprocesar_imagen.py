#!/usr/bin/env python3
"""
Herramientas de preprocesamiento de imágenes para OCR
"""

import cv2
import numpy as np
from PIL import Image, ImageEnhance, ImageFilter
import os

def mejorar_imagen_ocr(ruta_imagen, config=None):
    """
    Preprocesar imagen para mejorar resultados de OCR
    
    Args:
        ruta_imagen: Ruta de la imagen a procesar
        config: Diccionario con configuración de procesamiento
    
    Returns:
        PIL.Image: Imagen procesada
    """
    
    # Configuración por defecto
    default_config = {
        'grayscale': True,
        'contrast': 1.5,
        'brightness': 1.0,
        'denoise': True,
        'threshold': 'adaptive',
        'deskew': True,
        'remove_shadows': False,
        'enhance_edges': False
    }
    
    if config:
        default_config.update(config)
    
    config = default_config
    
    try:
        # Abrir imagen con PIL
        img_pil = Image.open(ruta_imagen)
        
        # Convertir a RGB si es necesario
        if img_pil.mode != 'RGB':
            img_pil = img_pil.convert('RGB')
        
        # 1. Convertir a escala de grises si está configurado
        if config['grayscale']:
            img_pil = img_pil.convert('L')
        
        # 2. Ajustar brillo y contraste
        enhancer = ImageEnhance.Brightness(img_pil)
        img_pil = enhancer.enhance(config['brightness'])
        
        enhancer = ImageEnhance.Contrast(img_pil)
        img_pil = enhancer.enhance(config['contrast'])
        
        # Convertir a OpenCV para procesamiento avanzado
        img_cv = np.array(img_pil)
        
        if len(img_cv.shape) == 2:  # Si es escala de grises
            img_cv = cv2.cvtColor(img_cv, cv2.COLOR_GRAY2BGR)
        
        # 3. Reducir ruido
        if config['denoise']:
            img_cv = cv2.medianBlur(img_cv, 3)
        
        # 4. Umbralización
        gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)
        
        if config['threshold'] == 'adaptive':
            # Umbral adaptativo
            thresh = cv2.adaptiveThreshold(
                gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                cv2.THRESH_BINARY, 11, 2
            )
        elif config['threshold'] == 'otsu':
            # Umbral Otsu
            _, thresh = cv2.threshold(
                gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU
            )
        else:
            # Umbral simple
            _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY)
        
        # 5. Enderezar imagen (deskew)
        if config['deskew']:
            thresh = corregir_inclinacion(thresh)
        
        # 6. Mejorar bordes
        if config['enhance_edges']:
            thresh = cv2.Canny(thresh, 50, 150)
        
        # 7. Eliminar sombras (si está habilitado)
        if config['remove_shadows']:
            thresh = eliminar_sombras(thresh)
        
        # Convertir de vuelta a PIL
        img_procesada = Image.fromarray(thresh)
        
        return img_procesada
        
    except Exception as e:
        print(f"Error procesando imagen: {e}")
        # Devolver imagen original si hay error
        return Image.open(ruta_imagen)

def corregir_inclinacion(imagen):
    """
    Corregir inclinación de texto en imagen
    
    Args:
        imagen: Imagen en escala de grises
    
    Returns:
        Imagen corregida
    """
    # Encontrar contornos de texto
    coords = np.column_stack(np.where(imagen > 0))
    
    if len(coords) < 10:  # No hay suficiente texto para corregir
        return imagen
    
    # Calcular ángulo de inclinación
    angle = cv2.minAreaRect(coords)[-1]
    
    if angle < -45:
        angle = 90 + angle
    elif angle > 45:
        angle = angle - 90
    
    # Rotar imagen para corregir
    if abs(angle) > 1.0:  # Solo corregir si hay inclinación significativa
        (h, w) = imagen.shape[:2]
        center = (w // 2, h // 2)
        M = cv2.getRotationMatrix2D(center, angle, 1.0)
        imagen = cv2.warpAffine(imagen, M, (w, h),
                               flags=cv2.INTER_CUBIC,
                               borderMode=cv2.BORDER_REPLICATE)
    
    return imagen

def eliminar_sombras(imagen):
    """
    Eliminar sombras de imagen
    
    Args:
        imagen: Imagen en escala de grises
    
    Returns:
        Imagen sin sombras
    """
    # Dilatar y erosionar para eliminar sombras
    rgb_planes = cv2.split(imagen)
    result_planes = []
    
    for plane in rgb_planes:
        dilated_img = cv2.dilate(plane, np.ones((7,7), np.uint8))
        bg_img = cv2.medianBlur(dilated_img, 21)
        diff_img = 255 - cv2.absdiff(plane, bg_img)
        result_planes.append(diff_img)
    
    result = cv2.merge(result_planes)
    return result

def redimensionar_imagen(imagen, ancho_max=2000, alto_max=2000):
    """
    Redimensionar imagen manteniendo proporción
    
    Args:
        imagen: PIL Image
        ancho_max: Ancho máximo
        alto_max: Alto máximo
    
    Returns:
        PIL.Image redimensionada
    """
    ancho, alto = imagen.size
    
    if ancho > ancho_max or alto > alto_max:
        ratio = min(ancho_max/ancho, alto_max/alto)
        nuevo_ancho = int(ancho * ratio)
        nuevo_alto = int(alto * ratio)
        imagen = imagen.resize((nuevo_ancho, nuevo_alto), Image.Resampling.LANCZOS)
    
    return imagen

def extraer_tabla_imagen(ruta_imagen):
    """
    Intentar extraer tabla de imagen usando detección de bordes
    
    Args:
        ruta_imagen: Ruta de la imagen
    
    Returns:
        Lista de imágenes de celdas detectadas
    """
    img = cv2.imread(ruta_imagen)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    
    # Detectar bordes
    edges = cv2.Canny(gray, 50, 150, apertureSize=3)
    
    # Detectar líneas
    lines = cv2.HoughLinesP(edges, 1, np.pi/180, 100, minLineLength=100, maxLineGap=10)
    
    if lines is not None:
        # Dibujar líneas detectadas
        for line in lines:
            x1, y1, x2, y2 = line[0]
            cv2.line(img, (x1, y1), (x2, y2), (0, 255, 0), 2)
    
    return Image.fromarray(cv2.cvtColor(img, cv2.COLOR_BGR2RGB))

if __name__ == "__main__":
    # Ejemplo de uso
    import sys
    
    if len(sys.argv) > 1:
        imagen = mejorar_imagen_ocr(sys.argv[1])
        imagen.save("imagen_procesada.png")
        print("Imagen procesada guardada como 'imagen_procesada.png'")
    else:
        print("Uso: python preprocesar_imagen.py <ruta_imagen>")