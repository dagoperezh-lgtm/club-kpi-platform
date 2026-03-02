# *****************************************************************************
# SECCIÓN 1: NÚCLEO DE NORMALIZACIÓN Y CONVERSIÓN ARITMÉTICA
# *****************************************************************************
# Esta sección garantiza que los nombres de los deportistas coincidan 
# perfectamente (MatchKey) y que los tiempos de entrenamiento se conviertan 
# a minutos numéricos para poder calcular los KPIs de adherencia (TPI).

import streamlit as st
import pandas as pd
import numpy as np
import re
import io
import zipfile
import unicodedata
import matplotlib.pyplot as plt
import random
from datetime import time, datetime, timedelta
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# -----------------------------------------------------------------------------
# 1.1 NORMALIZACIÓN DE IDENTIDAD (ELIMINA EL ERROR DE NOMBRES CON TILDES)
# -----------------------------------------------------------------------------
def clean_string(text):
    """
    Función de Normalización Absoluta.
    Convierte 'Dagoberto Pérez' y 'DAGOBERTO PEREZ' en el mismo código único.
    Sin esto, el sistema crea filas duplicadas o llena el Maestro con ceros.
    """
    if text is None or pd.isna(text):
        return ""
    
    # Paso 1: Convertir a string, eliminar espacios en los extremos y pasar a MAYÚSCULAS
    nombre_sucio = str(text).strip().upper()
    
    # Paso 2: Normalizar usando NFKD (Descompone caracteres como 'é' en 'e' + '´')
    nombre_normalizado_nfkd = unicodedata.normalize('NFKD', nombre_sucio)
    
    # Paso 3: Filtrar solo los caracteres que no sean marcas de combinación (tildes)
    # Esto elimina físicamente el acento del caracter.
    nombre_limpio_final = "".join([c for c in nombre_normalizado_nfkd if not unicodedata.combining(c)])
    
    return nombre_limpio_final

# -----------------------------------------------------------------------------
# 1.2 CONVERSOR ARITMÉTICO UNIVERSAL (ELIMINA EL ERROR DE FORMATOS DE HORA)
# -----------------------------------------------------------------------------
def to_mins(valor):
    """
    Motor de Conversión de Ingeniería. 
    Transforma cualquier entrada (Texto de Strava, Hora de Excel, Decimal) 
    en un número entero de minutos para permitir el cálculo del TPI.
    """
    # Si la celda está vacía, el valor es 0
    if pd.isna(valor):
        return 0
    
    # Convertir a texto para análisis de patrones
    s = str(valor).strip().lower()
    
    # Lista de exclusión: valores que Strava o Excel ponen cuando no hay actividad
    lista_basura = ['--:--', '0', '', '00:00:00', 'nc', '0:00:00', 'nan', '0:00']
    if s in lista_basura:
        return 0
    
    try:
        # CASO A: El valor es un decimal de Excel (ej: 0.5 equivale a 12 horas)
        # Excel guarda el tiempo como una fracción del día (1.0 = 24h)
        if isinstance(valor, (float, int)) and valor < 1 and valor > 0:
            return int(round(valor * 1440))
        
        # CASO B: El valor ya es un objeto de tiempo de Python (datetime.time)
        if isinstance(valor, (time, datetime)):
            return (valor.hour * 60) + valor.minute
            
        # CASO C: El valor es un texto con formato de reloj (HH:MM:SS o HH:MM)
        if ':' in s:
            partes = s.split(':')
            horas = int(partes[0])
            # Tomamos los minutos, ignorando segundos si los hay
            minutos = int(partes[1].split('.')[0])
            return (horas * 60) + minutos
            
        # CASO D: El valor es el formato de texto de Strava (ej: '1h 22min' o '45min')
        # Buscamos patrones de horas 'h' y minutos 'min'
        patron_horas = re.search(r'(\d+)\s*h', s)
        patron_minutos = re.search(r'(\d+)\s*min', s)
        
        total_h = int(patron_horas.group(1)) * 60 if patron_horas else 0
        total_m = int(patron_minutos.group(1)) if patron_minutos else 0
        
        return total_h + total_m
        
    except Exception:
        # Si algo falla en la conversión, devolvemos 0 para no romper el programa
        return 0

# -----------------------------------------------------------------------------
# 1.3 FORMATEADOR VISUAL PARA REPORTES (HH:MM:SS)
# -----------------------------------------------------------------------------
def to_hhmmss_display(minutos):
    """
    Convierte los minutos numéricos de vuelta a un formato legible 
    para las tablas de los reportes Word.
    """
    horas = int(minutos // 60)
    minutos_restantes = int(minutos % 60)
    return f"{horas:02d}:{minutos_restantes:02d}:00"

# *****************************************************************************
# SECCIÓN 2: MOTORES DE EXTRACCIÓN Y PARSEO (STRAVA Y PLAN INDIVIDUAL)
# *****************************************************************************
# Esta sección lee los archivos Excel de entrada y los transforma en tablas
# estandarizadas. Garantiza que las columnas se identifiquen correctamente
# sin importar ligeras variaciones en los nombres generados por Strava.

# -----------------------------------------------------------------------------
# 2.1 PARSER DE EXCEL SEMANAL (CAPTURADOR DE DATOS REALES DE STRAVA)
# -----------------------------------------------------------------------------
def procesar_strava_excel(archivo_excel):
    """
    Escanea el Excel descargado de Strava buscando las columnas clave.
    Utiliza búsqueda semántica explícita para ser inmune a cambios de formato.
    """
    # Cargar el archivo Excel crudo en un DataFrame
    df_raw = pd.read_excel(archivo_excel)
    columnas_originales = df_raw.columns
    
    # Variables de control para almacenar los nombres reales de las columnas encontradas
    columna_nombre = None
    columna_natacion = None
    columna_ciclismo = None
    columna_trote = None
    
    # --- Búsqueda de la columna de Identidad (Deportista / Nombre) ---
    for col in columnas_originales:
        col_limpia = clean_string(str(col))
        if 'DEPORTISTA' in col_limpia or 'NOMBRE' in col_limpia or 'NAME' in col_limpia:
            columna_nombre = col
            break
            
    # Fallback de seguridad: Si no encuentra columna de nombre, asume que es la primera
    if columna_nombre is None:
        columna_nombre = columnas_originales[0]
        
    # --- Búsqueda de la columna de Natación ---
    for col in columnas_originales:
        col_limpia = clean_string(str(col))
        if 'NATAC' in col_limpia or 'PISCINA' in col_limpia or 'SWIM' in col_limpia:
            columna_natacion = col
            break
            
    # --- Búsqueda de la columna de Ciclismo ---
    for col in columnas_originales:
        col_limpia = clean_string(str(col))
        if 'BICI' in col_limpia or 'CICLIS' in col_limpia or 'RIDE' in col_limpia:
            columna_ciclismo = col
            break
            
    # --- Búsqueda de la columna de Trote ---
    for col in columnas_originales:
        col_limpia = clean_string(str(col))
        if 'TROTE' in col_limpia or 'RUN' in col_limpia or 'CARRERA' in col_limpia:
            columna_trote = col
            break

    # --- Construcción del DataFrame Limpio y Purificado ---
    # Aquí creamos la estructura que el resto del sistema consumirá (N_Mins_Real, etc.)
    df_limpio = pd.DataFrame()
    df_limpio['Deportista'] = df_raw[columna_nombre]
    
    # Aplicamos el motor aritmético (to_mins de la Sección 1) a cada columna detectada
    if columna_natacion is not None:
        df_limpio['N_Mins_Real'] = df_raw[columna_natacion].apply(to_mins)
    else:
        df_limpio['N_Mins_Real'] = 0
        
    if columna_ciclismo is not None:
        df_limpio['B_Mins_Real'] = df_raw[columna_ciclismo].apply(to_mins)
    else:
        df_limpio['B_Mins_Real'] = 0
        
    if columna_trote is not None:
        df_limpio['R_Mins_Real'] = df_raw[columna_trote].apply(to_mins)
    else:
        df_limpio['R_Mins_Real'] = 0
        
    # Calculamos el tiempo total real de la semana como la suma exacta de las disciplinas
    df_limpio['T_Mins_Real'] = df_limpio['N_Mins_Real'] + df_limpio['B_Mins_Real'] + df_limpio['R_Mins_Real']
    
    return df_limpio

# -----------------------------------------------------------------------------
# 2.2 PARSER DEL PLAN INDIVIDUAL (METAS ESPECÍFICAS)
# -----------------------------------------------------------------------------
def procesar_plan_individual(archivo_plan):
    """
    Carga el archivo Excel con las metas individuales de los deportistas.
    Asegura que la identidad se procese con la misma llave (MatchKey) para un cruce perfecto.
    """
    if archivo_plan is None:
        # Si no se carga un plan individual, devolvemos un DataFrame vacío
        # La Sección 3 detectará esto y aplicará las Metas Globales del Sidebar
        return pd.DataFrame()
        
    # Cargar el Excel de plan de entrenamiento
    df_plan = pd.read_excel(archivo_plan)
    
    # Verificación de seguridad: Evitar procesar archivos vacíos
    if df_plan.empty:
        return pd.DataFrame()
    
    # Asumimos estrictamente que la primera columna contiene la identidad del deportista
    columna_nombres_plan = df_plan.columns[0]
    
    # Generamos la llave maestra de cruce usando la normalización de la Sección 1
    df_plan['MatchKey'] = df_plan[columna_nombres_plan].apply(clean_string)
    
    return df_plan

# *****************************************************************************
# SECCIÓN 3: MOTOR NARRATIVO PRO CHILE (INTEGRIDAD TOTAL - 80+ FRASES)
# *****************************************************************************
# Este motor dota de inteligencia y variabilidad a los reportes Word.
# Utiliza un sistema de "pilas" (stacks) para asegurar que en un mismo reporte
# grupal no se repitan los comentarios entre los distintos deportistas.

PILAS_COMENTARIOS = {}

def obtener_frase_base(categoria, pool_frases):
    """
    Sistema anti-repetición. Extrae frases de una pila barajada y la 
    recarga automáticamente si se agota durante la generación del lote.
    """
    global PILAS_COMENTARIOS
    if categoria not in PILAS_COMENTARIOS or not PILAS_COMENTARIOS[categoria]:
        temp_pool = [str(f) for f in pool_frases]
        random.shuffle(temp_pool)
        PILAS_COMENTARIOS[categoria] = temp_pool
    return PILAS_COMENTARIOS[categoria].pop()

def generar_comentario(datos_de_fila, nombre_categoria, rank_posicion):
    """
    Inyecta los datos reales del deportista (Nombre y Tiempo) en las 
    plantillas narrativas según la disciplina evaluada.
    """
    atleta = str(datos_de_fila.get('Deportista', 'Atleta TYM'))
    
    # Dependiendo de la categoría, extraemos el tiempo formateado
    # Asumimos que datos_de_fila trae los minutos reales ya procesados
    if nombre_categoria == 'Natación':
        tiempo = to_hhmmss_display(datos_de_fila.get('N_Mins_Real', 0))
    elif nombre_categoria == 'Bicicleta':
        tiempo = to_hhmmss_display(datos_de_fila.get('B_Mins_Real', 0))
    elif nombre_categoria == 'Trote':
        tiempo = to_hhmmss_display(datos_de_fila.get('R_Mins_Real', 0))
    else:
        tiempo = to_hhmmss_display(datos_de_fila.get('T_Mins_Real', 0))
    
    # BANCO DE FRASES EXTENDIDO (SIN OMISIONES)
    pools = {
        'General': [
            "La disciplina de {atleta} es el motor del club; liderar con este volumen es pura entrega.",
            "Semana de consolidación para {atleta}. No solo es cantidad, es la calidad del tiempo acumulado.",
            "El compromiso de {atleta} se refleja en cada sesión. Un pilar fundamental del ranking hoy.",
            "Rendimiento de alto nivel. {atleta} entiende que la base del éxito es este volumen sostenido.",
            "Impresionante despliegue de {atleta}. Gestionar estas cargas requiere madurez deportiva.",
            "La constancia de {atleta} marca el paso del equipo. Una semana de trabajo impecable.",
            "Fuerza mental y física. {atleta} asimila el volumen semanal con una resiliencia notable.",
            "Evolución sostenida de {atleta}. Estar en el top general es fruto de una planificación seria.",
            "La ética de trabajo de {atleta} es envidiable. Cada hora sumada construye su mejor versión.",
            "Control total de la fatiga. {atleta} cierra la semana en lo más alto con mérito propio.",
            "{atleta} demuestra que la regularidad es el camino corto hacia los objetivos de temporada.",
            "Capacidad de carga superior. {atleta} lidera la tabla con una solvencia técnica admirable.",
            "Una semana brillante para {atleta}, demostrando una solidez física que inspira al resto.",
            "Disciplina inquebrantable. {atleta} se mantiene en la cima con un enfoque envidiable.",
            "{atleta} cierra la jornada con un volumen que refleja ambición y preparación rigurosa.",
            "Poderío aeróbico de {atleta}. Registrar estas horas es señal de una base muy robusta.",
            "Planificación ejecutada a la perfección por {atleta}. La consistencia es su mayor virtud.",
            "Rendimiento de punta. {atleta} encabeza el grupo con una capacidad de recuperación única.",
            "El volumen de {atleta} es el resultado de una mentalidad enfocada en la larga distancia.",
            "Foco y determinación. {atleta} asume el liderato semanal con una carga de trabajo sólida.",
            "Notable fondo físico de {atleta}. Su presencia en el podio general es garantía de perseverancia.",
            "{atleta} proyecta una temporada sólida manteniendo este ritmo de entrenamientos semanales.",
            "Sello de calidad TYM: {atleta} pone el trabajo necesario para destacar en la tabla general.",
            "Madurez competitiva. {atleta} sabe que el volumen es el cimiento de su rendimiento futuro.",
            "Gran lectura de cargas de {atleta}, logrando un volumen total que marca diferencias claras."
        ],
        'CV': [
            "Equilibrio milimétrico. {atleta} entrena con la precisión de quien no deja nada al azar.",
            "La polivalencia de {atleta} es su mayor ventaja. Simetría total en las tres áreas.",
            "Control de carga magistral. {atleta} distribuye su energía de forma balanceada.",
            "Triatlón en estado puro: {atleta} demuestra que dominar la transición es dominar el balance.",
            "Eficiencia técnica destacada. {atleta} logra que la simetría parezca sencilla pero es pura gestión.",
            "Versatilidad técnica. {atleta} no descuida ningún frente, fortaleciendo sus debilidades.",
            "Arquitectura de entrenamiento impecable. {atleta} refleja la esencia del deportista integral.",
            "Cero puntos débiles. {atleta} mantiene una paridad envidiable entre agua, bici y trote.",
            "Gestión inteligente de las cargas. {atleta} prioriza la salud y el equilibrio deportivo.",
            "Sincronía total. {atleta} asimila las tres disciplinas con una regularidad asombrosa.",
            "El balance de {atleta} es la clave para evitar lesiones y potenciar el rendimiento global.",
            "{atleta} demuestra que ser completo es más importante que ser rápido en una sola área.",
            "Madurez deportiva de {atleta}. Su coeficiente de variación es de los mejores del club.",
            "Planificación equilibrada de {atleta}. Cada disciplina recibe la atención que merece.",
            "Solidez transversal. {atleta} se consolida como uno de los atletas más balanceados.",
            "Precisión técnica en la distribución. {atleta} entrena con inteligencia y visión global.",
            "La armonía de {atleta} en las tres áreas es fruto de un compromiso técnico superior.",
            "{atleta} destaca por su capacidad de mantener la calidad sin importar el medio.",
            "Consistencia simétrica. {atleta} es el referente de equilibrio para el equipo hoy.",
            "Desarrollo armónico. {atleta} fortalece su base con una distribución de tiempo magistral."
        ],
        'Natación': [
            "Fluidez y potencia. Los {tiempo} de {atleta} en la piscina son el cimiento de su base.",
            "Dominio acuático. {atleta} marca la pauta con un volumen técnico de {tiempo} en piscina.",
            "Calidad en el agua. {atleta} suma {tiempo} de nado con una técnica cada vez más depurada.",
            "{atleta} lidera la fase acuática con {tiempo}, demostrando que la piscina es su fortaleza.",
            "Brazada eficiente y constante. {atleta} asimila {tiempo} de natación con gran solvencia.",
            "Resistencia hidrodinámica de {atleta}. Registrar {tiempo} en piscina es un hito importante.",
            "El agua no miente: {atleta} ha trabajado duro para lograr estos {tiempo} de volumen neto.",
            "Foco técnico en natación. {atleta} cierra con {tiempo}, consolidando su fase de apertura.",
            "Disciplina en la piscina. {atleta} no falla y suma {tiempo} de alta relevancia técnica.",
            "Progreso acuático visible. {atleta} domina su carril con {tiempo} de trabajo serio.",
            "{atleta} demuestra solidez en el agua, acumulando {tiempo} de nado de alta calidad.",
            "Eficiencia en cada largo. {atleta} optimiza sus {tiempo} en piscina para mejorar su fondo.",
            "Control de ritmo acuático. {atleta} suma {tiempo} de natación con una técnica sólida.",
            "Fuerza en la piscina. {atleta} proyecta una gran base aeróbica con sus {tiempo} actuales.",
            "Consistencia en el agua. {atleta} aprovecha sus {tiempo} en piscina para pulir detalles."
        ],
        'Bicicleta': [
            "Kilometraje de calidad. {atleta} construye su fortaleza sobre los pedales con {tiempo} de rodaje.",
            "El gran motor del equipo. {atleta} asimila la carga de ciclismo con resiliencia envidiable.",
            "Potencia y fondo. {atleta} devoró la ruta sumando {tiempo}, demostrando preparación superior.",
            "Solidez sobre ruedas. {atleta} aprovecha cada sesión para sumar {tiempo} de base aeróbica.",
            "El asfalto es el hábitat de {atleta}. Su volumen de {tiempo} en bici es pilar de su plan.",
            "Resistencia sobre el pedal. {atleta} acumula {tiempo} de calidad para blindar sus piernas.",
            "Ciclismo de alto impacto. {atleta} se sitúa como líder con {tiempo} de rodaje neto.",
            "Fuerza y cadencia. {atleta} gestiona sus {tiempo} en bicicleta con una madurez notable.",
            "Fondo inquebrantable. {atleta} suma {tiempo} en la ruta, clave para la larga distancia.",
            "Dominio del segmento de ciclismo. {atleta} marca el ritmo con {tiempo} de trabajo duro.",
            "Potencia aeróbica en ruta. {atleta} consolida sus {tiempo} de pedaleo con determinación.",
            "{atleta} demuestra que la bicicleta es su fuerte, acumulando {tiempo} de volumen masivo.",
            "Resiliencia sobre el sillín. {atleta} asimila {tiempo} de ciclismo con una solvencia técnica única.",
            "Gestión de potencia de {atleta}. Sus {tiempo} de rodaje son fundamentales para la temporada.",
            "Control y resistencia. {atleta} suma {tiempo} de bicicleta, blindando su motor aeróbico."
        ],
        'Trote': [
            "Zancada resiliente. Cerrar la semana con {tiempo} de impacto en el asfalto define el carácter de {atleta}.",
            "Resistencia específica. {atleta} domina la fase de carrera con una gestión de fatiga admirable.",
            "Persistencia técnica. {atleta} asimila el volumen de {tiempo} en running fortaleciendo su base.",
            "El asfalto premia la constancia. {atleta} cierra con {tiempo} de trote muy sólidos.",
            "Capacidad de cierre. {atleta} demuestra su fondo aeróbico con {tiempo} de carrera a pie.",
            "Impacto controlado y eficiente. {atleta} suma {tiempo} de trote, clave para su evolución.",
            "Running de alta gama. {atleta} se posiciona en el top con {tiempo} de asimilación de carga.",
            "Fortaleza en la carrera. {atleta} no cede y registra {tiempo} de volumen neto en el asfalto.",
            "Zancada potente y rítmica. {atleta} asume sus {tiempo} de trote con una técnica ejemplar.",
            "Resiliencia en cada kilómetro. {atleta} demuestra que el trote es donde se ganan las carreras.",
            "Gestión de la fatiga en asfalto. {atleta} completa sus {tiempo} de trote con gran madurez.",
            "Fuerza mental en la carrera. {atleta} suma {tiempo} netos, esenciales para su progresión.",
            "Eficiencia de zancada. {atleta} asimila {tiempo} de trote, cuidando la técnica en cada tramo.",
            "Consistencia en el running. {atleta} cierra la semana con {tiempo} de carga aeróbica sólida.",
            "Resistencia de punta. {atleta} marca diferencias en el asfalto con sus {tiempo} de volumen."
        ]
    }

    # Determinamos la llave de la categoría (Fallback a 'General' si hay error)
    cat_key = 'General' if nombre_categoria in ['Completos', 'General'] else nombre_categoria
    
    # Extraemos la plantilla
    frase_plantilla = str(obtener_frase_base(cat_key, pools.get(cat_key, pools['General'])))
    
    # Inyectamos los datos del deportista
    comentario_final = frase_plantilla.replace("{atleta}", atleta).replace("{tiempo}", tiempo)
    
    # Bonus de liderazgo: Si es el número 1 general, añadimos un reconocimiento extra
    if rank_posicion == 1 and cat_key == 'General':
        comentario_final = f"🏆 {comentario_final.replace(atleta, f'nuestro líder {atleta}')}"
        
    return comentario_final
    
# *****************************************************************************
# SECCIÓN 4: MOTOR DE CÁLCULO DE ADHERENCIA (TPI - REGLA 4.3)
# *****************************************************************************
# Esta sección es el corazón analítico del sistema. Cruza los datos reales 
# obtenidos de Strava con las metas individuales o globales para calcular el TPI
# (Índice de Rendimiento TYM). Garantiza que no existan valores nulos (0.0%).

def calcular_kpis_tym(df_real, df_plan, metas_globales):
    """
    Calcula la adherencia al entrenamiento (TPI).
    Implementa la 'Cascada de Metas': Si un atleta no tiene un plan individual
    cargado en el Excel, el sistema utilizará automáticamente las metas globales
    definidas en el panel lateral (Sidebar).
    """
    # 1. Asegurar la existencia de la llave de cruce en los datos reales
    df_real['MatchKey'] = df_real['Deportista'].apply(clean_string)
    
    # 2. Cruce de datos (Merge) con el Plan Individual
    if df_plan is not None and not df_plan.empty:
        # Aseguramos que el plan también tenga su llave de cruce
        # Asumimos que la columna 0 del plan es el nombre del deportista
        nombre_col_plan = df_plan.columns[0]
        df_plan['MatchKey'] = df_plan[nombre_col_plan].apply(clean_string)
        
        # Realizamos un LEFT JOIN para mantener a todos los deportistas reales,
        # incluso si no tienen un plan individual cargado en su fila.
        df_merged = pd.merge(df_real, df_plan, on='MatchKey', how='left', suffixes=('', '_plan'))
    else:
        # Si no se subió el archivo de plan individual, trabajamos con los datos reales
        df_merged = df_real.copy()

    # 3. Función interna para procesar la aritmética fila por fila (Deportista por Deportista)
    def aplicar_logica_tpi_fila(row):
        resultados_kpi = {}
        
        # ---------------------------------------------------------------------
        # --- A. EVALUACIÓN DE NATACIÓN ---
        # ---------------------------------------------------------------------
        # Obtener meta de horas (Cascada: Plan Individual -> Meta Global)
        horas_meta_natacion = row.get('Natacion_Hrs_plan')
        if pd.isna(horas_meta_natacion):
            horas_meta_natacion = metas_globales['N_H']
            
        # Obtener meta de sesiones (Cascada: Plan Individual -> Meta Global)
        sesiones_meta_natacion = row.get('Natacion_Ses_plan')
        if pd.isna(sesiones_meta_natacion):
            sesiones_meta_natacion = metas_globales['N_S']
            
        # Obtener el dato real procesado en la Sección 2
        minutos_reales_natacion = row.get('N_Mins_Real', 0)
        
        # Cálculo de Índices para Natación
        # VCI (Volume Compliance Index): Porcentaje de horas cumplidas
        vci_natacion = (minutos_reales_natacion / (horas_meta_natacion * 60)) * 100 if horas_meta_natacion > 0 else 0
        
        # SEI (Session Execution Index): Cumplimiento de sesiones
        # Si entrenó más de 0 minutos y la meta es mayor a 0, le damos el 100% de la sesión
        sei_natacion = (100 / sesiones_meta_natacion) if (minutos_reales_natacion > 0 and sesiones_meta_natacion > 0) else 0
        
        # TPI Natación (40% Volumen + 60% Sesiones), tope máximo de 115% de cumplimiento
        tpi_natacion_crudo = (vci_natacion * 0.4) + (sei_natacion * 0.6)
        resultados_kpi['TPI_Natacion'] = min(tpi_natacion_crudo, 115)
        resultados_kpi['Natacion_Plan_Hrs'] = horas_meta_natacion

        # ---------------------------------------------------------------------
        # --- B. EVALUACIÓN DE CICLISMO ---
        # ---------------------------------------------------------------------
        # Obtener meta de horas
        horas_meta_ciclismo = row.get('Ciclismo_Hrs_plan')
        if pd.isna(horas_meta_ciclismo):
            horas_meta_ciclismo = metas_globales['B_H']
            
        # Obtener meta de sesiones
        sesiones_meta_ciclismo = row.get('Ciclismo_Ses_plan')
        if pd.isna(sesiones_meta_ciclismo):
            sesiones_meta_ciclismo = metas_globales['B_S']
            
        # Obtener el dato real procesado
        minutos_reales_ciclismo = row.get('B_Mins_Real', 0)
        
        # Cálculo de Índices para Ciclismo
        vci_ciclismo = (minutos_reales_ciclismo / (horas_meta_ciclismo * 60)) * 100 if horas_meta_ciclismo > 0 else 0
        sei_ciclismo = (100 / sesiones_meta_ciclismo) if (minutos_reales_ciclismo > 0 and sesiones_meta_ciclismo > 0) else 0
        
        # TPI Ciclismo
        tpi_ciclismo_crudo = (vci_ciclismo * 0.4) + (sei_ciclismo * 0.6)
        resultados_kpi['TPI_Ciclismo'] = min(tpi_ciclismo_crudo, 115)
        resultados_kpi['Ciclismo_Plan_Hrs'] = horas_meta_ciclismo

        # ---------------------------------------------------------------------
        # --- C. EVALUACIÓN DE TROTE ---
        # ---------------------------------------------------------------------
        # Obtener meta de horas
        horas_meta_trote = row.get('Trote_Hrs_plan')
        if pd.isna(horas_meta_trote):
            horas_meta_trote = metas_globales['T_H']
            
        # Obtener meta de sesiones
        sesiones_meta_trote = row.get('Trote_Ses_plan')
        if pd.isna(sesiones_meta_trote):
            sesiones_meta_trote = metas_globales['T_S']
            
        # Obtener el dato real procesado
        minutos_reales_trote = row.get('R_Mins_Real', 0)
        
        # Cálculo de Índices para Trote
        vci_trote = (minutos_reales_trote / (horas_meta_trote * 60)) * 100 if horas_meta_trote > 0 else 0
        sei_trote = (100 / sesiones_meta_trote) if (minutos_reales_trote > 0 and sesiones_meta_trote > 0) else 0
        
        # TPI Trote
        tpi_trote_crudo = (vci_trote * 0.4) + (sei_trote * 0.6)
        resultados_kpi['TPI_Trote'] = min(tpi_trote_crudo, 115)
        resultados_kpi['Trote_Plan_Hrs'] = horas_meta_trote

        # ---------------------------------------------------------------------
        # --- D. CÁLCULO DE INDICADORES GLOBALES DEL DEPORTISTA ---
        # ---------------------------------------------------------------------
        # Promedio aritmético de la adherencia en las 3 disciplinas
        resultados_kpi['TPI_Global'] = np.mean([
            resultados_kpi['TPI_Natacion'], 
            resultados_kpi['TPI_Ciclismo'], 
            resultados_kpi['TPI_Trote']
        ])
        
        # Validación Estricta de "Triatleta Completo"
        # Debe registrar estrictamente más de 0 minutos en las TRES disciplinas.
        # Esto soluciona de raíz el fallo del Word donde reportaba "0 Triatletas Completos".
        resultados_kpi['Es_Completo'] = (minutos_reales_natacion > 0) and (minutos_reales_ciclismo > 0) and (minutos_reales_trote > 0)
        
        # Retornamos los cálculos transformados en una Serie de Pandas
        return pd.Series(resultados_kpi)

    # 4. Ejecución del motor: Aplicamos la lógica fila por fila a todos los deportistas
    df_resultados_kpi = df_merged.apply(aplicar_logica_tpi_fila, axis=1)
    
    # 5. Concatenación: Unimos los datos reales (Strava) con los KPI recién calculados
    df_final_procesado = pd.concat([df_merged, df_resultados_kpi], axis=1)
    
    return df_final_procesado

# *****************************************************************************
# SECCIÓN 5: MOTOR DE PERSISTENCIA Y ACTUALIZACIÓN DEL MAESTRO
# *****************************************************************************
# Esta sección actualiza el libro Excel Histórico preservando las semanas 
# anteriores. Implementa un blindaje contra columnas duplicadas (sufijos _x, _y) 
# y purifica los tipos de datos para evitar errores aritméticos en los totales.

def actualizar_maestro_tym(dict_dfs_originales, df_semana_actual, etiqueta_semana):
    """
    Actualiza hoja por hoja el archivo Maestro.
    Garantiza que la información se indexe correctamente a cada atleta.
    """
    dict_dfs_actualizados = {}
    
    # 1. Mapeo Extendido de Disciplinas (Sincronización de Pestañas)
    # Permite que el sistema encuentre la pestaña incluso si se llama "Ciclismo" o "Bicicleta"
    mapeo_hojas_a_datos = {
        'TIEMPO TOTAL': 'T_Mins_Real',
        'NATACION': 'N_Mins_Real',
        'NATACIÓN': 'N_Mins_Real',
        'BICICLETA': 'B_Mins_Real',
        'CICLISMO': 'B_Mins_Real', 
        'TROTE': 'R_Mins_Real',
        'RUNNING': 'R_Mins_Real',
        'CV': 'TPI_Global' # En la hoja de Coeficiente (CV) guardamos la Adherencia (TPI)
    }
    
    # Normalizamos los nombres de las pestañas reales del archivo subido
    hojas_reales_normalizadas = {clean_string(nombre): nombre for nombre in dict_dfs_originales.keys()}
    
    # 2. Bucle de Procesamiento: Hoja por Hoja
    for nombre_hoja_original in dict_dfs_originales.keys():
        nombre_normalizado = clean_string(nombre_hoja_original)
        
        # Verificamos si la pestaña actual es una disciplina a actualizar
        if nombre_normalizado in mapeo_hojas_a_datos:
            columna_datos_a_extraer = mapeo_hojas_a_datos[nombre_normalizado]
            df_hoja_historia = dict_dfs_originales[nombre_hoja_original].copy()
            
            # --- IDENTIFICACIÓN DEL DEPORTISTA ---
            col_identidad_maestro = 'Nombre' if 'Nombre' in df_hoja_historia.columns else \
                                   ('Deportista' if 'Deportista' in df_hoja_historia.columns else df_hoja_historia.columns[0])
            
            # Generamos la llave maestra (MatchKey) usando el normalizador de la Sección 1
            df_hoja_historia['MatchKey'] = df_hoja_historia[col_identidad_maestro].apply(clean_string)
            
            # --- BLINDAJE ANTI-DUPLICADOS (Evita Sem 08_x) ---
            # Si la columna ya existe en el maestro (ej: el usuario reprocesa la misma semana),
            # la eliminamos completamente antes del merge para asegurar una sobreescritura limpia.
            if etiqueta_semana in df_hoja_historia.columns:
                df_hoja_historia = df_hoja_historia.drop(columns=[etiqueta_semana])
            
            # --- PURIFICACIÓN ARITMÉTICA DE SEMANAS ANTERIORES ---
            # Para evitar TypeError al sumar, convertimos todas las 'Sem XX' a números puros.
            columnas_semanas_viejas = [c for c in df_hoja_historia.columns if str(c).startswith('Sem')]
            for col_vieja in columnas_semanas_viejas:
                # 'coerce' convierte textos irreconocibles a NaN, y fillna(0) los hace sumables.
                df_hoja_historia[col_vieja] = pd.to_numeric(df_hoja_historia[col_vieja], errors='coerce').fillna(0.0)
            
            # --- PREPARACIÓN DE LA NOVEDAD (DATOS DE LA SEMANA) ---
            df_novedad = df_semana_actual[['MatchKey', columna_datos_a_extraer]].copy()
            
            # Regla de Conversión a Excel (1.0 = 24 horas)
            if columna_datos_a_extraer != 'TPI_Global':
                # Convertimos minutos a fracción de día de Excel
                df_novedad[etiqueta_semana] = df_novedad[columna_datos_a_extraer].apply(
                    lambda x: (x / 1440.0) if pd.notna(x) else 0.0
                )
            else:
                # Para el TPI (Porcentaje), dividimos por 100
                df_novedad[etiqueta_semana] = df_novedad[columna_datos_a_extraer] / 100.0
            
            # Eliminamos duplicados en la novedad por seguridad extrema
            df_novedad = df_novedad.drop_duplicates(subset=['MatchKey'], keep='first')
            
            # --- UNIÓN DEL HISTÓRICO CON LA NOVEDAD ---
            # LEFT JOIN asegura que la estructura del Maestro mande.
            df_actualizado = pd.merge(
                df_hoja_historia, 
                df_novedad[['MatchKey', etiqueta_semana]], 
                on='MatchKey', 
                how='left'
            )
            
            # Rellenamos con 0 a los atletas antiguos que no entrenaron esta semana
            df_actualizado[etiqueta_semana] = df_actualizado[etiqueta_semana].fillna(0.0)
            
            # --- RECALCULO DE MÉTRICAS (TIEMPO ACUMULADO) ---
            columnas_todas_semanas = [c for c in df_actualizado.columns if str(c).startswith('Sem')]
            if columna_datos_a_extraer != 'TPI_Global':
                if 'Tiempo Acumulado' in df_actualizado.columns:
                    # La suma ya no fallará porque garantizamos que todas las columnas son numéricas
                    df_actualizado['Tiempo Acumulado'] = df_actualizado[columnas_todas_semanas].sum(axis=1)
                
                if 'Promedio' in df_actualizado.columns:
                    df_actualizado['Promedio'] = df_actualizado[columnas_todas_semanas].mean(axis=1)
            
            # Guardamos la pestaña actualizada en el diccionario final (borrando la llave técnica)
            dict_dfs_actualizados[nombre_hoja_original] = df_actualizado.drop(columns=['MatchKey'], errors='ignore')
            
        else:
            # Si la hoja no es de disciplinas (Ej: "Número de Semana", "Calendario"),
            # la traspasamos intacta sin alterarla.
            dict_dfs_actualizados[nombre_hoja_original] = dict_dfs_originales[nombre_hoja_original]
            
    return dict_dfs_actualizados

def guardar_maestro_excel(dict_dfs):
    """
    Convierte el diccionario de DataFrames en un archivo binario Excel (.xlsx).
    Utiliza xlsxwriter para mayor estabilidad y retención de formatos en archivos con muchas hojas.
    """
    buffer_salida = io.BytesIO()
    with pd.ExcelWriter(buffer_salida, engine='xlsxwriter') as writer:
        for nombre_pestaña, df_contenido in dict_dfs.items():
            df_contenido.to_excel(writer, sheet_name=nombre_pestaña, index=False)
    
    return buffer_salida.getvalue()
# =============================================================================
# FIN DE SECCIÓN 5
# =============================================================================

# *****************************************************************************
# SECCIÓN 6: ORQUESTADOR DE ENTREGABLES (GENERADOR WORD Y EXCEL EN ZIP)
# *****************************************************************************
# Esta sección toma los datos calculados y purificados para generar el 
# paquete final descargable. Garantiza que la lectura de minutos y porcentajes
# se plasme en los reportes sin errores de formato o valores vacíos.

def generar_entregables_finales(df_semanal_procesado, dict_maestro_actualizado, etiqueta_semana):
    """
    Construye en memoria (RAM) el archivo ZIP que contendrá:
    1. El Excel Maestro Histórico actualizado.
    2. El Reporte General del Club (Word).
    3. Una carpeta con las fichas individuales de cada deportista activo (Word).
    """
    # Creamos el buffer en memoria para el archivo ZIP
    zip_buffer_final = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer_final, "a", zipfile.ZIP_DEFLATED) as archivo_zip:
        
        # ---------------------------------------------------------------------
        # 6.1: EMPAQUETADO DEL EXCEL MAESTRO ACTUALIZADO
        # ---------------------------------------------------------------------
        # Llamamos a la función de la Sección 5 que convierte el diccionario en Excel
        binario_excel = guardar_maestro_excel(dict_maestro_actualizado)
        nombre_excel = f"01_Estadisticas_Actualizadas_{etiqueta_semana}.xlsx"
        archivo_zip.writestr(nombre_excel, binario_excel)
        
        # ---------------------------------------------------------------------
        # 6.2: GENERACIÓN DEL REPORTE GRUPAL DEL CLUB (WORD)
        # ---------------------------------------------------------------------
        doc_grupal = Document()
        doc_grupal.add_heading(f"Reporte Semanal Club TYM - {etiqueta_semana}", 0)
        
        # --- Cálculo de Métricas Grupales Auditadas ---
        total_deportistas_activos = len(df_semanal_procesado)
        
        # Filtramos a los triatletas completos (aquellos con la bandera Es_Completo == True)
        # Esta bandera fue estrictamente calculada en la Sección 4
        df_triatletas_completos = df_semanal_procesado[df_semanal_procesado['Es_Completo'] == True]
        total_completos = len(df_triatletas_completos)
        
        # Sumamos el tiempo total real procesado por todo el club
        minutos_totales_club = df_semanal_procesado['T_Mins_Real'].sum()
        
        # --- Redacción del Párrafo de Resumen ---
        parrafo_resumen = doc_grupal.add_paragraph()
        parrafo_resumen.add_run("Total deportistas registrados esta semana: ").bold = True
        parrafo_resumen.add_run(f"{total_deportistas_activos}\n")
        
        parrafo_resumen.add_run("Triatletas con entrenamiento completo (N/B/R): ").bold = True
        parrafo_resumen.add_run(f"{total_completos}\n")
        
        parrafo_resumen.add_run("Volumen total acumulado por el club: ").bold = True
        parrafo_resumen.add_run(f"{to_hhmmss_display(minutos_totales_club)}")

        # --- Construcción de la Tabla de Podio TPI (TOP 15) ---
        doc_grupal.add_heading("🏆 TOP 15 ADHERENCIA (TPI GLOBAL)", level=1)
        
        tabla_podio = doc_grupal.add_table(rows=1, cols=3)
        tabla_podio.style = 'Light Grid Accent 1'
        
        # Encabezados de la tabla
        celdas_encabezado = tabla_podio.rows[0].cells
        celdas_encabezado[0].text = 'Posición'
        celdas_encabezado[1].text = 'Deportista'
        celdas_encabezado[2].text = 'TPI Global %'
        
        # Ordenamos a los deportistas por su TPI Global de mayor a menor y sacamos los primeros 15
        df_ranking_top15 = df_semanal_procesado.sort_values(by='TPI_Global', ascending=False).head(15)
        
        # Iteramos sobre los 15 mejores para rellenar las filas
        posicion = 1
        for index, fila_deportista in df_ranking_top15.iterrows():
            celdas_fila = tabla_podio.add_row().cells
            celdas_fila[0].text = str(posicion)
            celdas_fila[1].text = str(fila_deportista['Deportista'])
            # Formateamos el TPI con un decimal
            celdas_fila[2].text = f"{fila_deportista['TPI_Global']:.1f}%"
            posicion += 1

        # Guardar el Reporte Grupal en el buffer y añadirlo al ZIP
        buffer_word_grupal = io.BytesIO()
        doc_grupal.save(buffer_word_grupal)
        archivo_zip.writestr(f"02_Reporte_General_{etiqueta_semana}.docx", buffer_word_grupal.getvalue())

        # ---------------------------------------------------------------------
        # 6.3: GENERACIÓN DE FICHAS INDIVIDUALES (WORD)
        # ---------------------------------------------------------------------
        # Iteramos sobre absolutamente todos los deportistas procesados
        for index, row_atleta in df_semanal_procesado.iterrows():
            
            # Solo generamos ficha si el deportista registró al menos 1 minuto de actividad
            if row_atleta['T_Mins_Real'] > 0:
                doc_individual = Document()
                doc_individual.add_heading(f"Análisis de Rendimiento: {row_atleta['Deportista']}", 0)
                
                # Resumen Principal
                parrafo_tpi = doc_individual.add_paragraph()
                parrafo_tpi.add_run("Tu Índice de Adherencia (TPI Global) esta semana: ").bold = True
                parrafo_tpi.add_run(f"{row_atleta['TPI_Global']:.1f}%")
                
                # --- Construcción de Tabla de Desglose por Disciplina ---
                tabla_desglose = doc_individual.add_table(rows=1, cols=4)
                tabla_desglose.style = 'Table Grid'
                
                celdas_cabecera_ind = tabla_desglose.rows[0].cells
                celdas_cabecera_ind[0].text = 'Disciplina'
                celdas_cabecera_ind[1].text = 'Tiempo Real'
                celdas_cabecera_ind[2].text = 'Meta (Plan)'
                celdas_cabecera_ind[3].text = 'TPI %'
                
                # Mapeo manual y explícito para evitar fallos de lectura de columnas
                matriz_disciplinas = [
                    ('Natación', 'N_Mins_Real', 'Natacion_Plan_Hrs', 'TPI_Natacion'),
                    ('Ciclismo', 'B_Mins_Real', 'Ciclismo_Plan_Hrs', 'TPI_Ciclismo'),
                    ('Trote', 'R_Mins_Real', 'Trote_Plan_Hrs', 'TPI_Trote')
                ]
                
                # Rellenamos una fila por cada disciplina
                for nombre_disc, col_real, col_meta, col_tpi in matriz_disciplinas:
                    celdas_datos = tabla_desglose.add_row().cells
                    celdas_datos[0].text = nombre_disc
                    # Formateo visual del tiempo real
                    celdas_datos[1].text = to_hhmmss_display(row_atleta[col_real])
                    # Formateo visual de la meta en horas
                    celdas_datos[2].text = f"{row_atleta[col_meta]:.1f}h"
                    # Formateo visual del TPI individual
                    celdas_datos[3].text = f"{row_atleta[col_tpi]:.1f}%"

                # --- Inyección del Motor Narrativo (Sección 3) ---
                doc_individual.add_heading("Evaluación Técnica", level=1)
                
                # Determinamos el ranking de este atleta para saber si es el líder
                posicion_ranking = df_ranking_top15.index[df_ranking_top15['Deportista'] == row_atleta['Deportista']].tolist()
                rango_actual = posicion_ranking[0] + 1 if posicion_ranking else 99
                
                # Generamos e inyectamos la frase
                comentario_experto = generar_comentario(row_atleta, 'General', rango_actual)
                doc_individual.add_paragraph(comentario_experto)
                
                # Guardamos la ficha en el buffer y la empaquetamos en la subcarpeta "Fichas" del ZIP
                buffer_word_individual = io.BytesIO()
                doc_individual.save(buffer_word_individual)
                
                # Usamos clean_string para asegurar que el nombre del archivo sea seguro para Windows/Mac
                nombre_archivo_seguro = f"Fichas/Ficha_{clean_string(row_atleta['Deportista'])}.docx"
                archivo_zip.writestr(nombre_archivo_seguro, buffer_word_individual.getvalue())

    # Reiniciamos el puntero del buffer ZIP al inicio para que Streamlit pueda descargarlo
    zip_buffer_final.seek(0)
    return zip_buffer_final
    
# *****************************************************************************
# SECCIÓN 7: INTERFAZ DE USUARIO Y ORQUESTACIÓN (STREAMLIT)
# *****************************************************************************
# Esta sección construye el panel de control lateral (Sidebar) y ejecuta 
# secuencialmente todos los motores anteriores cuando el usuario presiona el botón.

# 1. Título principal de la aplicación en el panel central
st.title("🏆 Motor de Adherencia y Estadísticas - Club TYM")
st.markdown("Procesamiento de datos de Strava, cálculo de TPI y generación de reportes premium.")

# 2. Construcción del Panel Lateral (Sidebar)
with st.sidebar:
    st.header("⚙️ 1. Carga de Archivos")
    st.markdown("Sube los archivos Excel necesarios para el procesamiento.")
    
    # Cargadores de archivos
    archivo_maestro = st.file_uploader("A. Excel Maestro (Histórico)", type=["xlsx", "xls"])
    archivo_strava = st.file_uploader("B. Excel Strava (Semana Actual)", type=["xlsx", "xls"])
    archivo_plan = st.file_uploader("C. Excel Plan Individual (Opcional)", type=["xlsx", "xls"])
    
    st.divider()
    
    st.header("🎯 2. Metas Globales (Fallback)")
    st.markdown("Se usarán automáticamente si un deportista no tiene un plan individual cargado.")
    
    # Diccionario explícito para almacenar las metas globales
    metas_globales_sidebar = {}
    
    st.subheader("Natación")
    metas_globales_sidebar['N_H'] = st.number_input("Horas Natación (Meta)", min_value=0.0, value=3.0, step=0.5)
    metas_globales_sidebar['N_S'] = st.number_input("Sesiones Natación (Meta)", min_value=1, value=3, step=1)
    
    st.subheader("Ciclismo")
    metas_globales_sidebar['B_H'] = st.number_input("Horas Ciclismo (Meta)", min_value=0.0, value=5.0, step=0.5)
    metas_globales_sidebar['B_S'] = st.number_input("Sesiones Ciclismo (Meta)", min_value=1, value=3, step=1)
    
    st.subheader("Trote")
    metas_globales_sidebar['T_H'] = st.number_input("Horas Trote (Meta)", min_value=0.0, value=3.0, step=0.5)
    metas_globales_sidebar['T_S'] = st.number_input("Sesiones Trote (Meta)", min_value=1, value=3, step=1)
    
    st.divider()
    
    st.header("🏷️ 3. Etiqueta de la Semana")
    etiqueta_semana_input = st.text_input("Nombre de la columna a generar (Ej: Sem 08)", value="Sem 08")

# 3. Inicialización del estado de sesión (Evita que el botón de descarga desaparezca)
if 'paquete_zip_generado' not in st.session_state:
    st.session_state['paquete_zip_generado'] = None

# 4. Botón de Ejecución Principal y Flujo Lógico
if st.button("🚀 PROCESAR JORNADA Y GENERAR ENTREGABLES", use_container_width=True):
    
    # Verificación de seguridad estricta: Archivos obligatorios
    if archivo_maestro is not None and archivo_strava is not None:
        
        with st.spinner("Ejecutando motores de purificación y cálculo... Por favor espera."):
            try:
                # --- PASO 1: Extracción ---
                # Ejecuta la Sección 2 para convertir el Excel de Strava en un DataFrame puro
                df_datos_reales = procesar_strava_excel(archivo_strava)
                
                # --- PASO 2: Extracción de Plan ---
                # Ejecuta la Sección 2 para el plan (devuelve DataFrame vacío si no hay archivo)
                df_datos_plan = procesar_plan_individual(archivo_plan)
                
                # --- PASO 3: Cálculo de KPI (El Corazón del Sistema) ---
                # Ejecuta la Sección 4 aplicando la Cascada de Metas y calculando el TPI
                df_calculado_kpi = calcular_kpis_tym(df_datos_reales, df_datos_plan, metas_globales_sidebar)
                
                # --- PASO 4: Lectura del Historial ---
                # Carga todas las pestañas del Maestro en la memoria RAM
                diccionario_maestro_original = pd.read_excel(archivo_maestro, sheet_name=None)
                
                # --- PASO 5: Actualización y Blindaje ---
                # Ejecuta la Sección 5: Elimina columnas duplicadas, purifica datos antiguos y pega la novedad
                diccionario_maestro_actualizado = actualizar_maestro_tym(
                    diccionario_maestro_original, 
                    df_calculado_kpi, 
                    etiqueta_semana_input
                )
                
                # --- PASO 6: Generación de Entregables ---
                # Ejecuta la Sección 6: Crea el ZIP, redacta el Word Grupal y las Fichas Individuales
                buffer_zip_final = generar_entregables_finales(
                    df_calculado_kpi, 
                    diccionario_maestro_actualizado, 
                    etiqueta_semana_input
                )
                
                # Guardamos el ZIP en la memoria permanente de la sesión
                st.session_state['paquete_zip_generado'] = buffer_zip_final
                
                st.success("✅ ¡Procesamiento completado! Se han purificado los datos y generado los KPIs exitosamente.")
                
            except Exception as e:
                # Capturador de errores críticos para evitar que la pantalla se quede en blanco
                st.error(f"❌ Ocurrió un error crítico durante el procesamiento: {str(e)}")
                st.error("Revisa que los archivos de Excel tengan los nombres de columnas correctos.")
    else:
        # Mensaje de advertencia si el usuario olvidó cargar los archivos
        st.warning("⚠️ Debes cargar obligatoriamente el 'Excel Maestro' y el 'Excel Strava' en el panel izquierdo para continuar.")

# 5. Botón de Descarga (Renderizado persistente)
if st.session_state['paquete_zip_generado'] is not None:
    st.download_button(
        label="📥 DESCARGAR PAQUETE FINAL (ZIP)",
        data=st.session_state['paquete_zip_generado'],
        file_name=f"Entregables_TYM_{etiqueta_semana_input}.zip",
        mime="application/zip",
        use_container_width=True
    )
