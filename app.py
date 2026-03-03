# *****************************************************************************
# SECCIÓN 1: NÚCLEO DE NORMALIZACIÓN Y CONVERSIÓN (VERSIÓN EXTENDIDA HH:MM)
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
import random
from datetime import time, datetime

# Librerías de Word
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
# 1.2 CONVERSOR ARITMÉTICO UNIVERSAL (MOTOR ADAPTADO A HH:MM Y TIMEDELTA)
# -----------------------------------------------------------------------------
def to_mins(valor):
    """
    Motor de Conversión de Ingeniería. 
    Transforma cualquier entrada (Texto HH:MM, Timedelta, Decimal) 
    en un número entero de minutos para permitir el cálculo del TPI.
    """
    # Si la celda está vacía, el valor es 0
    if pd.isna(valor):
        return 0
        
    # SALVAVIDAS CRÍTICO 1: Manejo nativo de objetos Timedelta de Pandas
    # Esto evita el fallo de "0 days 02:30:00" que arrojaba ceros en las versiones previas
    if isinstance(valor, pd.Timedelta):
        return int(valor.total_seconds() // 60)
    
    # SALVAVIDAS CRÍTICO 2: El valor es un decimal de Excel (ej: 0.5 equivale a 12 horas)
    # Excel guarda el tiempo como una fracción del día (1.0 = 24h)
    if isinstance(valor, (float, int)):
        if valor == 0:
            return 0
        if valor < 1 and valor > 0:
            return int(round(valor * 1440))
        return int(valor)
        
    # SALVAVIDAS 3: El valor ya es un objeto de tiempo de Python (datetime.time)
    if isinstance(valor, (time, datetime)):
        return (valor.hour * 60) + valor.minute
    
    # Convertir a texto para análisis de patrones
    s = str(valor).strip().lower()
    
    # Lista de exclusión: valores que Strava o Excel ponen cuando no hay actividad
    lista_basura = ['--:--', '0', '', '00:00', '00:00:00', 'nc', 'nan', 'none']
    if s in lista_basura:
        return 0
    
    try:
        # Fallback de seguridad adicional para Timedeltas que llegan como texto
        if 'days' in s or 'day' in s:
            partes_espacio = s.split()
            dias = int(partes_espacio[0])
            tiempo_str = partes_espacio[-1]
            t_parts = tiempo_str.split(':')
            horas = int(t_parts[0])
            minutos = int(t_parts[1].split('.')[0])
            return (dias * 1440) + (horas * 60) + minutos

        # FORMATO PRINCIPAL ESPERADO: HH:MM (Opcionalmente HH:MM:SS)
        if ':' in s:
            partes = s.split(':')
            horas = int(partes[0])
            # Al tomar fijamente el índice [1], funciona perfecto tanto si el archivo
            # trae "02:30" (2 partes) como si trae "02:30:00" (3 partes).
            minutos = int(partes[1].split('.')[0])
            return (horas * 60) + minutos
            
        # SALVAVIDAS 4: Formato de texto de Strava (ej: '1h 22m' o '45min')
        patron_horas = re.search(r'(\d+)\s*h', s)
        patron_minutos = re.search(r'(\d+)\s*m', s)
        
        total_h = int(patron_horas.group(1)) * 60 if patron_horas else 0
        total_m = int(patron_minutos.group(1)) if patron_minutos else 0
        
        if total_h > 0 or total_m > 0:
            return total_h + total_m
            
        # SALVAVIDAS 5: Si por error de tipeo es solo un número suelto (ej: "45")
        if s.isdigit():
            return int(s)
            
    except Exception:
        # Si algo falla en la conversión de esta celda, devolvemos 0 para no romper el lote
        return 0
        
    return 0

# -----------------------------------------------------------------------------
# 1.3 FORMATEADOR VISUAL PARA REPORTES (HH:MM ESTRICTO)
# -----------------------------------------------------------------------------
def to_hhmm_display(minutos):
    """
    Convierte los minutos numéricos de vuelta a un formato legible 
    para las tablas de los reportes Word. Limitado a HH:MM por instrucción.
    """
    horas = int(minutos // 60)
    minutos_restantes = int(minutos % 60)
    
    # Entrega formato 02:30 (sin los segundos finales)
    return f"{horas:02d}:{minutos_restantes:02d}"


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
    # Cargar el archivo Excel crudo en un DataFrame de Pandas
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
            
    # Fallback de seguridad extrema: Si no encuentra columna de nombre, asume que es la primera
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
    # Aquí creamos la estructura interna que el resto del sistema consumirá sin errores
    df_limpio = pd.DataFrame()
    
    # Asignamos la identidad
    df_limpio['Deportista'] = df_raw[columna_nombre]
    
    # Aplicamos el motor aritmético (to_mins de la Sección 1) a cada celda de las disciplinas
    if columna_natacion is not None:
        df_limpio['N_Mins_Real'] = df_raw[columna_natacion].apply(to_mins)
    else:
        # Si Strava no trae la columna, rellenamos con 0 explícitamente
        df_limpio['N_Mins_Real'] = 0
        
    if columna_ciclismo is not None:
        df_limpio['B_Mins_Real'] = df_raw[columna_ciclismo].apply(to_mins)
    else:
        df_limpio['B_Mins_Real'] = 0
        
    if columna_trote is not None:
        df_limpio['R_Mins_Real'] = df_raw[columna_trote].apply(to_mins)
    else:
        df_limpio['R_Mins_Real'] = 0
        
    # Calculamos el tiempo total real de la semana como la suma exacta de las tres disciplinas
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
        # Si el usuario decide no cargar un plan individual, devolvemos un DataFrame vacío.
        # El sistema detectará esto y aplicará automáticamente las Metas Globales del Sidebar.
        return pd.DataFrame()
        
    # Cargar el Excel de plan de entrenamiento
    df_plan = pd.read_excel(archivo_plan)
    
    # Verificación de seguridad: Evitar procesar archivos que vengan en blanco
    if df_plan.empty:
        return pd.DataFrame()
    
    # Asumimos estrictamente que la primera columna contiene la identidad del deportista
    columna_nombres_plan = df_plan.columns[0]
    
    # Generamos la llave maestra de cruce usando la normalización de la Sección 1
    df_plan['MatchKey'] = df_plan[columna_nombres_plan].apply(clean_string)
    
    return df_plan
    
# *****************************************************************************
# SECCIÓN 3: MOTOR NARRATIVO PRO CHILE (INTEGRIDAD TOTAL - 90+ FRASES)
# *****************************************************************************
# Este motor dota de inteligencia y variabilidad a los reportes Word.
# Utiliza un sistema de "pilas" (stacks) para asegurar que en un mismo reporte
# grupal no se repitan los comentarios entre los distintos deportistas.

import random

PILAS_COMENTARIOS = {}

def obtener_frase_base(categoria, pool_frases):
    """
    Sistema anti-repetición. Extrae frases de una pila barajada y la 
    recarga automáticamente si se agota durante la generación del lote.
    """
    global PILAS_COMENTARIOS
    if categoria not in PILAS_COMENTARIOS or not PILAS_COMENTARIOS[categoria]:
        # Creamos una copia de la lista para no alterar la original y la barajamos
        temp_pool = [str(f) for f in pool_frases]
        random.shuffle(temp_pool)
        PILAS_COMENTARIOS[categoria] = temp_pool
    return PILAS_COMENTARIOS[categoria].pop()

def generar_comentario(datos_de_fila, nombre_categoria, rank_posicion):
    """
    Inyecta los datos reales del deportista (Nombre y Tiempo) en las 
    plantillas narrativas según la disciplina evaluada.
    Incluye lógica de podio para no llamar "líder" al segundo o tercer lugar.
    """
    atleta = str(datos_de_fila.get('Deportista', 'Atleta TYM'))
    
    # Dependiendo de la categoría, extraemos el tiempo formateado en HH:MM
    if nombre_categoria == 'Natación':
        tiempo = to_hhmm_display(datos_de_fila.get('N_Mins_Real', 0))
    elif nombre_categoria == 'Bicicleta':
        tiempo = to_hhmm_display(datos_de_fila.get('B_Mins_Real', 0))
    elif nombre_categoria == 'Trote':
        tiempo = to_hhmm_display(datos_de_fila.get('R_Mins_Real', 0))
    else:
        tiempo = to_hhmm_display(datos_de_fila.get('T_Mins_Real', 0))
    
    # BANCO DE FRASES EXTENDIDO (VERSIÓN ABSOLUTA SIN OMISIONES + TPI)
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
        'TPI': [
            "Ejecución impecable del plan. {atleta} demuestra una disciplina táctica de alto nivel.",
            "El entrenamiento invisible se hace visible aquí. Cumplimiento perfecto de las metas para {atleta}.",
            "Adherencia total. {atleta} respeta la planificación al pie de la letra y maximiza sus cargas.",
            "No es solo entrenar duro, es entrenar inteligente. {atleta} clava los porcentajes del plan.",
            "Respeto absoluto por las sesiones asignadas. {atleta} es un reloj suizo esta semana.",
            "La constancia vence al talento. {atleta} cierra con un nivel de cumplimiento envidiable.",
            "Planificación asimilada al máximo. {atleta} demuestra madurez para seguir las instrucciones técnicas.",
            "Disciplina inquebrantable. El TPI de {atleta} refleja un compromiso total con su propio proceso.",
            "Gestión perfecta de la agenda deportiva. {atleta} marca el estándar de adherencia del club.",
            "Cero excusas, puro cumplimiento. {atleta} se ajusta a la meta semanal con precisión quirúrgica.",
            "La estrategia da frutos cuando se respeta. {atleta} consolida su semana cumpliendo a cabalidad."
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

    # Determinamos la llave de la categoría (Fallback a 'General' si el nombre no cuadra exacto)
    cat_key = 'TPI' if nombre_categoria == 'TPI' else ('General' if nombre_categoria in ['Completos', 'General'] else nombre_categoria)
    
    # Extraemos la plantilla aleatoria
    frase_plantilla = str(obtener_frase_base(cat_key, pools.get(cat_key, pools['General'])))
    
    # Inyectamos los datos del deportista
    comentario_final = frase_plantilla.replace("{atleta}", atleta).replace("{tiempo}", tiempo)
    
    # LÓGICA DE PODIO: Filtrado estricto de palabras exclusivas para el 1° Lugar
    if rank_posicion > 1:
        comentario_final = comentario_final.replace("liderar", "destacar").replace("lidera", "destaca en")\
            .replace("líder", "referente").replace("en lo más alto", "en el podio")\
            .replace("en el top", "en el podio").replace("en la cima", "entre los mejores")\
            .replace("encabeza", "brilla en")
            
    # Bonus de liderazgo: Si es el número 1 del ranking, añadimos un reconocimiento extra
    if rank_posicion == 1 and cat_key in ['General', 'TPI']:
        comentario_final = f"🏆 {comentario_final.replace(atleta, f'nuestro líder {atleta}')}"
        
    return comentario_final
    
# *****************************************************************************
# SECCIÓN 4: MOTOR DE CÁLCULO DE ADHERENCIA (TPI - REGLA 4.3)
# *****************************************************************************
# Esta sección es el corazón analítico del sistema. Cruza los datos reales 
# obtenidos de Strava con las metas individuales o globales para calcular el TPI
# (Índice de Rendimiento TYM). Garantiza que no existan valores nulos (0.0%)
# ni errores de división por cero.

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
        # Si no se subió el archivo de plan individual, trabajamos directamente con los datos reales
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
        # VCI (Volume Compliance Index): Porcentaje de horas cumplidas respecto al plan
        if horas_meta_natacion > 0:
            vci_natacion = (minutos_reales_natacion / (horas_meta_natacion * 60)) * 100
        else:
            vci_natacion = 0
            
        # SEI (Session Execution Index): Cumplimiento de sesiones
        # Si el atleta entrenó más de 0 minutos y la meta es mayor a 0, asimilamos el 100% de la sesión
        if minutos_reales_natacion > 0 and sesiones_meta_natacion > 0:
            sei_natacion = (100 / sesiones_meta_natacion)
        else:
            sei_natacion = 0
        
        # TPI Natación (Regla de negocio: 40% Volumen + 60% Sesiones), tope máximo de 115% de cumplimiento
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
        if horas_meta_ciclismo > 0:
            vci_ciclismo = (minutos_reales_ciclismo / (horas_meta_ciclismo * 60)) * 100
        else:
            vci_ciclismo = 0
            
        if minutos_reales_ciclismo > 0 and sesiones_meta_ciclismo > 0:
            sei_ciclismo = (100 / sesiones_meta_ciclismo)
        else:
            sei_ciclismo = 0
        
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
        if horas_meta_trote > 0:
            vci_trote = (minutos_reales_trote / (horas_meta_trote * 60)) * 100
        else:
            vci_trote = 0
            
        if minutos_reales_trote > 0 and sesiones_meta_trote > 0:
            sei_trote = (100 / sesiones_meta_trote)
        else:
            sei_trote = 0
        
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
        if (minutos_reales_natacion > 0) and (minutos_reales_ciclismo > 0) and (minutos_reales_trote > 0):
            resultados_kpi['Es_Completo'] = True
        else:
            resultados_kpi['Es_Completo'] = False
        
        # Retornamos los cálculos transformados en una Serie de Pandas
        return pd.Series(resultados_kpi)

    # 4. Ejecución del motor: Aplicamos la lógica fila por fila a todos los deportistas
    df_resultados_kpi = df_merged.apply(aplicar_logica_tpi_fila, axis=1)
    
    # 5. Concatenación: Unimos los datos reales (Strava) con los KPI recién calculados
    df_final_procesado = pd.concat([df_merged, df_resultados_kpi], axis=1)
    
    return df_final_procesado


# *****************************************************************************
# SECCIÓN 5: MOTOR DE PERSISTENCIA Y ACTUALIZACIÓN DEL MAESTRO (FASE 1 - EXPANDIDA)
# *****************************************************************************
# Esta sección actualiza el libro Excel Histórico preservando las semanas
# anteriores. Implementa un blindaje contra columnas duplicadas (sufijos _x, _y),
# aplica la conversión a formato HH:MM estricto, reordena las columnas de 
# semanas, inyecta la hoja de KPIs, redondea el CV a 2 decimales y aplica
# un ordenamiento estricto y jerárquico a las pestañas del archivo final.

def actualizar_maestro_tym(dict_dfs_originales, df_semana_actual, etiqueta_semana):
    """
    Actualiza hoja por hoja el archivo Maestro.
    Garantiza que la información se indexe correctamente a cada atleta
    y que la historia no se borre ni se convierta en ceros, formateando todo a HH:MM.
    """
    dict_dfs_actualizados = {}

    # 1. Mapeo Extendido de Disciplinas (Sincronización de Pestañas)
    mapeo_hojas_a_datos = {
        'TIEMPO TOTAL': 'T_Mins_Real',
        'NATACION': 'N_Mins_Real',
        'NATACIÓN': 'N_Mins_Real',
        'BICICLETA': 'B_Mins_Real',
        'CICLISMO': 'B_Mins_Real',
        'TROTE': 'R_Mins_Real',
        'RUNNING': 'R_Mins_Real',
        'CV': 'TPI_Global'
    }

    # Normalizamos los nombres de las pestañas reales del archivo subido
    hojas_reales_normalizadas = {clean_string(nombre): nombre for nombre in dict_dfs_originales.keys()}

    # Diccionario para mapear MatchKey a Nombres reales de la semana actual (Para atletas nuevos)
    mapeo_nombres_nuevos = df_semana_actual.set_index('MatchKey')['Deportista'].to_dict()

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

            # Generamos la llave maestra (MatchKey) para un cruce seguro
            df_hoja_historia['MatchKey'] = df_hoja_historia[col_identidad_maestro].apply(clean_string)

            # --- BLINDAJE ANTI-DUPLICADOS Y LIMPIEZA DE BASURA ---
            columnas_basura = [c for c in df_hoja_historia.columns if etiqueta_semana in str(c) or str(c).endswith('_x') or str(c).endswith('_y')]
            if columnas_basura:
                df_hoja_historia = df_hoja_historia.drop(columns=columnas_basura)

            # --- SALVACIÓN Y PURIFICACIÓN DEL HISTORIAL ---
            columnas_semanas_viejas = [c for c in df_hoja_historia.columns if str(c).startswith('Sem')]

            for col_vieja in columnas_semanas_viejas:
                if columna_datos_a_extraer != 'TPI_Global':
                    df_hoja_historia[col_vieja] = df_hoja_historia[col_vieja].apply(to_mins)
                else:
                    # En la hoja CV forzamos a numérico y aplicamos el límite de 2 decimales
                    df_hoja_historia[col_vieja] = pd.to_numeric(df_hoja_historia[col_vieja], errors='coerce').fillna(0.0).round(2)

            # --- PREPARACIÓN DE LA NOVEDAD (DATOS DE LA SEMANA) ---
            df_novedad = df_semana_actual[['MatchKey', columna_datos_a_extraer]].copy()

            if columna_datos_a_extraer != 'TPI_Global':
                df_novedad[etiqueta_semana] = df_novedad[columna_datos_a_extraer]
            else:
                # La novedad de CV también se redondea a 2 decimales
                df_novedad[etiqueta_semana] = (df_novedad[columna_datos_a_extraer] / 100.0).round(2)

            # Eliminamos duplicados
            df_novedad = df_novedad.drop_duplicates(subset=['MatchKey'], keep='first')

            # --- UNIÓN DEL HISTÓRICO CON LA NOVEDAD ---
            df_actualizado = pd.merge(
                df_hoja_historia,
                df_novedad[['MatchKey', etiqueta_semana]],
                on='MatchKey',
                how='outer'
            )

            df_actualizado[etiqueta_semana] = df_actualizado[etiqueta_semana].fillna(0.0)

            # --- GESTIÓN DE ATLETAS NUEVOS ---
            mask_nuevos = df_actualizado[col_identidad_maestro].isna()
            df_actualizado.loc[mask_nuevos, col_identidad_maestro] = df_actualizado.loc[mask_nuevos, 'MatchKey'].map(mapeo_nombres_nuevos)

            for col_vieja in columnas_semanas_viejas:
                df_actualizado[col_vieja] = df_actualizado[col_vieja].fillna(0.0)

            # --- ORDENAMIENTO DE COLUMNAS DE SEMANAS ---
            columnas_fijas = [c for c in df_actualizado.columns if not str(c).startswith('Sem')]
            columnas_temporales = sorted([c for c in df_actualizado.columns if str(c).startswith('Sem')])
            df_actualizado = df_actualizado[columnas_fijas + columnas_temporales]

            # --- RECALCULO DE MÉTRICAS (TIEMPO ACUMULADO) ---
            columnas_todas_semanas = [c for c in df_actualizado.columns if str(c).startswith('Sem')]

            if columna_datos_a_extraer != 'TPI_Global':
                if 'Tiempo Acumulado' in df_actualizado.columns:
                    df_actualizado['Tiempo Acumulado'] = df_actualizado[columnas_todas_semanas].sum(axis=1)

                if 'Promedio' in df_actualizado.columns:
                    df_actualizado['Promedio'] = df_actualizado[columnas_todas_semanas].mean(axis=1)
            else:
                # Para el CV también calculamos el promedio histórico redondeado a 2 decimales
                if 'Promedio' in df_actualizado.columns:
                    df_actualizado['Promedio'] = df_actualizado[columnas_todas_semanas].mean(axis=1).round(2)

            # --- CONVERSIÓN VISUAL FINAL (DE MINUTOS A FORMATO HH:MM) ---
            if columna_datos_a_extraer != 'TPI_Global':
                columnas_a_formatear = columnas_todas_semanas.copy()
                if 'Tiempo Acumulado' in df_actualizado.columns: columnas_a_formatear.append('Tiempo Acumulado')
                if 'Promedio' in df_actualizado.columns: columnas_a_formatear.append('Promedio')
                
                for col_fmt in columnas_a_formatear:
                    df_actualizado[col_fmt] = df_actualizado[col_fmt].apply(to_hhmm_display)

            # Guardamos la pestaña actualizada
            dict_dfs_actualizados[nombre_hoja_original] = df_actualizado.drop(columns=['MatchKey'], errors='ignore')

        else:
            dict_dfs_actualizados[nombre_hoja_original] = dict_dfs_originales[nombre_hoja_original]

    # --- INYECCIÓN DEL KPI DE ADHERENCIA EN EL EXCEL ---
    df_kpi_excel = df_semana_actual[['Deportista', 'TPI_Global', 'TPI_Natacion', 'TPI_Ciclismo', 'TPI_Trote', 'Es_Completo']].copy()
    
    for col_tpi in ['TPI_Global', 'TPI_Natacion', 'TPI_Ciclismo', 'TPI_Trote']:
        df_kpi_excel[col_tpi] = df_kpi_excel[col_tpi].apply(lambda x: f"{x:.1f}%" if pd.notna(x) else "0.0%")
        
    df_kpi_excel['Es_Completo'] = df_kpi_excel['Es_Completo'].apply(lambda x: "Sí" if x else "No")
    
    df_kpi_excel = df_kpi_excel.rename(columns={
        'Deportista': 'Nombre del Deportista',
        'TPI_Global': 'Adherencia Global',
        'TPI_Natacion': 'Adherencia Natación',
        'TPI_Ciclismo': 'Adherencia Ciclismo',
        'TPI_Trote': 'Adherencia Trote',
        'Es_Completo': f'Completó Plan ({etiqueta_semana})'
    })

    # --- NUEVA EXPANSIÓN: CREACIÓN DE LA HOJA CRUDA DE LA SEMANA ---
    df_hoja_semana = df_semana_actual[['Deportista', 'N_Mins_Real', 'B_Mins_Real', 'R_Mins_Real', 'T_Mins_Real']].copy()
    for col_tiempo in ['N_Mins_Real', 'B_Mins_Real', 'R_Mins_Real', 'T_Mins_Real']:
        df_hoja_semana[col_tiempo] = df_hoja_semana[col_tiempo].apply(to_hhmm_display)
        
    df_hoja_semana = df_hoja_semana.rename(columns={
        'Deportista': 'Nombre del Deportista',
        'N_Mins_Real': 'Natación',
        'B_Mins_Real': 'Bicicleta',
        'R_Mins_Real': 'Trote',
        'T_Mins_Real': 'Tiempo Total'
    })

    # --- ORDENAMIENTO ESTRICTO DE PESTAÑAS FINALES (JERARQUÍA Y DESCENDENTE) ---
    dict_final_ordenado = {}
    
    # Insertamos temporalmente las nuevas hojas para el escaneo
    dict_dfs_actualizados['KPI_Adherencia_TPI'] = df_kpi_excel
    dict_dfs_actualizados[etiqueta_semana] = df_hoja_semana
    
    # Jerarquía exacta solicitada
    orden_jerarquia = [
        'NUMEROS DE SEMANA', 'NUMERO DE SEMANA', 'SEMANAS', 'CALENDARIO',
        'TIEMPO TOTAL', 
        'NATACION', 'NATACIÓN', 
        'BICICLETA', 'CICLISMO', 
        'TROTE', 'RUNNING', 
        'CV', 
        'KPI_ADHERENCIA_TPI'
    ]
    
    claves_base_ordenadas = []
    claves_semanas = []
    claves_otras = []
    
    for clave_real in dict_dfs_actualizados.keys():
        clave_limpia = clean_string(clave_real)
        
        # Aislar las hojas históricas de las semanas (Sem 08, Sem 07...)
        if str(clave_real).strip().upper().startswith('SEM '):
            claves_semanas.append(clave_real)
        else:
            if clave_limpia in orden_jerarquia:
                claves_base_ordenadas.append((orden_jerarquia.index(clave_limpia), clave_real))
            else:
                claves_otras.append(clave_real)
                
    # Ordenar las bases por la jerarquía solicitada
    claves_base_ordenadas.sort(key=lambda x: x[0])
    
    # Ordenar las semanas cronológicamente de forma descendente
    claves_semanas.sort(reverse=True)
    
    # Ensamblar el Excel Final
    for _, clave_real in claves_base_ordenadas:
        dict_final_ordenado[clave_real] = dict_dfs_actualizados[clave_real]
        
    for clave_semana in claves_semanas:
        dict_final_ordenado[clave_semana] = dict_dfs_actualizados[clave_semana]
        
    for clave_otra in claves_otras:
        dict_final_ordenado[clave_otra] = dict_dfs_actualizados[clave_otra]

    return dict_final_ordenado


# =============================================================================
# FIN DE SECCIÓN 5
# =============================================================================

# *****************************************************************************
# SECCIÓN 6: ORQUESTADOR DE ENTREGABLES (GRÁFICOS, COLORES Y REPORTES)
# *****************************************************************************
# Esta versión mantiene la estructura vertical extendida. Incluye el candado
# estricto de columnas para forzar a Word a respetar los márgenes, el centrado
# absoluto de tablas y el filtro de elegibilidad para TPI.

import matplotlib.pyplot as plt
import numpy as np
import io
import zipfile
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

def ajustar_anchos_y_centrar_tabla(tabla, anchos):
    """
    Candado de formato: Desactiva el autoajuste, centra la tabla y fuerza
    el ancho desde la raíz de la columna para evitar desbordes en los márgenes.
    """
    tabla.autofit = False
    tabla.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # 1. Fijar el ancho de la columna desde la estructura principal
    for i, col in enumerate(tabla.columns):
        if i < len(anchos):
            col.width = anchos[i]
            
    # 2. Reforzar el ancho celda por celda para someter a Word
    for row in tabla.rows:
        for idx, width in enumerate(anchos):
            if idx < len(row.cells):
                row.cells[idx].width = width

def generar_velocimetro_tpi(porcentaje):
    """
    Dibuja un gráfico de velocímetro (Gauge Chart) semicircular para el TPI.
    """
    fig, ax = plt.subplots(figsize=(4, 2), subplot_kw={'projection': 'polar'})
    
    colors = ['#FF4C4C', '#FFD700', '#4CAF50']
    theta = np.linspace(0, np.pi, 100)
    
    ax.fill_between(np.linspace(np.pi, np.pi - (0.7 * np.pi), 50), 0.6, 1, color=colors[0], alpha=0.6)
    ax.fill_between(np.linspace(np.pi - (0.7 * np.pi), np.pi - (0.9 * np.pi), 50), 0.6, 1, color=colors[1], alpha=0.6)
    ax.fill_between(np.linspace(np.pi - (0.9 * np.pi), 0, 50), 0.6, 1, color=colors[2], alpha=0.6)
    
    p_grafico = min(porcentaje, 115)
    angulo_aguja = np.pi * (1 - (p_grafico / 115))
    
    ax.plot([angulo_aguja, angulo_aguja], [0, 0.9], color='black', linewidth=3, solid_capstyle='round')
    ax.plot(angulo_aguja, 0.9, marker='^', color='black', markersize=8)
    ax.plot(0, 0, marker='o', color='black', markersize=10)
    
    ax.set_ylim(0, 1)
    ax.set_yticks([])
    ax.set_xticks(np.pi * np.array([1, 0.39, 0.21, 0]))
    ax.set_xticklabels(['0%', '70%', '90%', '115%'], fontsize=10, weight='bold')
    ax.spines['polar'].set_visible(False)
    
    plt.text(0.5, 0.1, f"{porcentaje:.1f}%", transform=ax.transAxes, fontsize=20, weight='bold', ha='center', va='center')
             
    buf = io.BytesIO()
    plt.savefig(buf, format='png', transparent=True, bbox_inches='tight', dpi=150)
    plt.close(fig)
    buf.seek(0)
    return buf

def generar_grafico_distribucion(mins_natacion, mins_ciclismo, mins_trote):
    """
    Dibuja un gráfico de anillo para mostrar la distribución de disciplinas.
    Con textos en NEGRO y ajustados para legibilidad total en el Word.
    """
    fig, ax = plt.subplots(figsize=(4, 4))
    
    sizes = [mins_natacion, mins_ciclismo, mins_trote]
    labels = ['Natación', 'Ciclismo', 'Trote']
    colors = ['#1f77b4', '#ff7f0e', '#2ca02c']
    
    if sum(sizes) == 0:
        sizes = [1, 1, 1]
        labels = ['Sin Datos', 'Sin Datos', 'Sin Datos']
        
    wedges, texts, autotexts = ax.pie(
        sizes, 
        labels=labels, 
        autopct='%1.1f%%', 
        startangle=90, 
        colors=colors, 
        wedgeprops=dict(width=0.4, edgecolor='w'), 
        pctdistance=0.75,
        textprops={'color': 'black', 'weight': 'bold', 'size': 11}
    )
    
    buf = io.BytesIO()
    plt.savefig(buf, format='png', transparent=True, bbox_inches='tight', dpi=150)
    plt.close(fig)
    buf.seek(0)
    return buf

def generar_entregables_separados(df_semanal_procesado, dict_maestro_actualizado, etiqueta_semana):
    
    # =========================================================================
    # FASE 2: PRE-CÁLCULOS MATEMÁTICOS PARA ANÁLISIS COMPARATIVO
    # =========================================================================
    promedio_equipo = {
        'Total': df_semanal_procesado['T_Mins_Real'].mean(),
        'Natacion': df_semanal_procesado['N_Mins_Real'].mean(),
        'Ciclismo': df_semanal_procesado['B_Mins_Real'].mean(),
        'Trote': df_semanal_procesado['R_Mins_Real'].mean()
    }

    def obtener_media_historica(deportista_matchkey, hoja_maestro):
        df_hoja = dict_maestro_actualizado.get(hoja_maestro)
        if df_hoja is None:
            return 0
        if df_hoja.empty:
            return 0
        if 'Promedio' not in df_hoja.columns:
            return 0
            
        fila_atleta = df_hoja[df_hoja['MatchKey'] == deportista_matchkey]
        if not fila_atleta.empty:
            return to_mins(fila_atleta['Promedio'].values[0])
        return 0

    def redactar_comparacion(minutos_reales, minutos_equipo, minutos_historicos):
        diff_equipo = minutos_reales - minutos_equipo
        if diff_equipo >= 0:
            txt_equipo = "MÁS"
        else:
            txt_equipo = "MENOS"
            
        val_equipo = to_hhmm_display(abs(diff_equipo))
        
        diff_hist = minutos_reales - minutos_historicos
        if diff_hist >= 0:
            txt_hist = "MÁS"
        else:
            txt_hist = "MENOS"
            
        val_hist = to_hhmm_display(abs(diff_hist))
        
        texto_final = f"Rendiste {val_equipo} {txt_equipo} que el promedio del equipo. Vs Tu Media Histórica: {val_hist} {txt_hist}."
        return texto_final

    # =========================================================================
    # 6.1: GENERACIÓN DEL EXCEL MAESTRO ACTUALIZADO
    # =========================================================================
    buffer_excel = io.BytesIO()
    with pd.ExcelWriter(buffer_excel, engine='xlsxwriter') as writer:
        for nombre_hoja, df_hoja in dict_maestro_actualizado.items():
            df_hoja.to_excel(writer, sheet_name=nombre_hoja, index=False)
    buffer_excel.seek(0)
    
    # =========================================================================
    # 6.2: GENERACIÓN DEL REPORTE GRUPAL DEL CLUB (WORD AVANZADO)
    # =========================================================================
    doc_grupal = Document()
    
    # FORZAR CALIBRI EN TODO EL DOCUMENTO GRUPAL
    style_normal = doc_grupal.styles['Normal']
    font_normal = style_normal.font
    font_normal.name = 'Calibri'
    font_normal.size = Pt(11)

    num_semana = str(etiqueta_semana).replace('Sem ', '')
    
    # --- PÁGINA 1: PORTADA, INTRODUCCIÓN Y DATOS GENERALES ---
    doc_grupal.add_heading(f"Reporte Semanal Club Tym Triatlón \nSemana {num_semana}", 0)
    doc_grupal.add_paragraph("") 
    
    doc_grupal.add_heading("1. Introducción General", level=1)
    p_intro = doc_grupal.add_paragraph()
    run_intro = p_intro.add_run("[ESPACIO PARA EDICIÓN MANUAL: Inserta aquí tu introducción de 4 líneas]")
    run_intro.bold = True
    run_intro.font.color.rgb = RGBColor(255, 0, 0)
    doc_grupal.add_paragraph("") 
    
    doc_grupal.add_heading("2. Datos Generales de la Semana", level=1)
    df_activos = df_semanal_procesado[df_semanal_procesado['T_Mins_Real'] > 0]
    mins_n = df_activos['N_Mins_Real'].sum()
    mins_b = df_activos['B_Mins_Real'].sum()
    mins_r = df_activos['R_Mins_Real'].sum()
    mins_totales = df_activos['T_Mins_Real'].sum()
    
    p_res = doc_grupal.add_paragraph()
    p_res.add_run(f"Atletas activos esta semana: {len(df_activos)}\n").bold = True
    p_res.add_run(f"Horas de entrenamiento total del Equipo: {to_hhmm_display(mins_totales)}\n\n").bold = True
    p_res.add_run(f"Distribución por disciplina:\n")
    p_res.add_run(f"• Natación: {to_hhmm_display(mins_n)}\n")
    p_res.add_run(f"• Ciclismo: {to_hhmm_display(mins_b)}\n")
    p_res.add_run(f"• Trote: {to_hhmm_display(mins_r)}")
    doc_grupal.add_paragraph("") 
    
    doc_grupal.add_heading("Distribución Gráfica", level=2)
    img_dist = generar_grafico_distribucion(mins_n, mins_b, mins_r)
    para_img_dist = doc_grupal.add_paragraph()
    para_img_dist.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para_img_dist.add_run().add_picture(img_dist, width=Inches(3.5))
    doc_grupal.add_page_break() 

    # --- PÁGINA 2: TOP 5 COMPLETOS ---
    doc_grupal.add_heading("3. TOP 5 TRIATLETAS COMPLETOS", level=1)
    doc_grupal.add_paragraph("(Entrenamiento registrado en las 3 disciplinas)")
    
    tabla_completos = doc_grupal.add_table(rows=1, cols=6)
    tabla_completos.style = 'Light Grid Accent 1'
    
    # Ancho total = 5.8 pulgadas (Garantiza no salirse de la hoja)
    ajustar_anchos_y_centrar_tabla(tabla_completos, [Inches(0.4), Inches(2.2), Inches(0.8), Inches(0.8), Inches(0.8), Inches(0.8)])
    
    cc = tabla_completos.rows[0].cells
    cc[0].text = '#'
    cc[1].text = 'Deportista'
    cc[2].text = 'Total'
    cc[3].text = 'Nat.'
    cc[4].text = 'Bici'
    cc[5].text = 'Trote'
    
    df_completos = df_semanal_procesado[df_semanal_procesado['Es_Completo'] == True]
    df_top_completos = df_completos.sort_values(by='T_Mins_Real', ascending=False).head(5)
    
    for pos, (_, fila) in enumerate(df_top_completos.iterrows(), 1):
        rc = tabla_completos.add_row().cells
        rc[0].text = str(pos)
        rc[1].text = str(fila['Deportista'])
        rc[2].text = to_hhmm_display(fila['T_Mins_Real'])
        rc[3].text = to_hhmm_display(fila['N_Mins_Real'])
        rc[4].text = to_hhmm_display(fila['B_Mins_Real'])
        rc[5].text = to_hhmm_display(fila['R_Mins_Real'])
        
    doc_grupal.add_paragraph("") 
        
    doc_grupal.add_heading("Comentarios del Grupo de Élite (Completos):", level=2)
    for pos, (_, fila) in enumerate(df_top_completos.iterrows(), 1):
        comentario = generar_comentario(fila, 'CV', pos)
        doc_grupal.add_paragraph(f"🏅 {fila['Deportista']}: {comentario}")
        
    doc_grupal.add_page_break() 

    # --- PÁGINA 3: TOP 15 ADHERENCIA AL PLAN ---
    doc_grupal.add_heading("📈 4. TOP 15 ADHERENCIA AL PLAN (TPI GLOBAL)", level=1)
    doc_grupal.add_paragraph("") 
    
    # Filtro estricto para Adherencia
    def es_elegible_para_podio_tpi(row):
        # Si la meta de plan es > 0 pero el real es 0, queda fuera.
        if row.get('Natacion_Plan_Hrs', 0) > 0 and row.get('N_Mins_Real', 0) == 0:
            return False
        if row.get('Ciclismo_Plan_Hrs', 0) > 0 and row.get('B_Mins_Real', 0) == 0:
            return False
        if row.get('Trote_Plan_Hrs', 0) > 0 and row.get('R_Mins_Real', 0) == 0:
            return False
        # Debe haber registrado al menos alguna actividad en total
        if row.get('T_Mins_Real', 0) == 0:
            return False
        return True
        
    df_elegibles_tpi = df_semanal_procesado[df_semanal_procesado.apply(es_elegible_para_podio_tpi, axis=1)]
    df_ranking_tpi = df_elegibles_tpi.sort_values(by='TPI_Global', ascending=False).head(15)
    
    tabla_tpi = doc_grupal.add_table(rows=1, cols=3)
    tabla_tpi.style = 'Light Grid Accent 1'
    
    # Ancho total = 5.5 pulgadas (Centrado y dentro de los márgenes)
    ajustar_anchos_y_centrar_tabla(tabla_tpi, [Inches(0.5), Inches(3.5), Inches(1.5)])
    
    c_tpi = tabla_tpi.rows[0].cells
    c_tpi[0].text = 'Pos.'
    c_tpi[1].text = 'Deportista'
    c_tpi[2].text = 'TPI Global %'
    
    for pos, (_, fila) in enumerate(df_ranking_tpi.iterrows(), 1):
        rc = tabla_tpi.add_row().cells
        rc[0].text = str(pos)
        rc[1].text = str(fila['Deportista'])
        rc[2].text = f"{fila['TPI_Global']:.1f}%"
        
    doc_grupal.add_paragraph("") 
        
    doc_grupal.add_heading("Comentarios del Podio (Adherencia):", level=2)
    medallas = {1: "🥇", 2: "🥈", 3: "🥉"}
    for pos, (_, fila) in enumerate(df_ranking_tpi.head(3).iterrows(), 1):
        comentario = generar_comentario(fila, 'TPI', pos)
        doc_grupal.add_paragraph(f"{medallas[pos]} {fila['Deportista']}: {comentario}")
        
    doc_grupal.add_page_break() 

    # --- PÁGINAS 4 a 7: TOP 15 POR TIEMPO Y DISCIPLINAS ---
    bloques_tops = [
        ("⏱️ 5. TOP 15 TIEMPO TOTAL", 'T_Mins_Real', 'General'),
        ("🏊‍♂️ 6. TOP 15 NATACIÓN", 'N_Mins_Real', 'Natación'),
        ("🚴‍♂️ 7. TOP 15 CICLISMO", 'B_Mins_Real', 'Bicicleta'),
        ("🏃‍♂️ 8. TOP 15 TROTE", 'R_Mins_Real', 'Trote')
    ]
    
    for titulo, columna, categoria_frase in bloques_tops:
        doc_grupal.add_heading(titulo, level=1)
        doc_grupal.add_paragraph("") 
        
        tabla_disc = doc_grupal.add_table(rows=1, cols=3)
        tabla_disc.style = 'Light Grid Accent 1'
        
        # Ancho total = 5.5 pulgadas (Centrado y dentro de los márgenes)
        ajustar_anchos_y_centrar_tabla(tabla_disc, [Inches(0.5), Inches(3.5), Inches(1.5)])
        
        cd = tabla_disc.rows[0].cells
        cd[0].text = 'Pos.'
        cd[1].text = 'Deportista'
        cd[2].text = 'Tiempo'
        
        df_disc = df_semanal_procesado[df_semanal_procesado[columna] > 0].sort_values(by=columna, ascending=False).head(15)
        for pos, (_, fila) in enumerate(df_disc.iterrows(), 1):
            rc = tabla_disc.add_row().cells
            rc[0].text = str(pos)
            rc[1].text = str(fila['Deportista'])
            rc[2].text = to_hhmm_display(fila[columna])
            
        doc_grupal.add_paragraph("") 
            
        doc_grupal.add_heading(f"Comentarios del Podio ({categoria_frase}):", level=2)
        for pos, (_, fila) in enumerate(df_disc.head(3).iterrows(), 1):
            comentario = generar_comentario(fila, categoria_frase, pos)
            doc_grupal.add_paragraph(f"{medallas[pos]} {fila['Deportista']}: {comentario}")
        
        doc_grupal.add_page_break()

    # --- PÁGINA 8: ESTRATEGIA Y CONCLUSIONES ---
    doc_grupal.add_heading("💡 9. Insights Estratégicos", level=1)
    doc_grupal.add_paragraph("") 
    
    p_ins = doc_grupal.add_paragraph()
    run_ins_1 = p_ins.add_run("Bici-dependencia Crónica: ")
    run_ins_1.bold = True
    p_ins.add_run("El ciclismo sigue siendo el \"hijo favorito\". Representa la mayor parte del tiempo total del club. Básicamente, si le quitamos las bicicletas a este equipo, el reporte semanal cabría en una servilleta.\n")
    
    run_ins_2 = p_ins.add_run("El \"Efecto Claudio\": ")
    run_ins_2.bold = True
    p_ins.add_run("Una lección magistral de cómo liderar el ciclismo y aun así bajar en el ranking general por el \"pequeño\" detalle de no tocar el agua. Recordatorio amistoso: el Triatlón, por definición, incluye nadar.\n")
    
    run_ins_3 = p_ins.add_run("Resiliencia en el Asfalto: ")
    run_ins_3.bold = True
    p_ins.add_run("Mientras el volumen general bajaba, el trote sacó la cara con registros notables de carrera a pie. Parece que algunos sí desayunaron ganas de correr.\n")
    
    doc_grupal.add_paragraph("") 
    
    doc_grupal.add_heading("10. Conclusiones Generales", level=1)
    p_cierre = doc_grupal.add_paragraph()
    run_cierre = p_cierre.add_run("[ESPACIO PARA EDICIÓN MANUAL: Inserta aquí tus conclusiones finales y el cierre de la semana]")
    run_cierre.bold = True
    run_cierre.font.color.rgb = RGBColor(255, 0, 0)

    buffer_word_grupal = io.BytesIO()
    doc_grupal.save(buffer_word_grupal)
    buffer_word_grupal.seek(0)

    # =========================================================================
    # 6.3: GENERACIÓN DE FICHAS INDIVIDUALES (ZIP) 
    # =========================================================================
    buffer_zip_fichas = io.BytesIO()
    
    with zipfile.ZipFile(buffer_zip_fichas, "a", zipfile.ZIP_DEFLATED) as archivo_zip:
        
        for index, row_atleta in df_semanal_procesado.iterrows():
            if row_atleta['T_Mins_Real'] > 0:
                doc_i = Document()
                
                style_ind = doc_i.styles['Normal']
                font_ind = style_ind.font
                font_ind.name = 'Calibri'
                font_ind.size = Pt(11)
                
                doc_i.add_heading(f"Análisis de Rendimiento Personal: {row_atleta['Deportista']}", 0)
                
                p_sub = doc_i.add_paragraph(f"Semana de Entrenamiento: {num_semana}")
                p_sub.bold = True
                doc_i.add_paragraph("") 
                
                # --- BLOQUE 1: VELOCÍMETRO DE ADHERENCIA (TPI) ---
                doc_i.add_heading("CUMPLIMIENTO DE PLAN (TPI)", level=2)
                p_tpi_intro = doc_i.add_paragraph("Tu Índice de Adherencia Global esta semana:")
                p_tpi_intro.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                img_velocimetro = generar_velocimetro_tpi(row_atleta['TPI_Global'])
                para_img = doc_i.add_paragraph()
                para_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                para_img.add_run().add_picture(img_velocimetro, width=Inches(3.5))
                doc_i.add_paragraph("") 
                
                tabla_desglose = doc_i.add_table(rows=1, cols=4)
                tabla_desglose.style = 'Table Grid'
                ajustar_anchos_y_centrar_tabla(tabla_desglose, [Inches(1.5), Inches(1.2), Inches(1.2), Inches(1.2)])
                
                c_cab = tabla_desglose.rows[0].cells
                c_cab[0].text = 'Disciplina'
                c_cab[1].text = 'Real'
                c_cab[2].text = 'Meta'
                c_cab[3].text = 'TPI %'
                
                matriz = [
                    ('Natación', 'N_Mins_Real', 'Natacion_Plan_Hrs', 'TPI_Natacion'),
                    ('Ciclismo', 'B_Mins_Real', 'Ciclismo_Plan_Hrs', 'TPI_Ciclismo'),
                    ('Trote', 'R_Mins_Real', 'Trote_Plan_Hrs', 'TPI_Trote')
                ]
                
                for d, c_r, c_m, c_tpi in matriz:
                    rc = tabla_desglose.add_row().cells
                    rc[0].text = d
                    rc[1].text = to_hhmm_display(row_atleta[c_r])
                    rc[2].text = f"{row_atleta[c_m]:.1f}h"
                    
                    run_tpi = rc[3].paragraphs[0].add_run(f"{row_atleta[c_tpi]:.1f}%")
                    run_tpi.bold = True
                    
                    tpi_val = row_atleta[c_tpi]
                    if tpi_val < 70:
                        run_tpi.font.color.rgb = RGBColor(255, 0, 0)
                    elif tpi_val <= 90:
                        run_tpi.font.color.rgb = RGBColor(204, 153, 0)
                    else:
                        run_tpi.font.color.rgb = RGBColor(0, 153, 0)
                        
                doc_i.add_paragraph("") 

                # --- BLOQUE 2: ANÁLISIS HISTÓRICO COMPARATIVO ---
                mk = row_atleta['MatchKey']
                bloques_analisis = [
                    ('TIEMPO TOTAL', 'T_Mins_Real', 'Total', 'TIEMPO TOTAL'),
                    ('NATACIÓN', 'N_Mins_Real', 'Natacion', 'NATACION'),
                    ('CICLISMO', 'B_Mins_Real', 'Ciclismo', 'BICICLETA'),
                    ('TROTE', 'R_Mins_Real', 'Trote', 'TROTE')
                ]
                
                for titulo, col_real, llave_equipo, hoja_historia in bloques_analisis:
                    doc_i.add_heading(titulo, level=2)
                    p_dato = doc_i.add_paragraph()
                    
                    run_volumen = p_dato.add_run(f"Volumen actual: {to_hhmm_display(row_atleta[col_real])} ")
                    run_volumen.bold = True
                    
                    hist_mins = obtener_media_historica(mk, hoja_historia)
                    texto_comp = redactar_comparacion(row_atleta[col_real], promedio_equipo[llave_equipo], hist_mins)
                    p_dato.add_run(texto_comp)
                    
                doc_i.add_paragraph("") 

                # --- BLOQUE 3: EVALUACIÓN TÉCNICA (MOTOR NARRATIVO) ---
                doc_i.add_heading("Evaluación Técnica", level=2)
                
                # El ranking para la narrativa siempre usará la posición del TPI General
                pos = df_ranking_tpi.index[df_ranking_tpi['Deportista'] == row_atleta['Deportista']].tolist()
                
                if pos:
                    rango_actual = pos[0] + 1
                else:
                    rango_actual = 99
                    
                texto_comentario = generar_comentario(row_atleta, 'General', rango_actual)
                doc_i.add_paragraph(texto_comentario)
                
                doc_i.add_paragraph("\n──────────────────────────────────────────────────")
                doc_i.add_paragraph("Generado por Plataforma TYM Performance")
                
                buffer_word_individual = io.BytesIO()
                doc_i.save(buffer_word_individual)
                
                nombre_archivo = f"Reporte_{clean_string(row_atleta['Deportista'])}.docx"
                archivo_zip.writestr(nombre_archivo, buffer_word_individual.getvalue())

    buffer_zip_fichas.seek(0)
    
    return {
        'excel_maestro': buffer_excel,
        'word_grupal': buffer_word_grupal,
        'zip_fichas': buffer_zip_fichas
    }
    
# *****************************************************************************
# SECCIÓN 7: INTERFAZ DE USUARIO, ORQUESTACIÓN Y CONSOLA DE DEBUG
# *****************************************************************************
# Esta sección construye el panel de control, ejecuta la cadena de motores
# (Secciones 2 a 6) de forma secuencial y despliega los tres botones de 
# descarga independientes generados por el orquestador.

# -----------------------------------------------------------------------------
# 7.1: CONFIGURACIÓN VISUAL DEL PANEL CENTRAL Y SIDEBAR
# -----------------------------------------------------------------------------
st.title("🏆 Plataforma de Rendimiento TYM v3.0")
st.markdown("Generador de KPIs de adherencia (TPI), Motor Narrativo e Integridad de Excel.")

with st.sidebar:
    st.header("⚙️ 1. Entradas de Sistema")
    st.markdown("Carga los archivos obligatorios para iniciar el procesamiento.")
    
    archivo_maestro = st.file_uploader("A. Maestro Histórico (Excel)", type=["xlsx", "xls"])
    archivo_strava = st.file_uploader("B. Strava Semanal (Excel)", type=["xlsx", "xls"])
    archivo_plan = st.file_uploader("C. Plan Individual (Opcional)", type=["xlsx", "xls"])
    
    st.divider()
    
    st.header("🎯 2. Metas Globales (Fallback)")
    st.markdown("Aplicables si un atleta no tiene Plan Individual.")
    metas_globales_sidebar = {
        'N_H': st.number_input("Natación - Horas", min_value=0.0, value=3.0, step=0.5),
        'N_S': st.number_input("Natación - Sesiones", min_value=1, value=3, step=1),
        'B_H': st.number_input("Ciclismo - Horas", min_value=0.0, value=5.0, step=0.5),
        'B_S': st.number_input("Ciclismo - Sesiones", min_value=1, value=3, step=1),
        'T_H': st.number_input("Trote - Horas", min_value=0.0, value=3.0, step=0.5),
        'T_S': st.number_input("Trote - Sesiones", min_value=1, value=3, step=1)
    }
    
    st.divider()
    
    st.header("🏷️ 3. Etiqueta Temporal")
    etiqueta_semana_input = st.text_input("Etiqueta a generar en el Maestro", "Sem 08")

# -----------------------------------------------------------------------------
# 7.2: INICIALIZACIÓN DEL ESTADO DE SESIÓN (SESSION STATE)
# -----------------------------------------------------------------------------
# Esto es vital para que los tres botones de descarga no desaparezcan 
# de la pantalla al hacer clic en uno de ellos.
if 'diccionario_entregables' not in st.session_state:
    st.session_state['diccionario_entregables'] = None

# -----------------------------------------------------------------------------
# 7.3: EJECUCIÓN PRINCIPAL (BOTÓN DE PROCESAMIENTO)
# -----------------------------------------------------------------------------
if st.button("🚀 PROCESAR EXCEL Y GENERAR ENTREGABLES", use_container_width=True):
    
    # Verificación de seguridad: No avanzar sin los archivos vitales
    if archivo_maestro is not None and archivo_strava is not None:
        
        with st.spinner("Purificando datos, calculando Adherencia TYM y generando gráficas..."):
            try:
                # --- FASE 1: EXTRACCIÓN (Sección 2) ---
                df_datos_reales = procesar_strava_excel(archivo_strava)
                df_datos_plan = procesar_plan_individual(archivo_plan)
                
                # --- FASE 2: CÁLCULO DE KPI (Sección 4) ---
                df_calculado_kpi = calcular_kpis_tym(df_datos_reales, df_datos_plan, metas_globales_sidebar)
                
                # --- FASE 3: CONSOLA DE AUDITORÍA (DEBUG VISUAL) ---
                # Mostrar en pantalla los resultados antes de empaquetarlos para validar precisión
                with st.expander("🕵️‍♂️ CONSOLA DE AUDITORÍA (VERIFICAR ANTES DE DESCARGAR)", expanded=True):
                    st.markdown("**1. Lectura Pura de Strava (Mins Reales Extraídos):**")
                    st.dataframe(df_calculado_kpi[['Deportista', 'N_Mins_Real', 'B_Mins_Real', 'R_Mins_Real', 'T_Mins_Real']].head(10))
                    
                    st.markdown("**2. Cálculo de KPIs (Adherencia):**")
                    st.dataframe(df_calculado_kpi[['Deportista', 'TPI_Natacion', 'TPI_Ciclismo', 'TPI_Trote', 'TPI_Global', 'Es_Completo']].head(10))

                # --- FASE 4: PERSISTENCIA DEL MAESTRO (Sección 5) ---
                diccionario_maestro_original = pd.read_excel(archivo_maestro, sheet_name=None)
                diccionario_maestro_actualizado = actualizar_maestro_tym(
                    diccionario_maestro_original, 
                    df_calculado_kpi, 
                    etiqueta_semana_input
                )
                
                # --- FASE 5: GENERACIÓN DE ENTREGABLES SEPARADOS (Sección 6) ---
                st.session_state['diccionario_entregables'] = generar_entregables_separados(
                    df_calculado_kpi, 
                    diccionario_maestro_actualizado, 
                    etiqueta_semana_input
                )
                
                st.success("✅ ¡Procesamiento completado con éxito! Gráficos generados. Revisa la consola y descarga tus archivos.")
                
            except Exception as error_critico:
                st.error(f"❌ Error crítico durante el procesamiento: {str(error_critico)}")
    else:
        st.warning("⚠️ Debes cargar el Maestro Histórico y el Excel de Strava en el panel lateral para iniciar.")

# -----------------------------------------------------------------------------
# 7.4: DESPLIEGUE DE LOS TRES BOTONES DE DESCARGA INDEPENDIENTES
# -----------------------------------------------------------------------------
if st.session_state['diccionario_entregables'] is not None:
    st.markdown("### 📥 Descarga de Archivos Procesados")
    st.markdown("Selecciona el entregable que deseas descargar:")
    
    columna_excel, columna_word, columna_zip = st.columns(3)
    
    with columna_excel:
        st.download_button(
            label="📊 1. Descargar Maestro Actualizado (.xlsx)",
            data=st.session_state['diccionario_entregables']['excel_maestro'],
            file_name=f"01_Estadisticas_TYM_{etiqueta_semana_input}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
    with columna_word:
        st.download_button(
            label="📄 2. Descargar Reporte Grupal (.docx)",
            data=st.session_state['diccionario_entregables']['word_grupal'],
            file_name=f"02_Reporte_General_{etiqueta_semana_input}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
        
    with columna_zip:
        st.download_button(
            label="🗂️ 3. Descargar Fichas Individuales (.zip)",
            data=st.session_state['diccionario_entregables']['zip_fichas'],
            file_name=f"03_Fichas_Individuales_{etiqueta_semana_input}.zip",
            mime="application/zip",
            use_container_width=True
        )
