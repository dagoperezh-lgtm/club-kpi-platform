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
        # Creamos una copia de la lista para no alterar la original y la barajamos
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
    
    # Dependiendo de la categoría, extraemos el tiempo formateado en HH:MM
    # Asumimos que datos_de_fila trae los minutos reales procesados en la Sección 2
    if nombre_categoria == 'Natación':
        tiempo = to_hhmm_display(datos_de_fila.get('N_Mins_Real', 0))
    elif nombre_categoria == 'Bicicleta':
        tiempo = to_hhmm_display(datos_de_fila.get('B_Mins_Real', 0))
    elif nombre_categoria == 'Trote':
        tiempo = to_hhmm_display(datos_de_fila.get('R_Mins_Real', 0))
    else:
        tiempo = to_hhmm_display(datos_de_fila.get('T_Mins_Real', 0))
    
    # BANCO DE FRASES EXTENDIDO (VERSIÓN ABSOLUTA SIN OMISIONES)
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

    # Determinamos la llave de la categoría (Fallback a 'General' si el nombre no cuadra exacto)
    cat_key = 'General' if nombre_categoria in ['Completos', 'General'] else nombre_categoria
    
    # Extraemos la plantilla aleatoria
    frase_plantilla = str(obtener_frase_base(cat_key, pools.get(cat_key, pools['General'])))
    
    # Inyectamos los datos del deportista
    comentario_final = frase_plantilla.replace("{atleta}", atleta).replace("{tiempo}", tiempo)
    
    # Bonus de liderazgo: Si es el número 1 del ranking general, añadimos un reconocimiento extra
    if rank_posicion == 1 and cat_key == 'General':
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
# SECCIÓN 5: MOTOR DE PERSISTENCIA Y ACTUALIZACIÓN DEL MAESTRO
# *****************************************************************************
# Esta sección actualiza el libro Excel Histórico preservando las semanas
# anteriores. Implementa un blindaje contra columnas duplicadas (sufijos _x, _y)
# y aplica la "vacuna" para purificar los tipos de datos antiguos SIN destruirlos.

def actualizar_maestro_tym(dict_dfs_originales, df_semana_actual, etiqueta_semana):
    """
    Actualiza hoja por hoja el archivo Maestro.
    Garantiza que la información se indexe correctamente a cada atleta
    y que la historia no se borre ni se convierta en ceros.
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

            # --- BLINDAJE ANTI-DUPLICADOS (Evita Sem 08_x) ---
            # Si el usuario procesa la misma semana dos veces, borramos la columna antes de escribir
            if etiqueta_semana in df_hoja_historia.columns:
                df_hoja_historia = df_hoja_historia.drop(columns=[etiqueta_semana])

            # --- SALVACIÓN Y PURIFICACIÓN DEL HISTORIAL (La Vacuna anti-ceros) ---
            # Identificamos todas las columnas históricas que empiezan con 'Sem'
            columnas_semanas_viejas = [c for c in df_hoja_historia.columns if str(c).startswith('Sem')]

            for col_vieja in columnas_semanas_viejas:
                if columna_datos_a_extraer != 'TPI_Global':
                    # Si es tiempo, usamos nuestra función to_mins de la Sección 1 y dividimos por 1440
                    # Esto garantiza que "02:30" o Timedeltas no se borren, sino que se conviertan seguro.
                    df_hoja_historia[col_vieja] = df_hoja_historia[col_vieja].apply(to_mins) / 1440.0
                else:
                    # Si es la hoja CV (TPI_Global), son porcentajes, forzamos numérico sin to_mins
                    df_hoja_historia[col_vieja] = pd.to_numeric(df_hoja_historia[col_vieja], errors='coerce').fillna(0.0)

            # --- PREPARACIÓN DE LA NOVEDAD (DATOS DE LA SEMANA) ---
            df_novedad = df_semana_actual[['MatchKey', columna_datos_a_extraer]].copy()

            if columna_datos_a_extraer != 'TPI_Global':
                df_novedad[etiqueta_semana] = df_novedad[columna_datos_a_extraer].apply(
                    lambda x: (x / 1440.0) if pd.notna(x) else 0.0
                )
            else:
                df_novedad[etiqueta_semana] = df_novedad[columna_datos_a_extraer] / 100.0

            # Eliminamos duplicados en la novedad por seguridad extrema
            df_novedad = df_novedad.drop_duplicates(subset=['MatchKey'], keep='first')

            # --- UNIÓN DEL HISTÓRICO CON LA NOVEDAD ---
            # Usamos OUTER JOIN. Si hay un atleta nuevo en Strava que no estaba en el Maestro,
            # esto garantiza que se agregue al final de la lista y no se pierda.
            df_actualizado = pd.merge(
                df_hoja_historia,
                df_novedad[['MatchKey', etiqueta_semana]],
                on='MatchKey',
                how='outer'
            )

            # Rellenamos los vacíos generados por cruces de atletas antiguos/nuevos
            df_actualizado[etiqueta_semana] = df_actualizado[etiqueta_semana].fillna(0.0)

            # --- GESTIÓN DE ATLETAS NUEVOS ---
            # Rellenar el nombre del deportista si acaba de aparecer en el club esta semana
            mask_nuevos = df_actualizado[col_identidad_maestro].isna()
            df_actualizado.loc[mask_nuevos, col_identidad_maestro] = df_actualizado.loc[mask_nuevos, 'MatchKey'].map(mapeo_nombres_nuevos)

            # Rellenar las semanas históricas con 0 para los atletas nuevos
            for col_vieja in columnas_semanas_viejas:
                df_actualizado[col_vieja] = df_actualizado[col_vieja].fillna(0.0)

            # --- RECALCULO DE MÉTRICAS (TIEMPO ACUMULADO) ---
            columnas_todas_semanas = [c for c in df_actualizado.columns if str(c).startswith('Sem')]

            if columna_datos_a_extraer != 'TPI_Global':
                if 'Tiempo Acumulado' in df_actualizado.columns:
                    # Esta suma ya no explotará porque todas las celdas fueron forzadas a números puros
                    df_actualizado['Tiempo Acumulado'] = df_actualizado[columnas_todas_semanas].sum(axis=1)

                if 'Promedio' in df_actualizado.columns:
                    df_actualizado['Promedio'] = df_actualizado[columnas_todas_semanas].mean(axis=1)

            # Guardamos la pestaña actualizada eliminando la llave técnica
            dict_dfs_actualizados[nombre_hoja_original] = df_actualizado.drop(columns=['MatchKey'], errors='ignore')

        else:
            # Si la hoja no es de disciplinas (ej: Calendario), la traspasamos intacta
            dict_dfs_actualizados[nombre_hoja_original] = dict_dfs_originales[nombre_hoja_original]

    return dict_dfs_actualizados


# =============================================================================
# FIN DE SECCIÓN 5
# =============================================================================

# *****************************************************************************
# SECCIÓN 6: ORQUESTADOR DE ENTREGABLES (DESCARGAS SEPARADAS)
# *****************************************************************************
# Esta sección toma los datos calculados y purificados para generar los archivos
# finales. Cumpliendo con los requerimientos operativos, genera tres buffers
# separados en memoria para permitir descargas independientes en la interfaz.

def generar_entregables_separados(df_semanal_procesado, dict_maestro_actualizado, etiqueta_semana):
    """
    Construye en memoria (RAM) los tres archivos solicitados por separado:
    1. buffer_excel: El Excel Maestro Histórico actualizado.
    2. buffer_word_grupal: El Reporte General del Club.
    3. buffer_zip_fichas: Un archivo ZIP que contiene únicamente las fichas individuales.
    
    Devuelve un diccionario con los tres objetos para que la interfaz los procese.
    """
    
    # ---------------------------------------------------------------------
    # 6.1: GENERACIÓN DEL EXCEL MAESTRO ACTUALIZADO
    # ---------------------------------------------------------------------
    buffer_excel = io.BytesIO()
    # Usamos xlsxwriter para mayor estabilidad y soporte multifichas
    with pd.ExcelWriter(buffer_excel, engine='xlsxwriter') as writer:
        for nombre_hoja, df_hoja in dict_maestro_actualizado.items():
            df_hoja.to_excel(writer, sheet_name=nombre_hoja, index=False)
            
    # Reiniciamos el puntero del buffer para que Streamlit pueda leerlo desde el principio
    buffer_excel.seek(0)
    
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
    # Usamos el nuevo formato estricto de horas y minutos (HH:MM)
    parrafo_resumen.add_run(f"{to_hhmm_display(minutos_totales_club)}")

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

    # Guardar el Reporte Grupal en su propio buffer independiente
    buffer_word_grupal = io.BytesIO()
    doc_grupal.save(buffer_word_grupal)
    buffer_word_grupal.seek(0)

    # ---------------------------------------------------------------------
    # 6.3: GENERACIÓN DE FICHAS INDIVIDUALES (EMPAQUETADAS EN UN ZIP)
    # ---------------------------------------------------------------------
    buffer_zip_fichas = io.BytesIO()
    
    with zipfile.ZipFile(buffer_zip_fichas, "a", zipfile.ZIP_DEFLATED) as archivo_zip:
        
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
                    # Formateo visual del tiempo real a HH:MM (Sin segundos)
                    celdas_datos[1].text = to_hhmm_display(row_atleta[col_real])
                    # Formateo visual de la meta en horas
                    celdas_datos[2].text = f"{row_atleta[col_meta]:.1f}h"
                    # Formateo visual del TPI individual
                    celdas_datos[3].text = f"{row_atleta[col_tpi]:.1f}%"

                # --- Inyección del Motor Narrativo (Sección 3) ---
                doc_individual.add_heading("Evaluación Técnica", level=1)
                
                # Determinamos el ranking de este atleta para saber si es el líder
                posicion_ranking = df_ranking_top15.index[df_ranking_top15['Deportista'] == row_atleta['Deportista']].tolist()
                rango_actual = posicion_ranking[0] + 1 if posicion_ranking else 99
                
                # Generamos e inyectamos la frase experta
                comentario_experto = generar_comentario(row_atleta, 'General', rango_actual)
                doc_individual.add_paragraph(comentario_experto)
                
                # Guardamos la ficha en un buffer temporal
                buffer_word_individual = io.BytesIO()
                doc_individual.save(buffer_word_individual)
                
                # Usamos clean_string para asegurar que el nombre del archivo sea seguro para Windows/Mac
                nombre_archivo_seguro = f"Ficha_{clean_string(row_atleta['Deportista'])}.docx"
                
                # Insertamos el Word directamente en la raíz del archivo ZIP
                archivo_zip.writestr(nombre_archivo_seguro, buffer_word_individual.getvalue())

    # Reiniciamos el puntero del buffer ZIP
    buffer_zip_fichas.seek(0)
    
    # ---------------------------------------------------------------------
    # 6.4: DEVOLUCIÓN DEL DICCIONARIO DE ENTREGABLES
    # ---------------------------------------------------------------------
    # Devolvemos los tres canales por separado para que la interfaz genere 3 botones
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
st.markdown("Generador de KPIs de adherencia (TPI) y Motor de Integridad de Excel.")

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
        
        with st.spinner("Purificando datos y calculando Adherencia TYM..."):
            try:
                # --- FASE 1: EXTRACCIÓN (Sección 2) ---
                df_datos_reales = procesar_strava_excel(archivo_strava)
                df_datos_plan = procesar_plan_individual(archivo_plan)
                
                # --- FASE 2: CÁLCULO DE KPI (Sección 4) ---
                df_calculado_kpi = calcular_kpis_tym(df_datos_reales, df_datos_plan, metas_globales_sidebar)
                
                # --- FASE 3: CONSOLA DE AUDITORÍA (DEBUG VISUAL) ---
                # Mostrar en pantalla los resultados antes de empaquetarlos para validar que no haya ceros
                with st.expander("🕵️‍♂️ CONSOLA DE AUDITORÍA (VERIFICAR ANTES DE DESCARGAR)", expanded=True):
                    st.markdown("**1. Lectura Pura de Strava (Mins Reales Extraídos):**")
                    # Mostramos los minutos procesados para confirmar que la función to_mins hizo su trabajo
                    st.dataframe(df_calculado_kpi[['Deportista', 'N_Mins_Real', 'B_Mins_Real', 'R_Mins_Real', 'T_Mins_Real']].head(10))
                    
                    st.markdown("**2. Cálculo de KPIs (Adherencia):**")
                    # Mostramos los porcentajes resultantes y la bandera de Triatleta Completo
                    st.dataframe(df_calculado_kpi[['Deportista', 'TPI_Natacion', 'TPI_Ciclismo', 'TPI_Trote', 'TPI_Global', 'Es_Completo']].head(10))

                # --- FASE 4: PERSISTENCIA DEL MAESTRO (Sección 5) ---
                diccionario_maestro_original = pd.read_excel(archivo_maestro, sheet_name=None)
                diccionario_maestro_actualizado = actualizar_maestro_tym(
                    diccionario_maestro_original, 
                    df_calculado_kpi, 
                    etiqueta_semana_input
                )
                
                # --- FASE 5: GENERACIÓN DE ENTREGABLES SEPARADOS (Sección 6) ---
                # Guardamos el diccionario con los tres buffers en la memoria de la sesión
                st.session_state['diccionario_entregables'] = generar_entregables_separados(
                    df_calculado_kpi, 
                    diccionario_maestro_actualizado, 
                    etiqueta_semana_input
                )
                
                st.success("✅ ¡Procesamiento completado con éxito! Revisa la consola de auditoría y descarga tus archivos.")
                
            except Exception as error_critico:
                st.error(f"❌ Error de Integridad durante el procesamiento: {str(error_critico)}")
    else:
        st.warning("⚠️ Debes cargar el Maestro Histórico y el Excel de Strava en el panel lateral para iniciar.")

# -----------------------------------------------------------------------------
# 7.4: DESPLIEGUE DE LOS TRES BOTONES DE DESCARGA INDEPENDIENTES
# -----------------------------------------------------------------------------
if st.session_state['diccionario_entregables'] is not None:
    st.markdown("### 📥 Descarga de Archivos Procesados")
    st.markdown("Selecciona el entregable que deseas descargar:")
    
    # Creamos tres columnas para organizar los botones visualmente
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
