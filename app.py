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

# *****************************************************************************
# --- 1. CONFIGURACIÓN DE PÁGINA (BLINDADO - NO TOCAR) ---
# *****************************************************************************

st.set_page_config(
    page_title="Plataforma TYM 2026 - V2.2.19", 
    page_icon="🏆", 
    layout="wide"
)

st.title("🏆 Gestión de Reportes y Estadísticas - Club TYM")

# *****************************************************************************
# --- 2. UTILIDADES DE PROCESAMIENTO Y TIEMPO (BLINDADO - NO SINTETIZAR) ---
# *****************************************************************************

def clean_string(text):
    """
    Normaliza nombres para asegurar coincidencias entre Strava y Excel.
    Elimina tildes, espacios extra y convierte a mayúsculas.
    """
    if text is None or pd.isna(text):
        return ""
    
    # Proceso de normalización de caracteres paso a paso
    nombre_limpio_temp = str(text).strip()
    
    nombre_limpio_temp = nombre_limpio_temp.upper()
    
    # Uso de NFKD para descomponer caracteres con tildes
    info_normalizada = unicodedata.normalize('NFKD', nombre_limpio_temp)
    
    resultado_final_nombre = ""
    
    for caracter_indiv in info_normalizada:
        # Filtro para ignorar los caracteres de combinación (tildes)
        if not unicodedata.combining(caracter_indiv):
            resultado_final_nombre = resultado_final_nombre + caracter_indiv
            
    return resultado_final_nombre

def to_mins(valor_entrada_tiempo):
    """
    Convierte cualquier formato de tiempo a minutos totales de forma explícita.
    Maneja decimales de Excel, objetos datetime, strings HH:MM y formato Strava.
    """
    if pd.isna(valor_entrada_tiempo):
        return 0
    
    string_valor = str(valor_entrada_tiempo).strip()
    
    # Listado exhaustivo de casos nulos detectados en la operativa real
    lista_casos_nulos = ['--:--', '0', '', '00:00:00', '0:00:00', '00:00', '0.0', 'NC', '0:00']
    
    if string_valor in lista_casos_nulos:
        return 0
        
    try:
        # 🛡️ REGLA ARITMÉTICA: Si el valor es numérico (fracción de día de Excel)
        if isinstance(valor_entrada_tiempo, (float, int)):
            # Excel almacena 1 día completo como 1.0. 
            # Multiplicamos por 1440 para obtener la cifra real de minutos.
            minutos_finales_calculados = int(round(valor_entrada_tiempo * 1440))
            return minutos_finales_calculados
        
        # Si el dato es un objeto de tiempo nativo de Python
        if isinstance(valor_entrada_tiempo, (time, datetime)):
            minutos_finales_calculados = (valor_entrada_tiempo.hour * 60) + valor_entrada_tiempo.minute
            return minutos_finales_calculados
            
        # Si el string representa un número decimal puro
        try:
            conversion_float = float(string_valor)
            minutos_finales_calculados = int(round(conversion_float * 1440))
            return minutos_finales_calculados
        except ValueError:
            # No es numérico, continuamos con la lógica de parsing de texto
            pass
            
        # Formato de hora estándar con separador de dos puntos (HH:MM)
        if ':' in string_valor:
            bloques_tiempo = string_valor.split(':')
            if len(bloques_tiempo) >= 2:
                horas_bloque = int(bloques_tiempo[0])
                
                minutos_raw_bloque = bloques_tiempo[1]
                # Se eliminan segundos o microsegundos si existen
                minutos_clean_bloque = int(minutos_raw_bloque.split('.')[0])
                
                total_minutos_bloque = (horas_bloque * 60) + minutos_clean_bloque
                return total_minutos_bloque
        
        # Formato nativo de Strava (ejemplo: 11h 6min)
        busqueda_horas = re.search(r'(\d+)h', string_valor)
        busqueda_minutos = re.search(r'(\d+)min', string_valor)
        
        h_resultado = 0
        if busqueda_horas:
            h_resultado = int(busqueda_horas.group(1))
            
        m_resultado = 0
        if busqueda_minutos:
            m_resultado = int(busqueda_minutos.group(1))
            
        resultado_total_minutos = (h_resultado * 60) + m_resultado
        return resultado_total_minutos
        
    except Exception:
        # Fallback de seguridad para evitar que la aplicación se detenga
        return 0

def to_excel_time_value(dato_entrada_original):
    """
    Transforma la entrada en la fracción decimal exacta que requiere el motor de Excel.
    Este paso es vital para que las celdas sean sumables y promediables.
    """
    minutos_para_excel = to_mins(dato_entrada_original)
    
    # 24 horas equivalen a 1440 minutos totales
    valor_decimal_excel = minutos_para_excel / 1440.0
    
    return valor_decimal_excel

def to_hhmmss_display(minutos_totales_input):
    """
    Formato de texto HH:MM:00 exclusivo para la estética del reporte Word.
    """
    valor_horas_v = int(minutos_totales_input // 60)
    valor_minutos_v = int(minutos_totales_input % 60)
    
    # Generación de la cadena de texto con formato de reloj
    string_formato_reloj = f"{valor_horas_v:02d}:{valor_minutos_v:02d}:00"
    
    return string_formato_reloj

# =============================================================================
# SECCIÓN 3: MOTOR NARRATIVO PRO CHILE (V2.2.28 - 20+ FRASES POR SECCIÓN)
# =============================================================================
import random

# Diccionario global para persistencia de frases durante la ejecución
PILAS_COMENTARIOS = {}

def obtener_frase_base(categoria, pool_frases):
    """Maneja el barajado de frases para garantizar 0 repeticiones."""
    global PILAS_COMENTARIOS
    if categoria not in PILAS_COMENTARIOS or not PILAS_COMENTARIOS[categoria]:
        temp_pool = [str(f) for f in pool_frases] # Forzamos a string para evitar TypeErrors
        random.shuffle(temp_pool)
        PILAS_COMENTARIOS[categoria] = temp_pool
    return PILAS_COMENTARIOS[categoria].pop()

def generar_comentario(datos_de_fila, nombre_categoria, rank_posicion):
    """
    Motor de Narrativa Pro Chile: 20+ variantes por sección.
    Léxico corregido (piscina) e inyección dinámica de identidad.
    """
    atleta_actual = str(datos_de_fila.get('Deportista', 'Atleta TYM'))
    tiempo_actual = str(datos_de_fila.get(nombre_categoria, "00:00:00"))
    
    # --- BANCO DE NARRATIVA CHILENA (20-25 FRASES POR SECCIÓN) ---
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

    # SELECCIÓN Y FORMATEO SEGURO
    cat_key = 'General' if nombre_categoria in ['Completos', 'General'] else nombre_categoria
    if cat_key not in pools:
        return f"Desempeño consistente de {atleta_actual} en {nombre_categoria}."

    frase_plantilla = str(obtener_frase_base(cat_key, pools[cat_key]))
    
    # REEMPLAZO DINÁMICO (Seguro contra duplicidad)
    comentario_final = frase_plantilla.replace("{atleta}", atleta_actual).replace("{tiempo}", tiempo_actual)
    
    # Distinción Líder
    if rank_posicion == 1 and cat_key == 'General':
        comentario_final = f"🏆 {comentario_final.replace(atleta_actual, f'nuestro líder {atleta_actual}')}"
    
    return comentario_final

# *****************************************************************************
# --- 4. PARSERS DE ENTRADA (BLINDADO - NO SINTETIZAR) ---
# *****************************************************************************

def parse_raw_data(bloque_input_strava):
    """
    Procesa el bloque de texto copiado de Strava (Tiempo Total).
    No utiliza síntesis; cada paso de extracción es explícito y visible.
    """
    lista_de_registros_atleta = []
    valor_rank_contador = 1
    
    # Limpieza de caracteres de control web (espacios de no ruptura)
    bloque_input_strava = bloque_input_strava.replace('\xa0', ' ')
    lineas_encontradas = bloque_input_strava.strip().split('\n')
    
    for fila_texto in lineas_encontradas:
        if not fila_texto:
            continue
            
        if 'Deportista' in fila_texto:
            continue
            
        try:
            # Expresión regular para detectar tiempos con formato h y min
            patron_tiempos = r'(\d+h\s*\d*min|\d+h|\d+min|--:--)'
            tiempos_en_linea = re.findall(patron_tiempos, fila_texto)
            
            # 🛡️ CORRECCIÓN SINTAXIS AUDITADA:
            if not tiempos_en_linea:
                continue
                
            # El Tiempo Total es siempre el primer elemento detectado
            string_del_total = tiempos_en_linea[0]
            ubicacion_del_tiempo = fila_texto.find(string_del_total)
            
            # El nombre del deportista precede a la cifra de tiempo
            segmento_del_nombre = fila_texto[:ubicacion_del_tiempo].strip()
            
            # Limpieza del número de ranking si está presente en el copiado (ej: "1 Rodrigo")
            nombre_limpio_final = re.sub(r'^\d+\s*', '', segmento_del_nombre).strip()
            
            # Conversión de los bloques de tiempo a minutos enteros
            minutos_volumen_total = to_mins(string_del_total)
            
            minutos_nat = 0
            if len(tiempos_en_linea) > 1:
                minutos_nat = to_mins(tiempos_en_linea[1])
                
            minutos_bici = 0
            if len(tiempos_en_linea) > 2:
                minutos_bici = to_mins(tiempos_en_linea[2])
                
            minutos_trote = 0
            if len(tiempos_en_linea) > 3:
                minutos_trote = to_mins(tiempos_en_linea[3])
                
            # Cálculo del Coeficiente de Variación (CV)
            lista_tiempos_cv = [minutos_nat, minutos_bici, minutos_trote]
            
            if 0 in lista_tiempos_cv:
                valor_cv_final = "NC"
            else:
                calculo_std = np.std(lista_tiempos_cv)
                calculo_mean = np.mean(lista_tiempos_cv)
                valor_cv_final = round(calculo_std / calculo_mean, 4)
            
            # Extracción del conteo de actividades (dato tras el tiempo total)
            segmento_final_linea = fila_texto[ubicacion_del_tiempo + len(string_del_total):]
            match_de_actividades = re.search(r'\d+', segmento_final_linea)
            
            numero_de_actividades = 0
            if match_de_actividades:
                numero_de_actividades = int(match_de_actividades.group())
            
            # Construcción del registro detallado por cada deportista
            diccionario_de_atleta = {
                '#': valor_rank_contador,
                'Deportista': nombre_limpio_final,
                'Tiempo Total': to_hhmmss_display(minutos_volumen_total),
                'Actividades': numero_de_actividades,
                'Natación': to_hhmmss_display(minutos_nat),
                'Bicicleta': to_hhmmss_display(minutos_bici),
                'Trote': to_hhmmss_display(minutos_trote),
                'CV': valor_cv_final,
                'T_Mins': minutos_volumen_total,
                'N_Mins': minutos_nat,
                'B_Mins': minutos_bici,
                'R_Mins': minutos_trote
            }
            
            lista_de_registros_atleta.append(diccionario_de_atleta)
            valor_rank_contador = valor_rank_contador + 1
            
        except Exception:
            # Omisión de líneas corruptas o sin formato válido
            continue
            
    # Retorno estructurado para procesamiento masivo en hojas de cálculo
    df_resultado_parsing = pd.DataFrame(lista_de_registros_atleta)
    
    return df_resultado_parsing

def parse_ocr_data(texto_ocr_crudo):
    """
    Parser de Ingeniería para Formato Vertical:
    Detecta el patrón [Nombre] [Nombre] [Valor] y lo traduce a podios.
    Blindado contra duplicidad de nombres en la misma línea.
    """
    # 1. Limpieza inicial: quitamos líneas vacías y encabezados de ruido
    lineas = [l.strip() for l in texto_ocr_crudo.split('\n') if l.strip()]
    palabras_ruido = ['tiempo', 'distancia', 'actividad', 'larga', 'total', 'clasificación']
    lineas_limpias = [l for l in lineas if not any(r in l.lower() for r in palabras_ruido)]
    
    podio_distancia = []
    podio_larga = []
    
    # 2. Procesamiento de bloques (Nombre -> Valor)
    # Iteramos saltando de 2 en 2, asumiendo que el nombre es la base
    i = 0
    while i < len(lineas_limpias) - 1:
        item_nombre = lineas_limpias[i]
        
        # Lógica para limpiar nombres duplicados (ej: "Claudio Claudio")
        palabras_nombre = item_nombre.split()
        mitad = len(palabras_nombre) // 2
        if mitad > 0 and palabras_nombre[:mitad] == palabras_nombre[mitad:]:
            nombre_final = " ".join(palabras_nombre[:mitad])
        else:
            nombre_final = item_nombre
            
        valor = lineas_limpias[i+1]
        
        # 3. Clasificación por naturaleza del dato
        # Si tiene ',' o 'km', es Distancia Total
        if ',' in valor or 'km' in valor.lower():
            podio_distancia.append({'nombre': nombre_final, 'valor': valor})
            i += 2
        # Si tiene ':' es un tiempo (Actividad Larga)
        elif ':' in valor:
            podio_larga.append({'nombre': nombre_final, 'valor': valor})
            i += 2
        else:
            # Si la línea siguiente no es un valor válido, saltamos solo 1 para buscar el par
            i += 1

    # Retornamos los Top 3 de cada categoría para el Reporte Word
    return podio_distancia[:3], podio_larga[:3]

# *****************************************************************************
# --- 5. MOTOR DE ACTUALIZACIÓN DEL MAESTRO (BLINDADO - TOTAL INTEGRIDAD) ---
# *****************************************************************************

def actualizar_maestro_tym(dict_dfs_originales, df_semana_actual, nombre_nueva_columna):
    """
    Actualiza el libro Excel completo preservando TODA la historia previa.
    Garantiza la existencia de: Tiempo Total, Natación, Bicicleta, Trote y CV.
    """
    dict_dfs_actualizados = {}
    
    # 1. Preparación de la llave de cruce (MatchKey) para evitar errores de nombre
    df_semana_actual['MatchKey'] = df_semana_actual['Deportista'].apply(clean_string)
    
    # Definición explícita de las hojas que componen el ecosistema del Maestro
    hojas_a_procesar = {
        'Tiempo Total': 'T_Mins',
        'Natación': 'N_Mins',
        'Bicicleta': 'B_Mins',
        'Trote': 'R_Mins',
        'CV': 'CV'
    }
    
    for nombre_hoja, col_origen in hojas_a_procesar.items():
        # Si la hoja existe en el archivo cargado, la procesamos
        if nombre_hoja in dict_dfs_originales:
            df_maestro_hoja = dict_dfs_originales[nombre_hoja].copy()
            
            # Identificamos la columna de identidad (Nombre, Deportista o la primera)
            col_identidad = 'Nombre' if 'Nombre' in df_maestro_hoja.columns else \
                            ('Deportista' if 'Deportista' in df_maestro_hoja.columns else df_maestro_hoja.columns[0])
            
            df_maestro_hoja['MatchKey'] = df_maestro_hoja[col_identidad].apply(clean_string)
            
            # 2. Extraemos la novedad de la semana actual
            df_novedad = df_semana_actual[['MatchKey', col_origen]].copy()
            
            # Regla 7: Conversión a fracción de Excel para que sea sumable (excepto CV)
            if col_origen != 'CV':
                df_novedad[nombre_nueva_columna] = df_novedad[col_origen].apply(lambda x: x / 1440.0)
            else:
                df_novedad[nombre_nueva_columna] = df_novedad[col_origen]
            
            # Eliminamos duplicados en la novedad para no corromper el merge
            df_novedad = df_novedad.drop_duplicates(subset=['MatchKey'], keep='first')
            
            # 3. MERGE OUTER: El corazón de la persistencia histórica
            # how='outer' mantiene Sem 01, Sem 02... y pega la nueva al final.
            df_final_hoja = pd.merge(
                df_maestro_hoja, 
                df_novedad[['MatchKey', nombre_nueva_columna]], 
                on='MatchKey', 
                how='outer'
            )
            
            # 4. Gestión de Atletas Nuevos
            # Si un atleta aparece en la semana pero no estaba en el maestro, llenamos su nombre
            mask_nombre_vacio = df_final_hoja[col_identidad].isna()
            nombres_mapeo = df_semana_actual.set_index('MatchKey')['Deportista'].to_dict()
            df_final_hoja.loc[mask_nombre_vacio, col_identidad] = df_final_hoja.loc[mask_nombre_vacio, 'MatchKey'].map(nombres_mapeo)
            
            # Rellenar con 0 (o NC) los vacíos de quienes no entrenaron esta semana
            if col_origen == 'CV':
                df_final_hoja[nombre_nueva_columna] = df_final_hoja[nombre_nueva_columna].fillna('NC')
            else:
                df_final_hoja[nombre_nueva_columna] = df_final_hoja[nombre_nueva_columna].fillna(0)
            
            # 5. Recálculo de Promedios y Totales (si existen en el archivo original)
            cols_semanas = [c for c in df_final_hoja.columns if 'Sem' in str(c)]
            
            if 'Promedio' in df_final_hoja.columns and col_origen != 'CV':
                # Promedio aritmético de todas las semanas registradas hasta ahora
                df_final_hoja['Promedio'] = df_final_hoja[cols_semanas].mean(axis=1)
                
            if 'Tiempo Acumulado' in df_final_hoja.columns and col_origen != 'CV':
                # Suma total de todas las semanas
                df_final_hoja['Tiempo Acumulado'] = df_final_hoja[cols_semanas].sum(axis=1)

            # Guardamos la hoja procesada
            dict_dfs_actualizados[nombre_hoja] = df_final_hoja.drop(columns=['MatchKey'], errors='ignore')
        else:
            # Si la hoja NO existe en el maestro, la creamos desde cero para no romper el libro
            # (Esto es útil si el usuario sube un Excel incompleto)
            df_nueva = df_semana_actual[['Deportista', col_origen]].copy()
            df_nueva.rename(columns={'Deportista': 'Nombre', col_origen: nombre_nueva_columna}, inplace=True)
            if col_origen != 'CV':
                df_nueva[nombre_nueva_columna] = df_nueva[nombre_nueva_columna].apply(lambda x: x / 1440.0)
            dict_dfs_actualizados[nombre_hoja] = df_nueva

    # 6. Preservación de Hojas de Referencia (Número de Semana, Calendario, etc.)
    for hoja_restante in dict_dfs_originales:
        if hoja_restante not in hojas_a_procesar:
            dict_dfs_actualizados[hoja_restante] = dict_dfs_originales[hoja_restante]
            
    return dict_dfs_actualizados

def save_maestro_to_excel(dict_dfs):
    """
    Escribe físicamente el libro Excel con todas las pestañas actualizadas.
    """
    output_binario = io.BytesIO()
    # Usamos xlsxwriter para máxima compatibilidad con formatos de Excel
    with pd.ExcelWriter(output_binario, engine='xlsxwriter') as writer:
        for nombre_hoja, df_contenido in dict_dfs.items():
            df_contenido.to_excel(writer, sheet_name=nombre_hoja, index=False)
    
    return output_binario.getvalue()

# *****************************************************************************
# --- 6. ORQUESTADOR DE ENTREGABLES (BLINDADO - NO SINTETIZAR) ---
# *****************************************************************************

def generar_entregables_finales(df_final, dict_maestro_upd, tag_semana, podio_dist, podio_larga):
    """
    Genera el ZIP que contiene:
    1. El Excel Maestro actualizado (Preservando toda la historia).
    2. Las Fichas Individuales Word (Con KPIs, Gráficos y Narrativa).
    3. El Reporte Grupal Semanal (Insights y TOP 15).
    """
    # Buffer principal para el archivo ZIP
    zip_buffer_final = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer_final, "a", zipfile.ZIP_DEFLATED, False) as zf:
        
        # --- SUB-PROCESO 6.1: EXCEL DE PROCESO (MAESTRO ACTUALIZADO) ---
        # Se genera el binario usando el motor de la Sección 5
        archivo_excel_maestro = save_maestro_to_excel(dict_maestro_upd)
        zf.writestr(f"01_Estadisticas_Actualizadas_{tag_semana}.xlsx", archivo_excel_maestro)
        
        # --- SUB-PROCESO 6.2: REPORTE GRUPAL (INSIGHTS) ---
        # Documento que resume el desempeño del club esta semana
        doc_grupal = Document()
        doc_grupal.add_heading(f"Reporte Semanal Club TYM Triatlón - {tag_semana}", 0)
        
        # Resumen Estadístico
        num_deportistas = len(df_final)
        num_completos = len(df_final[df_final['Es_Completo'] == True])
        mins_totales_club = df_final['T_Mins'].sum()
        
        p_resumen = doc_grupal.add_paragraph()
        p_resumen.add_run(f"Total deportistas registrados: ").bold = True
        p_resumen.add_run(f"{num_deportistas}\n")
        p_resumen.add_run(f"Triatletas completos: ").bold = True
        p_resumen.add_run(f"{num_completos}\n")
        p_resumen.add_run(f"Horas totales del club: ").bold = True
        p_resumen.add_run(f"{format_duracion_larga(mins_totales_club)}")

        # Tabla TOP 15 Adherencia (Regla 4.5)
        doc_grupal.add_heading(f"🏆 TOP 15 ADHERENCIA GLOBAL", level=1)
        tabla_top = doc_grupal.add_table(rows=1, cols=4)
        tabla_top.style = 'Light Grid Accent 1'
        h_top = tabla_top.rows[0].cells
        h_top[0].text, h_top[1].text = '#', 'Deportista'
        h_top[2].text, h_top[3].text = 'TPI Global', 'Tiempo Total'
        
        df_top_15 = df_final[df_final['Es_Completo'] == True].sort_values('TPI_Global', ascending=False).head(15)
        
        for i, (_, row_top) in enumerate(df_top_15.iterrows(), 1):
            r_top = tabla_top.add_row().cells
            r_top[0].text = str(i)
            r_top[1].text = str(row_top['Deportista'])
            r_top[2].text = f"{row_top['TPI_Global']:.1f}%"
            r_top[3].text = str(row_top['Tiempo Total'])

        # Guardar Reporte Grupal
        buffer_grupal = io.BytesIO()
        doc_grupal.save(buffer_grupal)
        zf.writestr(f"02_Reporte_General_{tag_semana}.docx", buffer_grupal.getvalue())

        # --- SUB-PROCESO 6.3: FICHAS INDIVIDUALES (CLIENTE) ---
        # Generación masiva de reportes por cada triatleta activo
        for index_f, row_f in df_final.iterrows():
            # Filtro de actividad: Si sumó minutos en cualquier disciplina
            if row_f['T_Mins'] > 0:
                doc_indiv = Document()
                
                # Encabezado con Identidad
                doc_indiv.add_heading(f"Análisis de Rendimiento: {row_f['Deportista']}", 0)
                doc_indiv.add_paragraph(f"Semana de Entrenamiento: {tag_semana}")
                
                # Bloque TPI (Adherencia)
                doc_indiv.add_heading("🎯 Adherencia al Plan", level=1)
                p_tpi = doc_indiv.add_paragraph("Tu índice TPI Global esta semana es de: ")
                p_tpi.add_run(f"{row_f['TPI_Global']:.1f}%").bold = True
                
                # Tabla de Desglose de Disciplinas
                tabla_ind = doc_indiv.add_table(rows=1, cols=4)
                tabla_ind.style = 'Table Grid'
                h_ind = tabla_ind.rows[0].cells
                h_ind[0].text, h_ind[1].text = 'Disciplina', 'Real (HH:MM)'
                h_ind[2].text, h_ind[3].text = 'Meta (Hrs)', 'TPI %'
                
                for d_nom in ['Natacion', 'Ciclismo', 'Trote']:
                    celdas_d = tabla_ind.add_row().cells
                    celdas_d[0].text = d_nom
                    # Mapeo de columnas internas de la Sección 3 y 4
                    # Natacion -> N_Mins, Ciclismo -> B_Mins, Trote -> R_Mins
                    col_mins = 'N_Mins' if d_nom == 'Natacion' else ('B_Mins' if d_nom == 'Ciclismo' else 'R_Mins')
                    celdas_d[1].text = format_duracion_larga(row_f[col_mins])
                    celdas_d[2].text = f"{row_f[f'{d_nom}_Hrs_Plan']:.1f}h"
                    celdas_d[3].text = f"{row_f[f'TPI_{d_nom}']:.1f}%"
                
                # Inyección de Gráfico Comparativo (Sección 4)
                reales_lista = [row_f['N_Mins']/60, row_f['B_Mins']/60, row_f['R_Mins']/60]
                metas_lista = [row_f['Natacion_Hrs_Plan'], row_f['Ciclismo_Hrs_Plan'], row_f['Trote_Hrs_Plan']]
                
                buffer_grafico = generar_grafico_comparativo(row_f['Deportista'], reales_lista, metas_lista)
                doc_indiv.add_paragraph("\n")
                doc_indiv.add_picture(buffer_grafico, width=Inches(5))
                
                # Inyección de Comentarios Narrativos (Sección 3)
                doc_indiv.add_heading("📝 Análisis Técnico", level=1)
                # Se utiliza el rank para la frase de líder
                ranking_atleta = int(row_f['#'])
                comentario_texto = generar_comentario(row_f, 'General', ranking_atleta)
                doc_indiv.add_paragraph(comentario_texto)
                
                # Firma de Marca
                doc_indiv.add_paragraph("\n" + "─"*50)
                doc_indiv.add_paragraph("Generado por Agente TYM 2026").alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # Empaquetado en el ZIP con nombre normalizado
                buffer_word_ind = io.BytesIO()
                doc_indiv.save(buffer_word_ind)
                nombre_archivo_word = f"Fichas/Ficha_{clean_string(row_f['Deportista'])}.docx"
                zf.writestr(nombre_archivo_word, buffer_word_ind.getvalue())
                
    zip_buffer_final.seek(0)
    return zip_buffer_final

# *****************************************************************************
# --- 7. INTERFAZ DE USUARIO (UI) Y FLUJO PRINCIPAL ---
# *****************************************************************************

# Inicialización de estados para evitar recargas innecesarias
if 'datos_procesados' not in st.session_state:
    st.session_state['datos_procesados'] = None
if 'maestro_actualizado' not in st.session_state:
    st.session_state['maestro_actualizado'] = None

with st.sidebar:
    st.image("https://images.unsplash.com/photo-1517649763962-0c623066013b?q=80&w=2070", caption="TYM Performance Lab")
    st.header("⚙️ Configuración de Entrada")
    
    # 📥 CARGA DE ARCHIVOS
    archivo_maestro = st.file_uploader("1. Subir Excel Maestro (Histórico)", type=["xlsx"])
    
    st.divider()
    st.subheader("🎯 Metas Globales (Plan)")
    # Estos valores se usan si el atleta no tiene un plan individual cargado
    col_n1, col_n2 = st.columns(2)
    with col_n1: g_n_h = st.number_input("Natación (Hrs)", 0.0, 10.0, 3.0)
    with col_n2: g_n_s = st.number_input("Natación (Ses)", 0, 10, 3)
    
    col_c1, col_c2 = st.columns(2)
    with col_c1: g_c_h = st.number_input("Ciclismo (Hrs)", 0.0, 20.0, 5.0)
    with col_c2: g_c_s = st.number_input("Ciclismo (Ses)", 0, 10, 3)
    
    col_t1, col_t2 = st.columns(2)
    with col_t1: g_t_h = st.number_input("Trote (Hrs)", 0.0, 15.0, 3.0)
    with col_t2: g_t_s = st.number_input("Trote (Ses)", 0, 10, 3)
    
    meta_global = {
        'Natacion_Hrs_Plan': g_n_h, 'Natacion_Ses_Plan': g_n_s,
        'Ciclismo_Hrs_Plan': g_c_h, 'Ciclismo_Ses_Plan': g_c_s,
        'Trote_Hrs_Plan': g_t_h, 'Trote_Ses_Plan': g_t_s
    }

# --- CUERPO PRINCIPAL ---
st.info("💡 Copia los datos de Strava 'Tiempo Total' y pégalos en el cuadro de abajo.")
entrada_texto_strava = st.text_area("Datos Crudos de Strava:", height=200, placeholder="1 Francisco Ramírez 14h 22min...")

tag_semana_actual = st.text_input("Etiqueta de la Semana (Ej: Sem 09):", "Sem 08")

# Botón de Procesamiento
if st.button("🚀 PROCESAR Y GENERAR ENTREGABLES"):
    if not entrada_texto_strava:
        st.error("❌ Error: No hay datos de Strava para procesar.")
    elif not archivo_maestro:
        st.error("❌ Error: Debes cargar el archivo Maestro para actualizar el historial.")
    else:
        with st.spinner("Ejecutando Pipeline de Ingeniería..."):
            # 1. PARSING (Sección 4)
            df_semanal = parse_raw_data(entrada_texto_strava)
            
            # 2. CARGA DE MAESTRO (Sección 5)
            # Leemos todas las hojas para no perder información
            dict_maestro_full = pd.read_excel(archivo_maestro, sheet_name=None)
            
            # 3. CÁLCULO DE KPIS Y ADHERENCIA (Lógica TPI)
            # Inyectamos las metas (esto puede expandirse a carga de Excel de Plan)
            for d in ['Natacion', 'Ciclismo', 'Trote']:
                df_semanal[f'{d}_Hrs_Plan'] = meta_global[f'{d}_Hrs_Plan']
                df_semanal[f'{d}_Ses_Plan'] = meta_global[f'{d}_Ses_Plan']
            
            # Cálculo de TPI Individual (Regla 4.3)
            def aplicar_tpi(row):
                # Natación
                vci_n = (row['N_Mins'] / (row['Natacion_Hrs_Plan']*60)) * 100 if row['Natacion_Hrs_Plan'] > 0 else 0
                sei_n = 100 if row['N_Mins'] > 0 else 0 # Simplificado para esta versión
                tpi_n = (vci_n * 0.4) + (sei_n * 0.6)
                
                # Ciclismo
                vci_c = (row['B_Mins'] / (row['Ciclismo_Hrs_Plan']*60)) * 100 if row['Ciclismo_Hrs_Plan'] > 0 else 0
                sei_c = 100 if row['B_Mins'] > 0 else 0
                tpi_c = (vci_c * 0.4) + (sei_c * 0.6)
                
                # Trote
                vci_t = (row['R_Mins'] / (row['Trote_Hrs_Plan']*60)) * 100 if row['Trote_Hrs_Plan'] > 0 else 0
                sei_t = 100 if row['R_Mins'] > 0 else 0
                tpi_t = (vci_t * 0.4) + (sei_t * 0.6)
                
                tpi_g = np.mean([tpi_n, tpi_c, tpi_t])
                es_comp = row['N_Mins'] > 0 and row['B_Mins'] > 0 and row['R_Mins'] > 0
                
                return pd.Series([tpi_n, tpi_c, tpi_t, tpi_g, es_comp])

            df_semanal[['TPI_Natacion', 'TPI_Ciclismo', 'TPI_Trote', 'TPI_Global', 'Es_Completo']] = df_semanal.apply(aplicar_tpi, axis=1)

            # 4. ACTUALIZACIÓN DE MAESTRO (Sección 5)
            maestro_upd = actualizar_maestro_tym(dict_maestro_full, df_semanal, tag_semana_actual)
            
            # 5. GENERACIÓN DE ZIP (Sección 6)
            # Nota: Los podios OCR pueden integrarse aquí o dejarse vacíos si no hay texto OCR
            zip_final = generar_entregables_finales(df_semanal, maestro_upd, tag_semana_actual, [], [])
            
            # Guardamos en Session State
            st.session_state['datos_procesados'] = df_semanal
            st.session_state['maestro_actualizado'] = zip_final
            st.success("✅ Procesamiento Exitoso.")

# --- ZONA DE DESCARGAS Y VISUALIZACIÓN ---
if st.session_state['datos_procesados'] is not None:
    df_f = st.session_state['datos_procesados']
    
    st.divider()
    st.subheader(f"🏆 Resumen de Desempeño: {tag_semana_actual}")
    
    col_m1, col_m2, col_m3 = st.columns(3)
    col_m1.metric("Atletas Activos", len(df_f))
    col_m2.metric("Triatletas Completos", len(df_f[df_f['Es_Completo'] == True]))
    col_m3.metric("TPI Promedio Club", f"{df_f['TPI_Global'].mean():.1f}%")

    # Mostrar TOP 15 en pantalla
    st.dataframe(df_f[df_f['Es_Completo'] == True].sort_values('TPI_Global', ascending=False).head(15)[['Deportista', 'TPI_Global', 'Tiempo Total']])

    # BOTÓN DE DESCARGA FINAL
    st.download_button(
        label="📥 DESCARGAR PACK COMPLETO (ZIP)",
        data=st.session_state['maestro_actualizado'].getvalue(),
        file_name=f"Pack_TYM_{tag_semana_actual}.zip",
        mime="application/zip",
        use_container_width=True
    )
