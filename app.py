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
    page_title="Plataforma TYM 2026 - V2.5.0 (Full Engine)", 
    page_icon="🏆", 
    layout="wide"
)

st.title("🏆 Gestión de Reportes y Estadísticas - Club TYM")

# *****************************************************************************
# --- 2. UTILIDADES DE PROCESAMIENTO Y TIEMPO (BLINDADO - NO SINTETIZAR) ---
# *****************************************************************************

def clean_string(text):
    if text is None or pd.isna(text): return ""
    nombre_limpio_temp = str(text).strip().upper()
    info_normalizada = unicodedata.normalize('NFKD', nombre_limpio_temp)
    return "".join([c for c in info_normalizada if not unicodedata.combining(c)])

def to_mins(valor):
    if pd.isna(valor): return 0
    s = str(valor).strip()
    if s in ['--:--', '0', '', '00:00:00', 'NC']: return 0
    try:
        if isinstance(valor, (float, int)): return int(round(valor * 1440))
        if isinstance(valor, (time, datetime)): return (valor.hour * 60) + valor.minute
        if ':' in s:
            partes = s.split(':')
            return (int(partes[0]) * 60) + int(partes[1].split('.')[0])
        h = re.search(r'(\d+)h', s); m = re.search(r'(\d+)min', s)
        return (int(h.group(1)) * 60 if h else 0) + (int(m.group(1)) if m else 0)
    except: return 0

def to_hhmmss_display(minutos):
    hrs = int(minutos // 60)
    mins = int(minutos % 60)
    return f"{hrs:02d}:{mins:02d}:00"

def format_duracion_larga(minutos):
    """Alias de seguridad para evitar NameError en Sección 6."""
    return to_hhmmss_display(minutos)

# =============================================================================
# SECCIÓN 3: MOTOR NARRATIVO PRO CHILE (INTEGRIDAD TOTAL - 75+ FRASES)
# =============================================================================

PILAS_COMENTARIOS = {}

def obtener_frase_base(categoria, pool_frases):
    global PILAS_COMENTARIOS
    if categoria not in PILAS_COMENTARIOS or not PILAS_COMENTARIOS[categoria]:
        temp_pool = [str(f) for f in pool_frases]
        random.shuffle(temp_pool)
        PILAS_COMENTARIOS[categoria] = temp_pool
    return PILAS_COMENTARIOS[categoria].pop()

def generar_comentario(datos_de_fila, nombre_categoria, rank_posicion):
    atleta = str(datos_de_fila.get('Deportista', 'Atleta TYM'))
    tiempo = str(datos_de_fila.get(nombre_categoria, "00:00:00"))
    
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
            "Solidez transversal. {atleta} se consolidida como uno de los atletas más balanceados.",
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
    cat_key = 'General' if nombre_categoria in ['Completos', 'General'] else nombre_categoria
    frase_plantilla = str(obtener_frase_base(cat_key, pools.get(cat_key, pools['General'])))
    comentario_final = frase_plantilla.replace("{atleta}", atleta).replace("{tiempo}", tiempo)
    if rank_posicion == 1 and cat_key == 'General':
        comentario_final = f"🏆 {comentario_final.replace(atleta, f'nuestro líder {atleta}')}"
    return comentario_final

# =============================================================================
# SECCIÓN 4: PARSERS DE ENTRADA (MAESTRO, SEMANAL Y PLAN)
# =============================================================================

def procesar_excel_semanal_robusto(archivo):
    """Detecta dinámicamente columnas de tiempo en el Excel de Strava."""
    df_raw = pd.read_excel(archivo)
    cols = df_raw.columns
    c_nom = next((c for c in cols if 'DEPORTISTA' in c.upper() or 'NOMBRE' in c.upper()), cols[0])
    c_nat = next((c for c in cols if 'NATAC' in c.upper() or 'PISCINA' in c.upper()), None)
    c_bic = next((c for c in cols if 'BICI' in c.upper() or 'CICLIS' in c.upper() or 'RODILLO' in c.upper()), None)
    c_tro = next((c for c in cols if 'TROTE' in c.upper() or 'RUN' in c.upper()), None)
    
    df_clean = pd.DataFrame()
    df_clean['Deportista'] = df_raw[c_nom]
    df_clean['N_Mins'] = df_raw[c_nat].apply(to_mins) if c_nat else 0
    df_clean['B_Mins'] = df_raw[c_bic].apply(to_mins) if c_bic else 0
    df_clean['R_Mins'] = df_raw[c_tro].apply(to_mins) if c_tro else 0
    df_clean['T_Mins'] = df_clean['N_Mins'] + df_clean['B_Mins'] + df_clean['R_Mins']
    df_clean['Tiempo Total'] = df_clean['T_Mins'].apply(to_hhmmss_display)
    return df_clean

def procesar_excel_plan_individual(archivo_plan):
    """Carga el plan individual y normaliza columnas para el cruce de KPIs."""
    if not archivo_plan: return pd.DataFrame()
    df_p = pd.read_excel(archivo_plan)
    df_p['MatchKey'] = df_p.iloc[:, 0].apply(clean_string)
    # Estandarización de nombres de columnas de plan (Hrs y Ses)
    return df_p

# *****************************************************************************
# --- 5. MOTOR DE ACTUALIZACIÓN DEL MAESTRO (VERSIÓN EXTENDIDA Y BLINDADA) ---
# *****************************************************************************

def actualizar_maestro_tym(dict_dfs_originales, df_semana_actual, nombre_nueva_columna):
    """
    Actualiza el libro Maestro con integridad total. 
    Limpia duplicados previos, purifica tipos de datos y sincroniza todas las disciplinas.
    Resuelve el KeyError asegurando que las columnas de origen existan antes de indexar.
    """
    dict_dfs_actualizados = {}
    
    # 1. PREPARACIÓN DE LA SEMANA ACTUAL
    # Generamos la llave de cruce MatchKey para todos los deportistas de la semana
    df_semana_actual['MatchKey'] = df_semana_actual['Deportista'].apply(clean_string)
    
    # 2. MAPEO EXTENDIDO DE DISCIPLINAS (PARA EVITAR HOJAS SIN ACTUALIZAR)
    # Vinculamos el nombre de la hoja con los posibles nombres de columnas generados en el proceso
    mapeo_hojas_a_datos = {
        'TIEMPO TOTAL': ['T_Mins', 'Tiempo Total', 'T_MINS'],
        'NATACION': ['N_Mins', 'Natación', 'N_MINS'],
        'NATACIÓN': ['N_Mins', 'Natación', 'N_MINS'],
        'BICICLETA': ['B_Mins', 'Bicicleta', 'B_MINS'],
        'CICLISMO': ['B_Mins', 'Bicicleta', 'B_MINS'], 
        'TROTE': ['R_Mins', 'Trote', 'R_MINS'],
        'RUNNING': ['R_Mins', 'Trote', 'R_MINS'],
        'CV': ['CV', 'cv']
    }
    
    # Identificamos las hojas reales disponibles en el Excel cargado
    hojas_reales_en_excel = {clean_string(nombre): nombre for nombre in dict_dfs_originales.keys()}
    
    # 3. PROCESAMIENTO SISTEMÁTICO DE CADA HOJA DEL MAESTRO
    for nombre_hoja_original in dict_dfs_originales.keys():
        nombre_normalizado = clean_string(nombre_hoja_original)
        
        # Verificamos si la hoja actual es una de las que requiere actualización
        if nombre_normalizado in mapeo_hojas_a_datos:
            lista_posibles_columnas = mapeo_hojas_a_datos[nombre_normalizado]
            
            # Buscamos cuál de los nombres de columna existe realmente en df_semana_actual
            columna_datos_strava = next((c for c in lista_posibles_columnas if c in df_semana_actual.columns), None)
            
            if columna_datos_strava:
                df_hoja_historia = dict_dfs_originales[nombre_hoja_original].copy()
                
                # Identificamos la columna de identidad (Nombre, Deportista o la primera)
                col_id_identidad = 'Nombre' if 'Nombre' in df_hoja_historia.columns else \
                                   ('Deportista' if 'Deportista' in df_hoja_historia.columns else df_hoja_historia.columns[0])
                
                # Generamos MatchKey en el historial para el cruce exacto
                df_hoja_historia['MatchKey'] = df_hoja_historia[col_id_identidad].apply(clean_string)
                
                # --- SOLUCIÓN AL ERROR DE DUPLICADOS (Sem 08_x / Sem 08_y) ---
                # Si en el Maestro ya existe una columna con el mismo nombre, la eliminamos antes de escribir
                if nombre_nueva_columna in df_hoja_historia.columns:
                    df_hoja_historia = df_hoja_historia.drop(columns=[nombre_nueva_columna])
                
                # Preparamos la novedad (datos Strava)
                df_novedad_semanal = df_semana_actual[['MatchKey', columna_datos_strava]].copy()
                
                # --- PURIFICACIÓN DE DATOS (Regla 7: 1.0 = 24h) ---
                if columna_datos_strava != 'CV':
                    # Aseguramos que el valor sea numérico antes de dividir por 1440
                    df_novedad_semanal[nombre_nueva_columna] = df_novedad_semanal[columna_datos_strava].apply(
                        lambda x: (to_mins(x) / 1440.0) if pd.notna(x) else 0.0
                    )
                else:
                    df_novedad_semanal[nombre_nueva_columna] = df_novedad_semanal[columna_datos_strava]
                
                # Limpieza de duplicados en la novedad antes de la unión
                df_novedad_semanal = df_novedad_semanal.drop_duplicates(subset=['MatchKey'], keep='first')
                
                # --- MERGE DE INTEGRIDAD ---
                df_actualizado_hoja = pd.merge(
                    df_hoja_historia, 
                    df_novedad_semanal[['MatchKey', nombre_nueva_columna]], 
                    on='MatchKey', 
                    how='outer'
                )
                
                # --- LIMPIEZA DE TYPES (EVITA ERROR DE SUMA Y TYPEERROR) ---
                # Convertimos todas las columnas "Sem XX" a numérico para que se puedan sumar
                columnas_semanales = [c for c in df_actualizado_hoja.columns if str(c).startswith('Sem')]
                for col_s in columnas_semanales:
                    df_actualizado_hoja[col_s] = pd.to_numeric(df_actualizado_hoja[col_s], errors='coerce').fillna(0.0)
                
                # --- GESTIÓN DE ATLETAS NUEVOS ---
                # Si el atleta es nuevo, su celda de identidad estará vacía (NaN)
                mask_atletas_nuevos = df_actualizado_hoja[col_id_identidad].isna()
                map_identidades = df_semana_actual.set_index('MatchKey')['Deportista'].to_dict()
                df_actualizado_hoja.loc[mask_atletas_nuevos, col_id_identidad] = df_actualizado_hoja.loc[mask_atletas_nuevos, 'MatchKey'].map(map_identidades)
                
                # Rellenar con 0 a quienes no entrenaron esta semana
                df_actualizado_hoja[nombre_nueva_columna] = df_actualizado_hoja[nombre_nueva_columna].fillna(0.0)
                
                # --- RECALCULO DE TOTALES ---
                if columna_datos_strava != 'CV':
                    if 'Tiempo Acumulado' in df_actualizado_hoja.columns:
                        df_actualizado_hoja['Tiempo Acumulado'] = df_actualizado_hoja[columnas_semanales].sum(axis=1)
                    
                    if 'Promedio' in df_actualizado_hoja.columns:
                        df_actualizado_hoja['Promedio'] = df_actualizado_hoja[columnas_semanales].mean(axis=1)
                
                # Guardamos la hoja limpia
                dict_dfs_actualizados[nombre_hoja_original] = df_actualizado_hoja.drop(columns=['MatchKey'], errors='ignore')
            else:
                # Si no hay datos en la semana para esta hoja, se mantiene original
                dict_dfs_actualizados[nombre_hoja_original] = dict_dfs_originales[nombre_hoja_original]
        else:
            # Hojas estáticas (Calendario, etc.) se mantienen sin cambios
            dict_dfs_actualizados[nombre_hoja_original] = dict_dfs_originales[nombre_hoja_original]
            
    return dict_dfs_actualizados

def save_maestro_to_excel(dict_dfs):
    """
    Exportación binaria multi-hoja.
    """
    output_buffer = io.BytesIO()
    with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
        for hoja, contenido in dict_dfs.items():
            contenido.to_excel(writer, sheet_name=hoja, index=False)
    return output_buffer.getvalue()
# =============================================================================
# FIN DE SECCIÓN 5
# =============================================================================

# *****************************************************************************
# --- 6. GENERADOR DE ENTREGABLES WORD Y ZIP (VERSIÓN EXTENDIDA) ---
# *****************************************************************************

def generar_entregables_finales(df_semanal_final, dict_maestro_final, tag_semana):
    """
    Genera el paquete ZIP con el Maestro y las fichas individuales.
    Resuelve el error de '0 triatletas completos' validando contra minutos reales.
    """
    zip_buffer_final = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer_final, "a", zipfile.ZIP_DEFLATED) as zf:
        
        # --- 6.1: EXCEL MAESTRO ACTUALIZADO ---
        binario_excel = save_maestro_to_excel(dict_maestro_final)
        zf.writestr(f"01_Estadisticas_Actualizadas_{tag_semana}.xlsx", binario_excel)
        
        # --- 6.2: REPORTE GRUPAL (WORD) ---
        doc_grupal = Document()
        doc_grupal.add_heading(f"Reporte Semanal Club TYM - {tag_semana}", 0)
        
        # Métricas grupales con validación profunda
        # Se considera completo si tiene minutos mayores a 0 en las 3 disciplinas
        df_completos = df_semanal_final[
            (df_semanal_final['N_Mins'] > 0) & 
            (df_semanal_final['B_Mins'] > 0) & 
            (df_semanal_final['R_Mins'] > 0)
        ]
        
        total_deportistas = len(df_semanal_final)
        total_completos = len(df_completos)
        horas_totales_club = df_semanal_final['T_Mins'].sum()
        
        # Párrafo de resumen
        p = doc_grupal.add_paragraph()
        p.add_run(f"Total deportistas registrados: ").bold = True
        p.add_run(f"{total_deportistas}\n")
        p.add_run(f"Triatletas completos esta semana (N/B/R): ").bold = True
        p.add_run(f"{total_completos}\n")
        p.add_run(f"Tiempo total acumulado por el club: ").bold = True
        p.add_run(f"{to_hhmmss_display(horas_totales_club)}")

        # Tabla de Podio TPI (TOP 15)
        doc_grupal.add_heading("🏆 TOP 15 ADHERENCIA (TPI GLOBAL)", level=1)
        tabla = doc_grupal.add_table(rows=1, cols=3)
        tabla.style = 'Light Grid Accent 1'
        hdr = tabla.rows[0].cells
        hdr[0].text, hdr[1].text, hdr[2].text = 'Posición', 'Deportista', 'TPI Global %'
        
        # Ordenamos por adherencia
        df_rank = df_semanal_final.sort_values('TPI_Global', ascending=False).head(15)
        for i, (idx, fila_r) in enumerate(df_rank.iterrows(), 1):
            r_c = tabla.add_row().cells
            r_c[0].text = str(i)
            r_c[1].text = str(fila_r['Deportista'])
            r_c[2].text = f"{fila_r['TPI_Global']:.1f}%"

        # Guardar Documento Grupal
        buf_g = io.BytesIO()
        doc_grupal.save(buf_g)
        zf.writestr(f"02_Reporte_General_{tag_semana}.docx", buf_g.getvalue())

        # --- 6.3: FICHAS INDIVIDUALES ---
        # Solo para atletas con actividad
        for _, row_a in df_semanal_final.iterrows():
            if row_a['T_Mins'] > 0:
                doc_i = Document()
                doc_i.add_heading(f"Reporte Semanal: {row_a['Deportista']}", 0)
                
                # KPIs principales
                doc_i.add_heading(f"Adherencia Semanal: {row_a['TPI_Global']:.1f}%", level=1)
                
                # Tabla Desglose
                t_i = doc_i.add_table(rows=1, cols=3); t_i.style = 'Table Grid'
                h_i = t_i.rows[0].cells
                h_i[0].text, h_i[1].text, h_i[2].text = 'Disciplina', 'Tiempo Real', 'TPI %'
                
                for disc, m_col in [('Natación', 'N_Mins'), ('Ciclismo', 'B_Mins'), ('Trote', 'R_Mins')]:
                    c_i = t_i.add_row().cells
                    c_i[0].text = disc
                    c_i[1].text = to_hhmmss_display(row_a[m_col])
                    # Verificamos si la columna de TPI existe para esa disciplina
                    tpi_key = f'TPI_{disc}' if f'TPI_{disc}' in row_a else 'TPI_Global'
                    c_i[2].text = f"{row_a.get(tpi_key, 0):.1f}%"

                doc_i.add_heading("Comentario Técnico", level=1)
                doc_i.add_paragraph(generar_comentario(row_a, 'General', 1))
                
                buf_i = io.BytesIO()
                doc_i.save(buf_i)
                nombre_archivo = f"Fichas/Ficha_{clean_string(row_a['Deportista'])}.docx"
                zf.writestr(nombre_archivo, buf_i.getvalue())

    zip_buffer_final.seek(0)
    return zip_buffer_final

# =============================================================================
# SECCIÓN 7: INTERFAZ DE USUARIO Y ORQUESTACIÓN (Sincronización Total)
# =============================================================================

# 1. Definición estática del Sidebar (Para que nunca desaparezca)
with st.sidebar:
    st.header("⚙️ Entradas de Ingeniería")
    f_maestro = st.file_uploader("1. Excel Maestro (Historial)", type=["xlsx"])
    f_semanal = st.file_uploader("2. Excel Semanal (Reales)", type=["xlsx"])
    f_plan = st.file_uploader("3. Excel Plan (Individual)", type=["xlsx"])
    st.divider()
    st.subheader("🎯 Metas Globales (Fallback)")
    
    # Definimos el diccionario meta_g explícitamente aquí
    meta_g = {
        'N_H': st.number_input("Natación (Hrs)", 3.0), 
        'N_S': st.number_input("Natación (Ses)", 3),
        'B_H': st.number_input("Ciclismo (Hrs)", 5.0), 
        'B_S': st.number_input("Ciclismo (Ses)", 3),
        'T_H': st.number_input("Trote (Hrs)", 3.0), 
        'T_S': st.number_input("Trote (Ses)", 3)
    }

# 2. Inputs en el cuerpo principal
tag_semana = st.text_input("Etiqueta de la Semana", "Sem 08")

# 3. Lógica de Procesamiento
if st.button("🚀 PROCESAR JORNADA COMPLETA"):
    if f_maestro and f_semanal:
        # A. Carga de datos
        df_s = procesar_excel_semanal_robusto(f_semanal)
        df_s['MatchKey'] = df_s['Deportista'].apply(clean_string)
        
        # B. Carga de Plan Individual (si existe)
        df_p = pd.read_excel(f_plan) if f_plan else pd.DataFrame(columns=['MatchKey'])
        if not df_p.empty:
            df_p['MatchKey'] = df_p.iloc[:, 0].apply(clean_string)
            # Unimos para tener metas individuales disponibles en la misma fila
            df_s = pd.merge(df_s, df_p, on='MatchKey', how='left', suffixes=('', '_plan'))
            
        # C. Función interna de cálculo de TPI (Regla 4.3 corregida)
        def aplicar_tpi_logica(row):
            res = {}
            mapeo = {'Natacion': 'N', 'Ciclismo': 'B', 'Trote': 'R'}
            
            for d, pref in mapeo.items():
                # Búsqueda jerárquica de metas para evitar KeyError
                # 1. ¿Está en el Excel de Plan?
                h_plan = row.get(f'{d}_Hrs_Plan_plan')
                if pd.isna(h_plan): 
                    h_plan = meta_g.get(f'{pref}_H', 0) # Fallback al Sidebar
                
                s_plan = row.get(f'{d}_Ses_Plan_plan')
                if pd.isna(s_plan): 
                    s_plan = meta_g.get(f'{pref}_S', 1) # Fallback al Sidebar (mínimo 1)
                
                # Identificación de columna de minutos reales (Sección 4)
                col_real = f'{"N" if d=="Natacion" else ("B" if d=="Ciclismo" else "R")}_Mins'
                real_m = row.get(col_real, 0)
                
                # Cálculo TPI: (Volumen * 0.4) + (Sesiones * 0.6)
                vci = (real_m / (h_plan * 60)) * 100 if h_plan > 0 else 0
                sei = (100 / s_plan) if (real_m > 0 and s_plan > 0) else 0
                
                res[f'TPI_{d}'] = min((vci * 0.4) + (sei * 0.6), 110)
                res[f'{d}_Hrs_Plan'] = h_plan
                res[f'{d}_Ses_Plan'] = s_plan

            # KPI Global y Bandera de Atleta Completo
            res['TPI_Global'] = np.mean([res['TPI_Natacion'], res['TPI_Ciclismo'], res['TPI_Trote']])
            res['Es_Completo'] = row['N_Mins']>0 and row['B_Mins']>0 and row['R_Mins']>0
            return pd.Series(res)

        # D. Ejecución del Pipeline
        # Concatenamos los resultados del TPI al DataFrame original
        df_tpi = df_s.apply(aplicar_tpi_logica, axis=1)
        df_final = pd.concat([df_s, df_tpi], axis=1)
        
        # E. Actualización del Maestro (Sección 5)
        dict_maestro_full = pd.read_excel(f_maestro, sheet_name=None)
        m_upd = actualizar_maestro_tym(dict_maestro_full, df_final, tag_semana)
        
        # F. Generación de ZIP (Sección 6)
        st.session_state['zip_out'] = generar_entregables_finales(df_final, m_upd, tag_semana)
        st.success("✅ Procesamiento completado con éxito.")
    else:
        st.error("⚠️ Debes cargar al menos el Maestro y el Semanal en el Sidebar.")

# 4. Botón de descarga (Persistente)
if 'zip_out' in st.session_state and st.session_state['zip_out'] is not None:
    st.download_button(
        "📥 DESCARGAR PACK FINAL (ZIP)", 
        st.session_state['zip_out'], 
        f"Pack_TYM_{tag_semana}.zip",
        use_container_width=True
    )
