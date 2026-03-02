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
# --- 5. MOTOR DE ACTUALIZACIÓN DEL MAESTRO (VERSION EXTENDIDA - GOLD) ---
# *****************************************************************************

def actualizar_maestro_tym(dict_dfs_originales, df_semana_actual, nombre_nueva_columna):
    """
    Actualiza el libro Maestro con integridad total. 
    Resuelve colisiones de columnas duplicadas (Sem 08_x) y descalce de nombres de hojas.
    """
    # Diccionario contenedor para el nuevo libro Excel
    dict_dfs_actualizados = {}
    
    # 1. PREPARACIÓN DE LA SEMANA ACTUAL
    # Aseguramos que la semana actual tenga la llave de cruce limpia
    df_semana_actual['MatchKey'] = df_semana_actual['Deportista'].apply(clean_string)
    
    # 2. MAPEO FLEXIBLE DE DISCIPLINAS
    # Mapeamos los nombres técnicos a todas las posibles variantes de nombres de hojas
    # Esto soluciona el problema de 'Ciclismo' vs 'Bicicleta'
    mapeo_disciplinas = {
        'TIEMPO TOTAL': 'T_Mins',
        'NATACION': 'N_Mins',
        'BICICLETA': 'B_Mins',
        'CICLISMO': 'B_Mins',  # Soporte para ambas nomenclaturas
        'TROTE': 'R_Mins',
        'RUNNING': 'R_Mins',
        'CV': 'CV'
    }
    
    # Normalizamos los nombres de las hojas que ya existen en el Maestro
    hojas_maestro_normalizadas = {clean_string(k): k for k in dict_dfs_originales.keys()}
    
    # 3. PROCESAMIENTO HOJA POR HOJA
    for nombre_hoja_maestro in dict_dfs_originales.keys():
        llave_norm = clean_string(nombre_hoja_maestro)
        
        # ¿Es esta una de las hojas que debemos actualizar con datos de Strava?
        if llave_norm in mapeo_disciplinas:
            col_origen_datos = mapeo_disciplinas[llave_norm]
            df_hoja_historia = dict_dfs_originales[nombre_hoja_maestro].copy()
            
            # Identificamos la columna de identidad (Nombre/Deportista)
            col_id = 'Nombre' if 'Nombre' in df_hoja_historia.columns else \
                     ('Deportista' if 'Deportista' in df_hoja_historia.columns else df_hoja_historia.columns[0])
            
            # Generamos MatchKey en el historial
            df_hoja_historia['MatchKey'] = df_hoja_historia[col_id].apply(clean_string)
            
            # --- SOLUCIÓN AL ERROR DE DUPLICADOS (Sem 08_x) ---
            # Si la columna 'Sem 08' ya existe en el Maestro, la eliminamos antes del merge
            # para que la nueva data sea la que mande y no se creen sufijos _x / _y
            if nombre_nueva_columna in df_hoja_historia.columns:
                df_hoja_historia = df_hoja_historia.drop(columns=[nombre_nueva_columna])
            
            # Preparamos la novedad de la semana
            df_novedad = df_semana_actual[['MatchKey', col_origen_datos]].copy()
            
            # Conversión a formato Excel (Fracción de día) excepto para CV
            if col_origen_datos != 'CV':
                df_novedad[nombre_nueva_columna] = df_novedad[col_origen_datos].apply(lambda x: x / 1440.0)
            else:
                df_novedad[nombre_nueva_columna] = df_novedad[col_origen_datos]
            
            # Eliminamos duplicados en la novedad por seguridad
            df_novedad = df_novedad.drop_duplicates(subset=['MatchKey'], keep='first')
            
            # --- MERGE DE INTEGRIDAD ---
            df_final_hoja = pd.merge(
                df_hoja_historia, 
                df_novedad[['MatchKey', nombre_nueva_columna]], 
                on='MatchKey', 
                how='outer'
            )
            
            # --- GESTIÓN DE ATLETAS NUEVOS Y NULOS ---
            # Rellenar nombres para atletas que no estaban en el maestro
            mask_nuevos = df_final_hoja[col_id].isna()
            mapeo_nombres_nuevos = df_semana_actual.set_index('MatchKey')['Deportista'].to_dict()
            df_final_hoja.loc[mask_nuevos, col_id] = df_final_hoja.loc[mask_nuevos, 'MatchKey'].map(mapeo_nombres_nuevos)
            
            # Rellenar con 0 (o NC) a quienes no entrenaron esta semana
            valor_vacio = 'NC' if col_origen_datos == 'CV' else 0
            df_final_hoja[nombre_nueva_columna] = df_final_hoja[nombre_nueva_columna].fillna(valor_vacio)
            
            # --- RECALCULO DE TOTALES (Opcional según estructura) ---
            # Si existen columnas de Promedio o Acumulado, se actualizan aquí
            # Buscamos todas las columnas que empiecen con 'Sem'
            cols_semanas = [c for c in df_final_hoja.columns if str(c).startswith('Sem')]
            
            if 'Tiempo Acumulado' in df_final_hoja.columns and col_origen_datos != 'CV':
                df_final_hoja['Tiempo Acumulado'] = df_final_hoja[cols_semanas].sum(axis=1)
                
            if 'Promedio' in df_final_hoja.columns and col_origen_datos != 'CV':
                # El promedio solo cuenta semanas con datos > 0 si se desea, 
                # aquí lo hacemos sobre el total de semanas procesadas.
                df_final_hoja['Promedio'] = df_final_hoja[cols_semanas].mean(axis=1)

            # Guardamos la hoja limpia
            dict_dfs_actualizados[nombre_hoja_maestro] = df_final_hoja.drop(columns=['MatchKey'], errors='ignore')
            
        else:
            # Si la hoja no es de una disciplina (ej: Calendario, Número de Semana), la pasamos intacta
            dict_dfs_actualizados[nombre_hoja_maestro] = dict_dfs_originales[nombre_hoja_maestro]
            
    return dict_dfs_actualizados

def save_maestro_to_excel(dict_dfs):
    """
    Genera el archivo Excel binario manteniendo todas las pestañas.
    """
    output_binario = io.BytesIO()
    # Usamos xlsxwriter para asegurar que los formatos numéricos se mantengan
    with pd.ExcelWriter(output_binario, engine='xlsxwriter') as writer:
        for nombre_hoja, df_contenido in dict_dfs.items():
            df_contenido.to_excel(writer, sheet_name=nombre_hoja, index=False)
            
            # Auto-ajuste de columnas básico (Opcional)
            worksheet = writer.sheets[nombre_hoja]
            for i, col in enumerate(df_contenido.columns):
                column_len = max(df_contenido[col].astype(str).str.len().max(), len(col)) + 2
                worksheet.set_column(i, i, column_len)
                
    return output_binario.getvalue()

# =============================================================================
# FIN DE SECCIÓN 5
# =============================================================================

# =============================================================================
# SECCIÓN 6: GENERADOR DE ENTREGABLES (BLINDADO)
# =============================================================================

def generar_grafico_comparativo(nombre, reales, metas):
    labels = ['N', 'B', 'R']
    fig, ax = plt.subplots(figsize=(4, 2.5))
    x = np.arange(len(labels))
    ax.bar(x - 0.2, metas, 0.4, label='Plan', color='#BDC3C7')
    ax.bar(x + 0.2, reales, 0.4, label='Real', color='#1E90FF')
    ax.set_xticks(x); ax.set_xticklabels(labels); ax.legend()
    buf = io.BytesIO(); plt.savefig(buf, format='png', bbox_inches='tight'); plt.close(fig)
    return buf

def generar_entregables_finales(df_final, dict_maestro_upd, tag_semana):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED) as zf:
        # Excel
        ex_buf = save_maestro_to_excel(dict_maestro_upd)
        zf.writestr(f"01_Estadisticas_Actualizadas_{tag_semana}.xlsx", ex_buf)
        
        # Reporte General
        doc_g = Document()
        doc_g.add_heading(f"Reporte Semanal Club TYM - {tag_semana}", 0)
        p = doc_g.add_paragraph()
        p.add_run(f"Total deportistas: {len(df_final)}\n").bold = True
        p.add_run(f"Triatletas completos: {len(df_final[df_final['Es_Completo']])}")
        buf_g = io.BytesIO(); doc_g.save(buf_g); zf.writestr(f"02_Reporte_General_{tag_semana}.docx", buf_g.getvalue())
        
        # Fichas Individuales
        for _, row in df_final.iterrows():
            if row['T_Mins'] > 0:
                doc = Document()
                doc.add_heading(f"Reporte TYM: {row['Deportista']}", 0)
                doc.add_paragraph(f"TPI Global: {row['TPI_Global']:.1f}%")
                ti = doc.add_table(rows=1, cols=4); ti.style = 'Table Grid'
                hi = ti.rows[0].cells
                hi[0].text, hi[1].text, hi[2].text, hi[3].text = 'Disciplina', 'Real', 'Plan', 'TPI %'
                for d, m_col in [('Natacion', 'N_Mins'), ('Ciclismo', 'B_Mins'), ('Trote', 'R_Mins')]:
                    rc = ti.add_row().cells
                    rc[0].text = d
                    rc[1].text = to_hhmmss_display(row[m_col])
                    rc[2].text = f"{row[f'{d}_Hrs_Plan']:.1f}h"
                    rc[3].text = f"{row[f'TPI_{d}']:.1f}%"
                
                r_h = [row['N_Mins']/60, row['B_Mins']/60, row['R_Mins']/60]
                m_h = [row['Natacion_Hrs_Plan'], row['Ciclismo_Hrs_Plan'], row['Trote_Hrs_Plan']]
                g_buf = generar_grafico_comparativo(row['Deportista'], r_h, m_h)
                doc.add_picture(g_buf, width=Inches(4))
                doc.add_heading("Análisis Técnico", level=1)
                doc.add_paragraph(generar_comentario(row, 'General', 1))
                w_buf = io.BytesIO(); doc.save(w_buf)
                zf.writestr(f"Fichas/Ficha_{clean_string(row['Deportista'])}.docx", w_buf.getvalue())
    zip_buffer.seek(0)
    return zip_buffer

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
