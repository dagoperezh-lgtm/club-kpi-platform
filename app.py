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
    
    nombre_limpio_temp = str(text).strip()
    nombre_limpio_temp = nombre_limpio_temp.upper()
    info_normalizada = unicodedata.normalize('NFKD', nombre_limpio_temp)
    
    resultado_final_nombre = ""
    for caracter_indiv in info_normalizada:
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
    lista_casos_nulos = ['--:--', '0', '', '00:00:00', '0:00:00', '00:00', '0.0', 'NC', '0:00']
    
    if string_valor in lista_casos_nulos:
        return 0
        
    try:
        if isinstance(valor_entrada_tiempo, (float, int)):
            return int(round(valor_entrada_tiempo * 1440))
        
        if isinstance(valor_entrada_tiempo, (time, datetime)):
            return (valor_entrada_tiempo.hour * 60) + valor_entrada_tiempo.minute
            
        try:
            conversion_float = float(string_valor)
            return int(round(conversion_float * 1440))
        except ValueError:
            pass
            
        if ':' in string_valor:
            bloques_tiempo = string_valor.split(':')
            if len(bloques_tiempo) >= 2:
                horas_bloque = int(bloques_tiempo[0])
                minutos_raw_bloque = bloques_tiempo[1]
                minutos_clean_bloque = int(minutos_raw_bloque.split('.')[0])
                return (horas_bloque * 60) + minutos_clean_bloque
        
        busqueda_horas = re.search(r'(\d+)h', string_valor)
        busqueda_minutos = re.search(r'(\d+)min', string_valor)
        
        h_resultado = int(busqueda_horas.group(1)) if busqueda_horas else 0
        m_resultado = int(busqueda_minutos.group(1)) if busqueda_minutos else 0
            
        return (h_resultado * 60) + m_resultado
    except Exception:
        return 0

def to_hhmmss_display(minutos_totales_input):
    """
    Formato de texto HH:MM:00 exclusivo para la estética del reporte Word.
    """
    valor_horas_v = int(minutos_totales_input // 60)
    valor_minutos_v = int(minutos_totales_input % 60)
    return f"{valor_horas_v:02d}:{valor_minutos_v:02d}:00"

# =============================================================================
# SECCIÓN 3: MOTOR NARRATIVO PRO CHILE (V2.2.28 - 75+ FRASES)
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
    atleta_actual = str(datos_de_fila.get('Deportista', 'Atleta TYM'))
    tiempo_actual = str(datos_de_fila.get(nombre_categoria, "00:00:00"))
    
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

    cat_key = 'General' if nombre_categoria in ['Completos', 'General'] else nombre_categoria
    if cat_key not in pools:
        return f"Desempeño consistente de {atleta_actual} en {nombre_categoria}."

    frase_plantilla = str(obtener_frase_base(cat_key, pools[cat_key]))
    comentario_final = frase_plantilla.replace("{atleta}", atleta_actual).replace("{tiempo}", tiempo_actual)
    
    if rank_posicion == 1 and cat_key == 'General':
        comentario_final = f"🏆 {comentario_final.replace(atleta_actual, f'nuestro líder {atleta_actual}')}"
    
    return comentario_final

# =============================================================================
# SECCIÓN 4: PARSERS DE ENTRADA (STRAVA & OCR)
# =============================================================================

def parse_raw_data(bloque_input_strava):
    lista_de_registros_atleta = []
    valor_rank_contador = 1
    bloque_input_strava = bloque_input_strava.replace('\xa0', ' ')
    lineas_encontradas = bloque_input_strava.strip().split('\n')
    
    for fila_texto in lineas_encontradas:
        if not fila_texto or 'Deportista' in fila_texto: continue
        try:
            patron_tiempos = r'(\d+h\s*\d*min|\d+h|\d+min|--:--|\d{2}:\d{2}:\d{2})'
            tiempos_en_linea = re.findall(patron_tiempos, fila_texto)
            if not tiempos_en_linea: continue
            
            string_del_total = tiempos_en_linea[0]
            ubicacion_del_tiempo = fila_texto.find(string_del_total)
            segmento_del_nombre = fila_texto[:ubicacion_del_tiempo].strip()
            nombre_limpio_final = re.sub(r'^\d+\s*', '', segmento_del_nombre).strip()
            
            minutos_volumen_total = to_mins(string_del_total)
            minutos_nat = to_mins(tiempos_en_linea[1]) if len(tiempos_en_linea) > 1 else 0
            minutos_bici = to_mins(tiempos_en_linea[2]) if len(tiempos_en_linea) > 2 else 0
            minutos_trote = to_mins(tiempos_en_linea[3]) if len(tiempos_en_linea) > 3 else 0
            
            lista_tiempos_cv = [minutos_nat, minutos_bici, minutos_trote]
            valor_cv_final = round(np.std(lista_tiempos_cv) / np.mean(lista_tiempos_cv), 4) if 0 not in lista_tiempos_cv else "NC"
            
            diccionario_de_atleta = {
                '#': valor_rank_contador,
                'Deportista': nombre_limpio_final,
                'Tiempo Total': to_hhmmss_display(minutos_volumen_total),
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
            valor_rank_contador += 1
        except Exception: continue
    return pd.DataFrame(lista_de_registros_atleta)

# =============================================================================
# SECCIÓN 5: MOTOR DE ACTUALIZACIÓN DEL MAESTRO (BLINDADO - NO SINTETIZAR)
# =============================================================================

def actualizar_maestro_tym(dict_dfs_originales, df_semana_actual, nombre_nueva_columna):
    """
    Actualiza el libro Excel completo preservando TODA la historia previa.
    Inmune a tildes y variaciones de nombre en las pestañas del Maestro.
    """
    dict_dfs_actualizados = {}
    df_semana_actual['MatchKey'] = df_semana_actual['Deportista'].apply(clean_string)
    hojas_en_maestro_norm = {clean_string(k): k for k in dict_dfs_originales.keys()}
    
    hojas_a_procesar = {
        'TIEMPO TOTAL': 'T_Mins',
        'NATACION': 'N_Mins',
        'BICICLETA': 'B_Mins',
        'TROTE': 'R_Mins',
        'CV': 'CV'
    }
    
    for key_norm, col_origen in hojas_a_procesar.items():
        orig_key = hojas_en_maestro_norm.get(key_norm)
        if orig_key:
            df_maestro_hoja = dict_dfs_originales[orig_key].copy()
            col_identidad = 'Nombre' if 'Nombre' in df_maestro_hoja.columns else \
                            ('Deportista' if 'Deportista' in df_maestro_hoja.columns else df_maestro_hoja.columns[0])
            
            df_maestro_hoja['MatchKey'] = df_maestro_hoja[col_identidad].apply(clean_string)
            df_novedad = df_semana_actual[['MatchKey', col_origen]].copy()
            
            if col_origen != 'CV':
                df_novedad[nombre_nueva_columna] = df_novedad[col_origen].apply(lambda x: x / 1440.0)
            else:
                df_novedad[nombre_nueva_columna] = df_novedad[col_origen]
            
            df_novedad = df_novedad.drop_duplicates(subset=['MatchKey'], keep='first')
            df_final_hoja = pd.merge(df_maestro_hoja, df_novedad[['MatchKey', nombre_nueva_columna]], on='MatchKey', how='outer')
            
            mask_nombre_vacio = df_final_hoja[col_identidad].isna()
            nombres_mapeo = df_semana_actual.set_index('MatchKey')['Deportista'].to_dict()
            df_final_hoja.loc[mask_nombre_vacio, col_identidad] = df_final_hoja.loc[mask_nombre_vacio, 'MatchKey'].map(nombres_mapeo)
            df_final_hoja[nombre_nueva_columna] = df_final_hoja[nombre_nueva_columna].fillna('NC' if col_origen == 'CV' else 0)
            
            dict_dfs_actualizados[orig_key] = df_final_hoja.drop(columns=['MatchKey'], errors='ignore')
        else:
            df_nueva = df_semana_actual[['Deportista', col_origen]].copy()
            nombre_hoja_crear = key_norm.capitalize()
            df_nueva.rename(columns={'Deportista': 'Nombre', col_origen: nombre_nueva_columna}, inplace=True)
            if col_origen != 'CV':
                df_nueva[nombre_nueva_columna] = df_nueva[nombre_nueva_columna].apply(lambda x: x / 1440.0)
            dict_dfs_actualizados[nombre_hoja_crear] = df_nueva

    for k in dict_dfs_originales:
        if clean_string(k) not in hojas_a_procesar:
            dict_dfs_actualizados[k] = dict_dfs_originales[k]
    return dict_dfs_actualizados

def save_maestro_to_excel(dict_dfs):
    output_binario = io.BytesIO()
    with pd.ExcelWriter(output_binario, engine='xlsxwriter') as writer:
        for nombre_hoja, df_contenido in dict_dfs.items():
            df_contenido.to_excel(writer, sheet_name=nombre_hoja, index=False)
    return output_binario.getvalue()

# =============================================================================
# SECCIÓN 6: GENERADOR DE ENTREGABLES (WORD, GRAFICOS Y ZIP)
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

def generar_entregables_finales(df, maestro_upd, tag_sem):
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED) as zf:
        # 1. Maestro
        ex_buf = io.BytesIO()
        with pd.ExcelWriter(ex_buf, engine='xlsxwriter') as writer:
            for k, v in maestro_upd.items(): v.to_excel(writer, sheet_name=k, index=False)
        zf.writestr(f"01_Estadisticas_{tag_sem}.xlsx", ex_buf.getvalue())
        
        # 2. Reporte Grupal
        doc_g = Document()
        doc_g.add_heading(f"Reporte Semanal Club TYM - {tag_sem}", 0)
        mins_tot = df['T_Mins'].sum()
        p = doc_g.add_paragraph()
        p.add_run(f"Total deportistas: {len(df)}\n").bold = True
        p.add_run(f"Tiempo Total Club: {to_hhmmss_display(mins_tot)}")
        
        doc_g.add_heading("🏆 TOP 15 ADHERENCIA GLOBAL", level=1)
        tabla = doc_g.add_table(rows=1, cols=3); tabla.style = 'Light Grid Accent 1'
        h = tabla.rows[0].cells
        h[0].text, h[1].text, h[2].text = 'Pos', 'Deportista', 'TPI %'
        
        df_top = df[df['Es_Completo']].sort_values('TPI_Global', ascending=False).head(15)
        for i, (_, r) in enumerate(df_top.iterrows(), 1):
            rc = tabla.add_row().cells
            rc[0].text, rc[1].text, rc[2].text = str(i), r['Deportista'], f"{r['TPI_Global']:.1f}%"
        
        buf_g = io.BytesIO(); doc_g.save(buf_g)
        zf.writestr(f"02_Reporte_General_{tag_sem}.docx", buf_g.getvalue())

        # 3. Fichas Individuales
        for _, row in df.iterrows():
            if row['T_Mins'] > 0:
                doc = Document()
                doc.add_heading(f"Reporte TYM: {row['Deportista']}", 0)
                doc.add_paragraph(f"TPI Global: {row['TPI_Global']:.1f}%")
                
                # Tabla Desglose
                ti = doc.add_table(rows=1, cols=4); ti.style = 'Table Grid'
                hi = ti.rows[0].cells
                hi[0].text, hi[1].text, hi[2].text, hi[3].text = 'Disciplina', 'Real', 'Plan', 'TPI'
                for d, m_col in [('Natacion', 'N_Mins'), ('Ciclismo', 'B_Mins'), ('Trote', 'R_Mins')]:
                    rc = ti.add_row().cells
                    rc[0].text = d
                    rc[1].text = to_hhmmss_display(row[m_col])
                    rc[2].text = f"{row[f'{d}_Hrs_Plan']:.1f}h"
                    rc[3].text = f"{row[f'TPI_{d}']:.1f}%"
                
                # Gráfico
                r_h = [row['N_Mins']/60, row['B_Mins']/60, row['R_Mins']/60]
                m_h = [row['Natacion_Hrs_Plan'], row['Ciclismo_Hrs_Plan'], row['Trote_Hrs_Plan']]
                g_buf = generar_grafico_comparativo(row['Deportista'], r_h, m_h)
                doc.add_picture(g_buf, width=Inches(4))
                
                doc.add_heading("Análisis Técnico", level=1)
                doc.add_paragraph(generar_comentario(row, 'General', 1))
                
                w_buf = io.BytesIO(); doc.save(w_buf)
                zf.writestr(f"Fichas/Ficha_{clean_string(row['Deportista'])}.docx", w_buf.getvalue())
    zip_buf.seek(0)
    return zip_buf

# =============================================================================
# SECCIÓN 7: INTERFAZ DE USUARIO Y ORQUESTACIÓN (ST)
# =============================================================================

if 'maestro_upd' not in st.session_state: st.session_state['maestro_upd'] = None
if 'df_final' not in st.session_state: st.session_state['df_final'] = None

with st.sidebar:
    f_maestro = st.file_uploader("Subir Maestro (.xlsx)", type=["xlsx"])
    st.divider()
    st.subheader("Plan Global")
    meta_g = {
        'N_H': st.number_input("Nat (Hrs)", 3.0), 'N_S': st.number_input("Nat (Ses)", 3),
        'B_H': st.number_input("Cic (Hrs)", 5.0), 'B_S': st.number_input("Cic (Ses)", 3),
        'T_H': st.number_input("Tro (Hrs)", 3.0), 'T_S': st.number_input("Tro (Ses)", 3)
    }

raw_strava = st.text_area("Datos Strava:", height=200)
tag_sem = st.text_input("Etiqueta Semana:", "Sem 08")

if st.button("🚀 PROCESAR JORNADA"):
    if f_maestro and raw_strava:
        df = parse_raw_data(raw_strava)
        
        # Cálculo de TPI Individual (Regla 4.3 corregida por Auditoría)
        def aplicar_tpi(row):
            # Natación
            vci_n = (row['N_Mins'] / (row['Natacion_Hrs_Plan']*60)) * 100 if row['Natacion_Hrs_Plan'] > 0 else 0
            # SEI (Disciplina): Si entrenó al menos una vez, cuenta como 1 sesión realizada frente al plan
            sei_n = (1 / row['Natacion_Ses_Plan']) * 100 if (row['N_Mins'] > 0 and row['Natacion_Ses_Plan'] > 0) else 0
            tpi_n = min((vci_n * 0.4) + (sei_n * 0.6), 110)
            
            # Ciclismo
            vci_c = (row['B_Mins'] / (row['Ciclismo_Hrs_Plan']*60)) * 100 if row['Ciclismo_Hrs_Plan'] > 0 else 0
            sei_c = (1 / row['Ciclismo_Ses_Plan']) * 100 if (row['B_Mins'] > 0 and row['Ciclismo_Ses_Plan'] > 0) else 0
            tpi_c = min((vci_c * 0.4) + (sei_c * 0.6), 110)
            
            # Trote
            vci_t = (row['R_Mins'] / (row['Trote_Hrs_Plan']*60)) * 100 if row['Trote_Hrs_Plan'] > 0 else 0
            sei_t = (1 / row['Trote_Ses_Plan']) * 100 if (row['R_Mins'] > 0 and row['Trote_Ses_Plan'] > 0) else 0
            tpi_t = min((vci_t * 0.4) + (sei_t * 0.6), 110)
            
            tpi_g = np.mean([tpi_n, tpi_c, tpi_t])
            es_comp = row['N_Mins'] > 0 and row['B_Mins'] > 0 and row['R_Mins'] > 0
            
            return pd.Series([tpi_n, tpi_c, tpi_t, tpi_g, es_comp])

        # Inyectar metas y aplicar cálculo
        for d, pref in [('Natacion', 'N'), ('Ciclismo', 'B'), ('Trote', 'R')]:
            df[f'{d}_Hrs_Plan'] = meta_g[f'{pref}_H']
            df[f'{d}_Ses_Plan'] = meta_g[f'{pref}_S']
            
        df[['TPI_Natacion', 'TPI_Ciclismo', 'TPI_Trote', 'TPI_Global', 'Es_Completo']] = df.apply(aplicar_tpi, axis=1)
        
        m_upd = actualizar_maestro_tym(pd.read_excel(f_maestro, sheet_name=None), df, tag_sem)
        zip_p = generar_entregables_finales(df, m_upd, tag_sem)
        
        st.session_state['maestro_upd'] = zip_p
        st.session_state['df_final'] = df
        st.success("✅ Procesamiento completado.")

if st.session_state['df_final'] is not None:
    st.download_button("📥 DESCARGAR PACK COMPLETO", st.session_state['maestro_upd'], f"Pack_{tag_sem}.zip")
