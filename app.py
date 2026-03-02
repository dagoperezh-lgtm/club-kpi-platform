import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile
import unicodedata
import matplotlib.pyplot as plt
import plotly.graph_objects as go
from datetime import time, datetime
from docx import Document
from docx.shared import Inches, Pt

# =============================================================================
# INICIO SECCIÓN 1: MOTOR ARITMÉTICO Y NORMALIZACIÓN (TimeEngine)
# Regla 7 (Aritmética) y Regla 8 (Tiempos > 24h)
# =============================================================================

def to_mins(valor):
    if pd.isna(valor) or str(valor).strip() in ['0', '', 'NC', '--:--']:
        return 0.0
    try:
        if isinstance(valor, (float, int)):
            return float(valor * 1440 if valor < 1 else valor)
        if isinstance(valor, (time, datetime)):
            return float(valor.hour * 60 + valor.minute)
        s = str(valor).strip()
        if ':' in s:
            partes = s.split(':')
            return float(int(partes[0]) * 60 + int(partes[1]))
    except:
        return 0.0
    return 0.0

def format_duracion_larga(minutos):
    if minutos is None or minutos < 0: return "00:00"
    tot_m = int(round(minutos))
    hrs = tot_m // 60
    mins = tot_m % 60
    return f"{hrs:02d}:{mins:02d}"

def normalizar_nombre(txt):
    if not txt: return ""
    norm = unicodedata.normalize('NFKD', str(txt).strip().upper())
    return "".join([c for c in norm if not unicodedata.combining(c)])

# =============================================================================
# INICIO SECCIÓN 2: PROCESAMIENTO DE IDENTIDADES (MatchEngine)
# Gestión de atletas nuevos y merge de archivos
# =============================================================================

def pipeline_identidad(df_m, df_s, df_p):
    df_s['MatchKey'] = df_s.iloc[:, 0].apply(normalizar_nombre)
    df_p['MatchKey'] = df_p.iloc[:, 0].apply(normalizar_nombre)
    df_m['MatchKey'] = df_m.iloc[:, 0].apply(normalizar_nombre)
    
    df_merged = pd.merge(df_s, df_p, on='MatchKey', how='left', suffixes=('', '_plan'))
    nuevos = set(df_s['MatchKey']) - set(df_m['MatchKey'])
    return df_merged, nuevos

# =============================================================================
# INICIO SECCIÓN 3: MOTOR DE KPIS Y ADHERENCIA (KPIDriven)
# Regla 4.3 (TPI) y Regla 4.5 (Filtro TOP 15)
# =============================================================================

def calcular_tpi_individual(row, dict_g):
    res = {}
    mapping = {'Natacion': 'Natación', 'Ciclismo': 'Bicicleta', 'Trote': 'Trote'}
    completo = True
    
    for d_nom, col_ex in mapping.items():
        h_plan = row.get(f'{d_nom}_Hrs_Plan', dict_g[f'{d_nom}_Hrs_Plan'])
        s_plan = row.get(f'{d_nom}_Ses_Plan', dict_g[f'{d_nom}_Ses_Plan'])
        real_m = to_mins(row.get(col_ex, 0.0))
        
        if real_m <= 0: completo = False
        
        vci = (real_m / (h_plan * 60)) * 100 if h_plan > 0 else 0.0
        sei = (100.0 if real_m > 0 else 0.0) / s_plan * 100 if s_plan > 0 else 0.0
        
        res[f'TPI_{d_nom}'] = (vci * 0.4) + (sei * 0.6)
        res[f'{d_nom}_Hrs_Plan'] = h_plan
        res[f'{d_nom}_Mins_Real'] = real_m

    res['TPI_Global'] = np.mean([res['TPI_Natacion'], res['TPI_Ciclismo'], res['TPI_Trote']])
    res['Es_Completo'] = completo
    return pd.Series(res)

# =============================================================================
# INICIO SECCIÓN 4: MOTOR DE VISUALIZACIÓN (VizEngine)
# Regla 3: Velocímetro
# =============================================================================

def generar_velocimetro_tpi(valor, nombre):
    fig = go.Figure(go.Indicator(
        mode = "gauge+number", value = valor,
        title = {'text': f"TPI Global: {nombre}", 'font': {'size': 20}},
        gauge = {'axis': {'range': [0, 100]},
                 'bar': {'color': "#1E90FF"},
                 'steps': [{'range': [0, 85], 'color': "lightgray"},
                           {'range': [85, 100], 'color': "#32CD32"}]}))
    fig.update_layout(height=300, margin=dict(l=20, r=20, t=50, b=20))
    return fig

# =============================================================================
# INICIO SECCIÓN 5: PERSISTENCIA Y ACTUALIZACIÓN (HistoryEngine)
# =============================================================================

def actualizar_historial_maestro(df_m_orig, df_procesado, tag_semana):
    """
    Crea el nuevo archivo Maestro.
    Garantiza que las semanas anteriores no se borren (Merge Outer).
    """
    # 1. Extraemos solo la identidad y el KPI de interés de la semana actual
    df_novedades = df_procesado[['MatchKey', 'TPI_Global']].copy()
    df_novedades.rename(columns={'TPI_Global': tag_semana}, inplace=True)
    
    # 2. Unión Atómica: Preservamos TODAS las columnas del maestro anterior
    # El 'how=outer' asegura que si hay atletas nuevos en df_procesado, se agreguen.
    df_nuevo_maestro = pd.merge(df_m_orig, df_novedades, on='MatchKey', how='outer')
    
    # 3. Limpieza de nulos: Atletas que no entrenaron esta semana quedan con 0
    df_nuevo_maestro[tag_semana] = df_nuevo_maestro[tag_semana].fillna(0)
    
    return df_nuevo_maestro

# =============================================================================
# FIN SECCIÓN 5
# =============================================================================

# =============================================================================
# INICIO SECCIÓN 6: ORQUESTADOR DE ENTREGABLES (ZipEngine)
# =============================================================================

def generar_pack_entregables(df_full, df_maestro_upd, tag_semana):
    """
    Construye el archivo ZIP con Fichas Word y Excel de Historial.
    Garantiza el encapsulamiento correcto de cada documento Word.
    """
    zip_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zf:
        
        # --- SUB-PROCESO A: EXCEL DE PROCESO ---
        buffer_excel = io.BytesIO()
        # Eliminamos MatchKey para el Excel final para mantener limpieza
        df_maestro_upd.drop(columns=['MatchKey'], errors='ignore').to_excel(buffer_excel, index=False)
        zf.writestr(f"01_Historial_Actualizado_{tag_semana}.xlsx", buffer_excel.getvalue())
        
        # --- SUB-PROCESO B: FICHAS INDIVIDUALES (CLIENTE) ---
        for _, row in df_full.iterrows():
            # Regla: Solo generar ficha si hubo minutos de actividad
            if (row['Natacion_Mins_Real'] + row['Ciclismo_Mins_Real'] + row['Trote_Mins_Real']) > 0:
                
                # 1. Instancia independiente por cada atleta
                doc_ficha = Document()
                doc_ficha.add_heading(f"Análisis de Desempeño: {row.iloc[0]}", 0)
                doc_ficha.add_paragraph(f"Reporte correspondiente a: {tag_semana}")
                
                # 2. Resumen TPI Global
                doc_ficha.add_heading("📊 Adherencia al Plan", level=1)
                p_resumen = doc_ficha.add_paragraph(f"Tu índice de cumplimiento global es de ")
                p_resumen.add_run(f"{row['TPI_Global']:.1f}%").bold = True
                
                # 3. Tabla de Disciplinas (Aritmética de Salida)
                tabla = doc_ficha.add_table(rows=1, cols=4)
                tabla.style = 'Light Grid Accent 1'
                encabezados = tabla.rows[0].cells
                encabezados[0].text = 'Disciplina'
                encabezados[1].text = 'Real (HH:MM)'
                encabezados[2].text = 'Meta (Hrs)'
                encabezados[3].text = 'TPI %'
                
                for disc in ['Natacion', 'Ciclismo', 'Trote']:
                    celdas = tabla.add_row().cells
                    celdas[0].text = disc
                    celdas[1].text = format_duracion_larga(row[f'{disc}_Mins_Real'])
                    celdas[2].text = f"{row[f'{disc}_Hrs_Plan']:.1f}"
                    celdas[3].text = f"{row[f'TPI_{disc}']:.1f}%"
                
                # 4. Inyección de Gráfico Comparativo (Sección 4)
                reales_h = [row['Natacion_Mins_Real']/60, row['Ciclismo_Mins_Real']/60, row['Trote_Mins_Real']/60]
                metas_h = [row['Natacion_Hrs_Plan'], row['Ciclismo_Hrs_Plan'], row['Trote_Hrs_Plan']]
                
                # Llamada al motor gráfico de la Sección 4
                buffer_graf = generar_grafico_comparativo(row.iloc[0], reales_h, metas_h)
                doc_ficha.add_paragraph("\n")
                doc_ficha.add_picture(buffer_graf, width=Inches(5))
                
                # 5. Comentarios Narrativos (Sección 3)
                doc_ficha.add_heading("📝 Comentario Técnico", level=1)
                doc_ficha.add_paragraph(generar_comentario(row, 'General', 0))
                
                # 6. Guardado en Buffer y Cierre de documento
                buffer_word = io.BytesIO()
                doc_ficha.save(buffer_word)
                zf.writestr(f"Fichas/Ficha_{normalizar_nombre(row.iloc[0])}.docx", buffer_word.getvalue())
                
    zip_buffer.seek(0)
    return zip_buffer

# =============================================================================
# FIN SECCIÓN 6
# =============================================================================
