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
# INICIO SECCIÓN 5: GENERADOR DE GRÁFICOS PARA REPORTES (ReportEngine)
# Objetivo: Regla 1 (Gráfico Real vs Plan)
# Inputs: Datos calculados / Output: Buffer de imagen PNG
# =============================================================================

def generar_grafico_comparativo(nombre, reales, planes):
    """Genera gráfico de barras para el Word."""
    labels = ['Natación', 'Ciclismo', 'Trote']
    x = np.arange(len(labels))
    width = 0.35
    
    fig, ax = plt.subplots(figsize=(6, 4))
    ax.bar(x - width/2, planes, width, label='Plan (Hrs)', color='#BDC3C7')
    ax.bar(x + width/2, reales, width, label='Real (Hrs)', color='#1E90FF')
    
    ax.set_ylabel('Horas')
    ax.set_title(f'Cumplimiento Semanal: {nombre}')
    ax.set_xticks(x)
    ax.set_xticklabels(labels)
    ax.legend()
    
    buf = io.BytesIO()
    plt.savefig(buf, format='png', bbox_inches='tight')
    plt.close(fig)
    return buf

# =============================================================================
# INICIO SECCIÓN 6: INTERFAZ Y ORQUESTACIÓN (AppMain)
# Objetivo: Unir todas las cajas y generar entregables (Regla 6)
# =============================================================================

st.set_page_config(page_title="Club KPI Platform", layout="wide")
st.title("📊 Club KPI Platform - Modo Desarrollo")

if 'ok' not in st.session_state: st.session_state['ok'] = False

with st.sidebar:
    st.header("📥 Carga de Archivos")
    f_maestro = st.file_uploader("1. Excel Maestro (Histórico)", type="xlsx")
    f_semanal = st.file_uploader("2. Excel Semanal (Reales)", type="xlsx")
    f_plan = st.file_uploader("3. Excel Plan (Individual)", type="xlsx")
    
    st.divider()
    st.header("🌍 Plan Global (Backup)")
    gn_h = st.number_input("Natación (Hrs)", 3.0)
    gn_s = st.number_input("Natación (Ses)", 3)
    gb_h = st.number_input("Ciclismo (Hrs)", 4.0)
    gb_s = st.number_input("Ciclismo (Ses)", 3)
    gt_h = st.number_input("Trote (Hrs)", 1.5)
    gt_s = st.number_input("Trote (Ses)", 2)
    
    dict_g = {
        'Natacion_Hrs_Plan': gn_h, 'Natacion_Ses_Plan': gn_s,
        'Ciclismo_Hrs_Plan': gb_h, 'Ciclismo_Ses_Plan': gb_s,
        'Trote_Hrs_Plan': gt_h, 'Trote_Ses_Plan': gt_s
    }

nombre_sem = st.text_input("Etiqueta de esta semana (ej: Sem 09):", "Sem 08")

if st.button("🚀 PROCESAR JORNADA"):
    if f_maestro and f_semanal:
        df_m = pd.read_excel(f_maestro)
        df_s = pd.read_excel(f_semanal)
        df_p = pd.read_excel(f_plan) if f_plan else pd.DataFrame(columns=['Deportista'])
        
        # Orquestación de cajas
        df_merged, nuevos = pipeline_identidad(df_m, df_s, df_p)
        kpis = df_merged.apply(lambda r: calcular_tpi_individual(r, dict_g), axis=1)
        df_full = pd.concat([df_merged, kpis], axis=1)
        
        st.session_state['df_full'] = df_full
        st.session_state['df_m'] = df_m
        st.session_state['ok'] = True
    else:
        st.error("Faltan archivos obligatorios (Maestro y Semanal).")

if st.session_state['ok']:
    df_f = st.session_state['df_full']
    
    # Visualización TOP 15 (Regla 4.5)
    st.subheader(f"🏆 TOP 15 Adherencia Global - {nombre_sem}")
    top_15 = df_f[df_f['Es_Completo']].sort_values('TPI_Global', ascending=False).head(15)
    st.table(top_15[['Deportista', 'TPI_Global', 'TPI_Natacion', 'TPI_Ciclismo', 'TPI_Trote']])

    if st.button("📦 GENERAR ZIP DE ENTREGABLES"):
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "a") as zf:
            # 1. Excel Histórico Actualizado (Regla 6.2)
            df_hist = pd.merge(st.session_state['df_m'], df_f[['MatchKey', 'TPI_Global']], on='MatchKey', how='outer')
            df_hist.rename(columns={'TPI_Global': nombre_sem}, inplace=True)
            
            ex_buf = io.BytesIO()
            df_hist.drop(columns=['MatchKey']).to_excel(ex_buf, index=False)
            zf.writestr(f"Historial_Actualizado_{nombre_sem}.xlsx", ex_buf.getvalue())
            
            # 2. Fichas Individuales
            for _, row in df_f.iterrows():
                # Solo si tuvo actividad (Regla triatletas nuevos sin actividad)
                if (row['Natacion_Mins_Real'] + row['Ciclismo_Mins_Real'] + row['Trote_Mins_Real']) > 0:
                    doc = Document()
                    doc.add_heading(f"Reporte KPI: {row.iloc[0]}", 0)
                    doc.add_paragraph(f"Cumplimiento TPI: {row['TPI_Global']:.1f}%")
                    
                    # Tabla
                    tbl = doc.add_table(rows=1, cols=4); tbl.style = 'Light Grid Accent 1'
                    h_cells = tbl.rows[0].cells
                    h_cells[0].text, h_cells[1].text = 'Disciplina', 'Real'
                    h_cells[2].text, h_cells[3].text = 'Plan (Hrs)', 'TPI'
                    
                    for d in ['Natacion', 'Ciclismo', 'Trote']:
                        rc = tbl.add_row().cells
                        rc[0].text = d
                        rc[1].text = format_duracion_larga(row[f'{d}_Mins_Real'])
                        rc[2].text = str(row[f'{d}_Hrs_Plan'])
                        rc[3].text = f"{row[f'TPI_{d}']:.1f}%"
                    
                    # Gráfico
                    reales = [row['Natacion_Mins_Real']/60, row['Ciclismo_Mins_Real']/60, row['Trote_Mins_Real']/60]
                    planes = [row['Natacion_Hrs_Plan'], row['Ciclismo_Hrs_Plan'], row['Trote_Hrs_Plan']]
                    g_buf = generar_grafico_comparativo(row.iloc[0], reales, planes)
                    doc.add_picture(g_buf, width=Inches(5))
                    
                    d_buf = io.BytesIO(); doc.save(d_buf)
                    zf.writestr(f"Fichas/Ficha_{row.iloc[0]}.docx", d_buf.getvalue())
        
        st.download_button("⬇️ Descargar ZIP", zip_buf.getvalue(), f"Pack_{nombre_sem}.zip")
