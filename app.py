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
# CAJA 1: MOTOR ARITMÉTICO (TimeEngine)
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
    hrs = int(minutos // 60)
    mins = int(minutos % 60)
    return f"{hrs:02d}:{mins:02d}"

def normalizar_nombre(txt):
    if not txt: return ""
    norm = unicodedata.normalize('NFKD', str(txt).strip().upper())
    return "".join([c for c in norm if not unicodedata.combining(c)])

# =============================================================================
# CAJA 2: PROCESAMIENTO DE IDENTIDAD (MatchEngine)
# =============================================================================
def pipeline_identidad(df_m, df_s, df_p):
    df_s['MatchKey'] = df_s.iloc[:, 0].apply(normalizar_nombre)
    df_p['MatchKey'] = df_p.iloc[:, 0].apply(normalizar_nombre)
    df_m['MatchKey'] = df_m.iloc[:, 0].apply(normalizar_nombre)
    df_merged = pd.merge(df_s, df_p, on='MatchKey', how='left', suffixes=('', '_plan'))
    nuevos = set(df_s['MatchKey']) - set(df_m['MatchKey'])
    return df_merged, nuevos

# =============================================================================
# CAJA 3: MOTOR DE KPIS Y ADHERENCIA (KPIDriven)
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

def generar_comentario(row, cat, rank):
    return f"Desempeño sólido de {row.iloc[0]}. Su adherencia del {row['TPI_Global']:.1f}% refleja compromiso técnico."

# =============================================================================
# CAJA 4: VISUALIZACIÓN (VizEngine)
# =============================================================================
def generar_grafico_comparativo(nombre, reales, planes):
    labels = ['Natación', 'Ciclismo', 'Trote']
    x = np.arange(len(labels))
    fig, ax = plt.subplots(figsize=(5, 3))
    ax.bar(x - 0.2, planes, 0.4, label='Plan (h)', color='#BDC3C7')
    ax.bar(x + 0.2, reales, 0.4, label='Real (h)', color='#1E90FF')
    ax.set_xticks(x); ax.set_xticklabels(labels); ax.legend()
    buf = io.BytesIO(); plt.savefig(buf, format='png', bbox_inches='tight'); plt.close(fig)
    return buf

# =============================================================================
# CAJA 5 Y 6: ORQUESTACIÓN Y ENTREGABLES (AppMain)
# =============================================================================
st.set_page_config(page_title="Club KPI Platform", layout="wide")
st.title("📊 Club KPI Platform")

with st.sidebar:
    f_m = st.file_uploader("1. Maestro", type="xlsx")
    f_s = st.file_uploader("2. Semanal", type="xlsx")
    f_p = st.file_uploader("3. Plan", type="xlsx")
    dict_g = {
        'Natacion_Hrs_Plan': st.number_input("N Hrs", 3.0), 'Natacion_Ses_Plan': st.number_input("N Ses", 3),
        'Ciclismo_Hrs_Plan': st.number_input("C Hrs", 4.0), 'Ciclismo_Ses_Plan': st.number_input("C Ses", 3),
        'Trote_Hrs_Plan': st.number_input("T Hrs", 1.5), 'Trote_Ses_Plan': st.number_input("T Ses", 2)
    }

nombre_sem = st.text_input("Etiqueta Semana", "Sem 08")

if st.button("🚀 PROCESAR"):
    if f_m and f_s:
        df_m = pd.read_excel(f_m)
        df_s = pd.read_excel(f_s)
        df_p = pd.read_excel(f_p) if f_p else pd.DataFrame(columns=['Deportista'])
        
        df_merged, nuevos = pipeline_identidad(df_m, df_s, df_p)
        kpis = df_merged.apply(lambda r: calcular_tpi_individual(r, dict_g), axis=1)
        df_full = pd.concat([df_merged, kpis], axis=1)
        
        # TOP 15
        st.subheader(f"🏆 TOP 15 Adherencia - {nombre_sem}")
        top = df_full[df_full['Es_Completo']].sort_values('TPI_Global', ascending=False).head(15)
        st.table(top[['Deportista', 'TPI_Global']])

        # Generar ZIP
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED, False) as zf:
            # Excel Histórico (Outer Merge para preservar historia)
            df_hist = pd.merge(df_m, df_full[['MatchKey', 'TPI_Global']], on='MatchKey', how='outer')
            df_hist.rename(columns={'TPI_Global': nombre_sem}, inplace=True)
            ex_buf = io.BytesIO()
            df_hist.drop(columns=['MatchKey'], errors='ignore').to_excel(ex_buf, index=False)
            zf.writestr(f"Historial_Actualizado_{nombre_sem}.xlsx", ex_buf.getvalue())

            # Fichas Word
            for _, row in df_full.iterrows():
                if (row['Natacion_Mins_Real'] + row['Ciclismo_Mins_Real'] + row['Trote_Mins_Real']) > 0:
                    doc = Document()
                    doc.add_heading(f"Reporte: {row.iloc[0]}", 0)
                    doc.add_paragraph(f"TPI Global: {row['TPI_Global']:.1f}%")
                    
                    # Gráfico
                    r_h = [row['Natacion_Mins_Real']/60, row['Ciclismo_Mins_Real']/60, row['Trote_Mins_Real']/60]
                    p_h = [row['Natacion_Hrs_Plan'], row['Ciclismo_Hrs_Plan'], row['Trote_Hrs_Plan']]
                    g_buf = generar_grafico_comparativo(row.iloc[0], r_h, p_h)
                    doc.add_picture(g_buf, width=Inches(4))
                    
                    doc.add_paragraph(generar_comentario(row, 'General', 0))
                    w_buf = io.BytesIO(); doc.save(w_buf)
                    zf.writestr(f"Fichas/Ficha_{normalizar_nombre(row.iloc[0])}.docx", w_buf.getvalue())
        
        st.download_button("⬇️ Descargar Pack", zip_buf.getvalue(), f"Pack_{nombre_sem}.zip")
