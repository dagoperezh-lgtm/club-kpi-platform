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
