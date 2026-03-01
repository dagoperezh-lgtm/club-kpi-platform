# ==========================================================
# CLUB TYМ KPI PLATFORM - MVP VOLUMEN v1
# Dagoberto Pérez
# ==========================================================

import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO

st.set_page_config(page_title="Club KPI Platform", layout="wide")

st.title("🏊🚴🏃 Club KPI Platform - MVP Volumen")



# ==========================================================
# SECCIÓN 1 - FUNCIONES UTILITARIAS (EDITABLE)
# ==========================================================

import unicodedata

def time_to_seconds(time_str):
    """
    Convierte:
    - 7h 39min
    - 38min
    - 00:33:44
    - --:--
    a segundos
    """
    if pd.isna(time_str):
        return 0

    time_str = str(time_str).strip()

    if time_str in ["--:--", "", "nan"]:
        return 0

    # Formato HH:MM:SS
    if ":" in time_str and "h" not in time_str:
        parts = time_str.split(":")
        if len(parts) == 3:
            h, m, s = parts
            return int(h)*3600 + int(m)*60 + int(s)

    hours = 0
    minutes = 0

    h_match = re.search(r"(\d+)h", time_str)
    m_match = re.search(r"(\d+)min", time_str)

    if h_match:
        hours = int(h_match.group(1))
    if m_match:
        minutes = int(m_match.group(1))

    return hours*3600 + minutes*60


def normalize_column(series):
    max_val = series.max()
    if max_val == 0:
        return series
    return series / max_val


def normalizar_nombre(nombre):
    """
    Quita tildes y normaliza para comparación interna
    """
    if pd.isna(nombre):
        return ""

    nombre = str(nombre).strip().lower()
    nombre = unicodedata.normalize('NFD', nombre)
    nombre = ''.join(c for c in nombre if unicodedata.category(c) != 'Mn')

    return nombre

# ==========================================================
# SECCIÓN 2 - CARGA DE ARCHIVOS
# ==========================================================

st.sidebar.header("📂 Carga de Archivos")

historico_file = st.sidebar.file_uploader("Subir Histórico (opcional)", type=["xlsx"])
semana_file = st.sidebar.file_uploader("Subir Semana RAW", type=["xlsx", "csv"])

if semana_file is None:
    st.warning("Sube el archivo de la semana para continuar.")
    st.stop()

# ==========================================================
# SECCIÓN 3 - PROCESAMIENTO SEMANA
# ==========================================================

if semana_file.name.endswith(".csv"):
    df_raw = pd.read_csv(semana_file)
else:
    df_raw = pd.read_excel(semana_file)

# Asegurar nombres estándar
df = df_raw.copy()

# Convertir tiempos a segundos
df["swim_sec"] = df["Natacion"].apply(time_to_seconds)
df["bike_sec"] = df["Ciclismo"].apply(time_to_seconds)
df["run_sec"] = df["Trote"].apply(time_to_seconds)
df["total_sec"] = df["swim_sec"] + df["bike_sec"] + df["run_sec"]

# ==========================================================
# SECCIÓN 4 - KPI VOLUMEN
# ==========================================================

df["VN"] = normalize_column(df["total_sec"])
df["VN"] = df["VN"].round(2)

# Ranking Volumen
df["Rank_Volumen"] = df["VN"].rank(ascending=False, method="min")

# ==========================================================
# SECCIÓN 4B - KPI ADHERENCIA (CONVERSIÓN AUTOMÁTICA)
# ==========================================================

plan_file = st.sidebar.file_uploader("Subir Plan Global", type=["xlsx"])

def clasificar_disciplina(texto):
    if pd.isna(texto):
        return None
    
    texto = str(texto).lower()
    
    if "trote" in texto or "run" in texto:
        return "Trote"
    elif "rodillo" in texto or "cicl" in texto or "bike" in texto:
        return "Ciclismo"
    elif "nat" in texto or "nado" in texto:
        return "Natación"
    elif "descanso" in texto:
        return None
    else:
        return None

if plan_file:

    plan_original = pd.read_excel(plan_file)

    # Detectar filas clave
    fila_actividad = plan_original[
        plan_original.iloc[:,0].astype(str).str.contains("Actividad", case=False, na=False)
    ].index[0]

    fila_duracion = plan_original[
        plan_original.iloc[:,0].astype(str).str.contains("Duración", case=False, na=False)
    ].index[0]

    actividades = plan_original.iloc[fila_actividad, 1:8]
    duraciones = plan_original.iloc[fila_duracion, 1:8]

    dias = ["Lunes","Martes","Miércoles","Jueves","Viernes","Sábado","Domingo"]

    plan_convertido = {
        "Natación": [0]*7,
        "Ciclismo": [0]*7,
        "Trote": [0]*7
    }

    # Clasificación automática
    for i in range(7):
        disciplina = clasificar_disciplina(actividades.iloc[i])
        tiempo_sec = time_to_seconds(duraciones.iloc[i])
        
        if disciplina in plan_convertido:
            plan_convertido[disciplina][i] += tiempo_sec

    # 🔵 AGREGAR 3H NATACIÓN OBLIGATORIAS
    horas_extra_natacion = 3 * 3600
    sesiones_natacion = 3
    tiempo_por_sesion = horas_extra_natacion // sesiones_natacion
    dias_para_natacion = [0, 2, 4]  # Lunes, Miércoles, Viernes

    for d in dias_para_natacion:
        plan_convertido["Natación"][d] += tiempo_por_sesion

    df_plan_global = pd.DataFrame(plan_convertido, index=dias).T

    plan_total_sec = df_plan_global.sum().sum()

    # ==========================
    # CÁLCULO ADHERENCIA
    # ==========================

    def calcular_adherencia(real_total, plan_total):
        if plan_total == 0:
            return np.nan
        valor = real_total / plan_total
        return round(min(valor, 1.10), 2)

    df["Adherencia"] = df["total_sec"].apply(
        lambda x: calcular_adherencia(x, plan_total_sec)
    )

    df["Rank_Adherencia"] = df["Adherencia"].rank(
        ascending=False, method="min"
    )

    st.subheader("📋 Plan Global Convertido (+3h Natación)")
    st.dataframe(df_plan_global)
    
# ==========================================================
# SECCIÓN 5 - RANKING POR DISCIPLINA
# ==========================================================

df["Swim_Index"] = normalize_column(df["swim_sec"]).round(2)
df["Bike_Index"] = normalize_column(df["bike_sec"]).round(2)
df["Run_Index"] = normalize_column(df["run_sec"]).round(2)

# Rankings individuales
df["Rank_Swim"] = df["swim_sec"].rank(ascending=False, method="min")
df["Rank_Bike"] = df["bike_sec"].rank(ascending=False, method="min")
df["Rank_Run"] = df["run_sec"].rank(ascending=False, method="min")

# ==========================================================
# SECCIÓN 6 - COMPARACIÓN VS PROMEDIO EQUIPO
# ==========================================================

team_avg = df["total_sec"].mean()
df["%_vs_Team_Avg"] = (df["total_sec"] / team_avg).round(2)

# ==========================================================
# SECCIÓN 7 - MOSTRAR RESULTADOS
# ==========================================================

st.subheader("🔵 Ranking Volumen Total")
st.dataframe(df.sort_values("Rank_Volumen"))

st.subheader("🏊 Ranking Natación")
st.dataframe(df[df["swim_sec"] > 0].sort_values("Rank_Swim"))

st.subheader("🚴 Ranking Ciclismo")
st.dataframe(df[df["bike_sec"] > 0].sort_values("Rank_Bike"))

st.subheader("🏃 Ranking Trote")
st.dataframe(df[df["run_sec"] > 0].sort_values("Rank_Run"))

# ==========================================================
# SECCIÓN 8 - ACTUALIZAR HISTÓRICO (VERSIÓN EXTENDIDA)
# ==========================================================

st.subheader("📚 Gestión de Histórico")

# Cargar histórico si existe
if historico_file:
    df_hist = pd.read_excel(historico_file)
else:
    df_hist = pd.DataFrame()

semana_label = st.text_input("Nombre Semana (ej: 2026-03-01)")

if st.button("Actualizar Histórico"):

    if semana_label == "":
        st.warning("Debes ingresar un nombre de semana.")
        st.stop()

    # Exportar todas las métricas necesarias
    df_export = df[[
        "Nombre",
        "total_sec",
        "swim_sec",
        "bike_sec",
        "run_sec",
        "VN"
    ]].copy()

    df_export["Semana"] = semana_label

    # Concatenar histórico
    if df_hist.empty:
        df_final = df_export
    else:
        df_final = pd.concat([df_hist, df_export], ignore_index=True)

    # Exportar archivo actualizado
    output = BytesIO()
    df_final.to_excel(output, index=False)
    output.seek(0)

    st.success("Histórico actualizado correctamente.")

    st.download_button(
        label="⬇ Descargar Histórico Actualizado",
        data=output,
        file_name="historico_actualizado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ==========================================================
# SECCIÓN 9 - FICHA INDIVIDUAL PDF (NORMALIZADA)
# ==========================================================

from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
import matplotlib.pyplot as plt
import tempfile

def seconds_to_hhmm(seconds):
    h = int(seconds // 3600)
    m = int((seconds % 3600) // 60)
    return f"{h}h {m}min"

st.subheader("📄 Generar Ficha Individual")

# Crear columna normalizada temporal
df["Nombre_norm"] = df["Nombre"].apply(normalizar_nombre)

selected_athlete = st.selectbox("Seleccionar atleta", df["Nombre"].unique())

if st.button("Generar PDF"):

    atleta_norm = normalizar_nombre(selected_athlete)
    atleta_df = df[df["Nombre_norm"] == atleta_norm].iloc[0]

  # ==========================================================
# SECCIÓN 9 - FICHA INDIVIDUAL PDF (ROBUSTA UTF-8 + HISTÓRICO)
# ==========================================================

from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
import matplotlib.pyplot as plt
import tempfile

def seconds_to_hhmm(seconds):
    h = int(seconds // 3600)
    m = int((seconds % 3600) // 60)
    return f"{h}h {m}min"

def limpiar_texto_mal_codificado(texto):
    try:
        return texto.encode('latin1').decode('utf-8')
    except:
        return texto

st.subheader("📄 Generar Ficha Individual")

# Corregir posibles errores de codificación en nombres
df["Nombre"] = df["Nombre"].astype(str).apply(limpiar_texto_mal_codificado)

selected_athlete = st.selectbox("Seleccionar atleta", df["Nombre"].unique())

if st.button("Generar PDF"):

    atleta_df = df[df["Nombre"] == selected_athlete].iloc[0]

    # ==========================
    # CARGAR HISTÓRICO DE FORMA SEGURA
    # ==========================
    if historico_file:
        df_hist = pd.read_excel(historico_file)

        # Limpiar nombres columnas
        df_hist.columns = df_hist.columns.str.strip()

        # Verificar si existe columna nombre (case insensitive)
        nombre_col = None
        for col in df_hist.columns:
            if col.lower() == "nombre":
                nombre_col = col

        if nombre_col:
            df_hist[nombre_col] = df_hist[nombre_col].astype(str).apply(limpiar_texto_mal_codificado)
            df_hist_atleta = df_hist[df_hist[nombre_col] == selected_athlete]
        else:
            df_hist_atleta = pd.DataFrame()
    else:
        df_hist_atleta = pd.DataFrame()

    if not df_hist_atleta.empty and "total_sec" in df_hist_atleta.columns:
        promedio_hist_total = df_hist_atleta["total_sec"].mean()
    else:
        promedio_hist_total = 0

    def comparar(actual, promedio):
        if promedio == 0:
            return "Sin histórico previo"
        diff = actual - promedio
        if diff > 0:
            return f"{seconds_to_hhmm(abs(diff))} MÁS"
        elif diff < 0:
            return f"{seconds_to_hhmm(abs(diff))} MENOS"
        else:
            return "Igual a tu promedio histórico"

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    elements = []

    styles = getSampleStyleSheet()
    elements.append(Paragraph(f"Ficha Individual - {selected_athlete}", styles["Heading1"]))
    elements.append(Spacer(1, 0.3 * inch))

    data_kpi = [
        ["Total Semana", seconds_to_hhmm(atleta_df["total_sec"])],
        ["Vs Histórico", comparar(atleta_df["total_sec"], promedio_hist_total)],
        ["Ranking Volumen", int(atleta_df["Rank_Volumen"])],
        ["Adherencia", atleta_df.get("Adherencia", "N/A")],
    ]

    table_kpi = Table(data_kpi, colWidths=[220, 200])
    table_kpi.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey)
    ]))

    elements.append(table_kpi)
    elements.append(Spacer(1, 0.4 * inch))

    fig, ax = plt.subplots()
    disciplinas = ["Swim", "Bike", "Run"]
    valores = [
        atleta_df["swim_sec"]/3600,
        atleta_df["bike_sec"]/3600,
        atleta_df["run_sec"]/3600
    ]
    ax.bar(disciplinas, valores)
    ax.set_ylabel("Horas")
    ax.set_title("Distribución por Disciplina")

    temp_img = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
    plt.savefig(temp_img.name)
    plt.close(fig)

    elements.append(Image(temp_img.name, width=4*inch, height=3*inch))

    doc.build(elements)
    buffer.seek(0)

    st.download_button(
        label="⬇ Descargar Ficha PDF",
        data=buffer,
        file_name=f"Ficha_{selected_athlete}.pdf",
        mime="application/pdf"
    )
