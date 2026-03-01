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

def time_to_seconds(time_str):
    """
    Convierte formatos:
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
# SECCIÓN 9 - FICHA INDIVIDUAL PDF (CON HISTÓRICO REAL)
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

selected_athlete = st.selectbox("Seleccionar atleta", df["Nombre"].unique())

if st.button("Generar PDF"):

    atleta_df = df[df["Nombre"] == selected_athlete].iloc[0]

    # ==========================
    # CARGAR HISTÓRICO
    # ==========================
    if historico_file:
        df_hist = pd.read_excel(historico_file)
        df_hist_atleta = df_hist[df_hist["Nombre"] == selected_athlete]
    else:
        df_hist_atleta = pd.DataFrame()

    # Excluir semana actual del histórico
    if not df_hist_atleta.empty:
        promedio_hist_total = df_hist_atleta["total_sec"].mean()
        promedio_hist_swim = df_hist_atleta["swim_sec"].mean()
        promedio_hist_bike = df_hist_atleta["bike_sec"].mean()
        promedio_hist_run = df_hist_atleta["run_sec"].mean()
    else:
        promedio_hist_total = 0
        promedio_hist_swim = 0
        promedio_hist_bike = 0
        promedio_hist_run = 0

    # Función comparación histórica
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

    # ==========================
    # CREAR PDF
    # ==========================
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    elements = []

    styles = getSampleStyleSheet()
    title_style = styles["Heading1"]

    elements.append(Paragraph(f"Ficha Individual - {selected_athlete}", title_style))
    elements.append(Spacer(1, 0.3 * inch))

    data_kpi = [
        ["Total Semana", seconds_to_hhmm(atleta_df["total_sec"])],
        ["Vs Histórico", comparar(atleta_df["total_sec"], promedio_hist_total)],
        ["Ranking Volumen", int(atleta_df["Rank_Volumen"])],
        ["VN (0-1)", atleta_df["VN"]],
    ]

    table_kpi = Table(data_kpi, colWidths=[220, 200])
    table_kpi.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey)
    ]))

    elements.append(table_kpi)
    elements.append(Spacer(1, 0.4 * inch))

    data_disc = [
        ["Disciplina", "Tiempo", "Vs Histórico"],
        ["Natación",
         seconds_to_hhmm(atleta_df["swim_sec"]),
         comparar(atleta_df["swim_sec"], promedio_hist_swim)],
        ["Ciclismo",
         seconds_to_hhmm(atleta_df["bike_sec"]),
         comparar(atleta_df["bike_sec"], promedio_hist_bike)],
        ["Trote",
         seconds_to_hhmm(atleta_df["run_sec"]),
         comparar(atleta_df["run_sec"], promedio_hist_run)],
    ]

    table_disc = Table(data_disc, colWidths=[120, 120, 180])
    table_disc.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey)
    ]))

    elements.append(table_disc)
    elements.append(Spacer(1, 0.4 * inch))

    # Gráfico barras verticales
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
