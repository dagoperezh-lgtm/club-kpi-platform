import streamlit as st
import requests

try:
    import gspread
    from google.oauth2.service_account import Credentials
except ImportError:
    gspread = None
    Credentials = None

st.set_page_config(page_title="Portal Atletas - Metri KM", page_icon="🏃‍♂️", layout="centered")

# =====================================================================
# 1. CREDENCIALES DE STRAVA
# =====================================================================
CLIENT_ID = "162131"
CLIENT_SECRET = st.secrets["STRAVA_CLIENT_SECRET"]

# ATENCIÓN: Esta URL la cambiaremos en el Paso 3, por ahora déjala así
REDIRECT_URI = "https://club-kpi-platform-cej3krxg3wgv3a9on4a4d6.streamlit.app/"

# =====================================================================
# 1.1 BÓVEDA PERSISTENTE (Google Sheets, en vez de un archivo local
# que se borra cada vez que la app se reinicia)
# =====================================================================
NOMBRE_HOJA_BOVEDA = "Tokens"

def get_google_client():
    if gspread is None:
        return None
    try:
        scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"⚠️ Error de conexión a Google Cloud: {e}")
        return None

def obtener_hoja_boveda():
    client = get_google_client()
    if client is None:
        st.error("⚠️ No se pudo conectar a la bóveda de datos (Google Sheets).")
        return None
    doc = client.open_by_url(st.secrets["google_sheet_url"])
    try:
        return doc.worksheet(NOMBRE_HOJA_BOVEDA)
    except gspread.exceptions.WorksheetNotFound:
        hoja = doc.add_worksheet(title=NOMBRE_HOJA_BOVEDA, rows=200, cols=4)
        hoja.append_row(["Atleta", "access_token", "refresh_token", "expires_at"])
        return hoja

def guardar_token_atleta(atleta_matchkey, access_token, refresh_token, expires_at):
    hoja = obtener_hoja_boveda()
    if hoja is None:
        return False
    fila_valores = [atleta_matchkey, access_token, refresh_token, str(expires_at)]
    celda = hoja.find(atleta_matchkey, in_column=1)
    if celda:
        hoja.update(f"A{celda.row}:D{celda.row}", [fila_valores])
    else:
        hoja.append_row(fila_valores)
    return True

# =====================================================================
# 2. DISEÑO DE CABECERA
# =====================================================================
col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    try:
        st.image("logo_metrikm.png", use_container_width=True)
    except:
        st.markdown("### Metri KM - TYM")

st.markdown("<h2 style='text-align: center;'>Portal de Sincronización</h2>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center;'>Vincula tu cuenta de Strava para automatizar tu reporte semanal del Club TYM.</p>", unsafe_allow_html=True)
st.divider()

# =====================================================================
# 3. LÓGICA DE CONEXIÓN
# =====================================================================
# Detectamos si el atleta viene rebotando desde Strava con éxito
codigo_auth = st.query_params.get("code")
atleta_matchkey = st.query_params.get("state")

if not codigo_auth:
    st.info("Paso 1: Identifícate en el sistema")
    # El atleta escribe su nombre tal cual aparece en tu Excel Maestro
    nombre_atleta = st.text_input("Ingresa tu Nombre y Apellido (Ej: Tomas Galmez)")
    
    if nombre_atleta:
        # En el enlace, pegamos su nombre en el parámetro "state" para no perderlo
        auth_url = f"https://www.strava.com/oauth/authorize?client_id={CLIENT_ID}&response_type=code&redirect_uri={REDIRECT_URI}&scope=activity:read_all&state={nombre_atleta}"
        st.link_button("🔗 Paso 2: Conectar mi Strava con Metri KM", auth_url, use_container_width=True)

else:
    st.warning(f"Procesando llaves de seguridad para: **{atleta_matchkey}**...")
    url_token = "https://www.strava.com/oauth/token"
    payload = {
        "client_id": CLIENT_ID, 
        "client_secret": CLIENT_SECRET,
        "code": codigo_auth, 
        "grant_type": "authorization_code"
    }
    res = requests.post(url_token, data=payload).json()
    
    if "access_token" in res:
        # --- GUARDAR LA LLAVE EN LA BÓVEDA PERMANENTE (Google Sheets) ---
        guardado_ok = guardar_token_atleta(
            atleta_matchkey,
            res["access_token"],
            res["refresh_token"],
            res["expires_at"]
        )

        if guardado_ok:
            st.success(f"✅ ¡Éxito, **{atleta_matchkey}**! Tus entrenamientos están sincronizados con el Club TYM.")
            st.balloons()

            if st.button("Sincronizar otro atleta"):
                st.query_params.clear()
                st.rerun()
        else:
            st.error("No se pudo guardar tu conexión en la bóveda. Intenta de nuevo.")
    else:
        st.error("Hubo un error de comunicación con Strava. Intenta de nuevo.")
