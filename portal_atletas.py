import streamlit as st
import requests
import json
import os

st.set_page_config(page_title="Portal Atletas - Metri KM", page_icon="🏃‍♂️", layout="centered")

# =====================================================================
# 1. CREDENCIALES DE STRAVA
# =====================================================================
CLIENT_ID = "162131"
CLIENT_SECRET = "TU_CLIENT_SECRET" # <-- ¡PON TU SECRET REAL AQUÍ!

# ATENCIÓN: Esta URL la cambiaremos en el Paso 3, por ahora déjala así
REDIRECT_URI = "https://tu-nueva-url-del-portal.streamlit.app/" 
ARCHIVO_BOVEDA = "boveda_strava.json"

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
        # --- ABRIR LA BÓVEDA Y GUARDAR LA LLAVE ---
        boveda = {}
        # Si la bóveda ya existe, la abrimos para no borrar a los otros atletas
        if os.path.exists(ARCHIVO_BOVEDA):
            with open(ARCHIVO_BOVEDA, "r") as f:
                try:
                    boveda = json.load(f)
                except:
                    boveda = {}
                
        # Guardamos la información específica de este atleta usando su nombre como llave
        boveda[atleta_matchkey] = {
            "access_token": res["access_token"],
            "refresh_token": res["refresh_token"],
            "expires_at": res["expires_at"]
        }
        
        # Cerramos la bóveda con candado
        with open(ARCHIVO_BOVEDA, "w") as f:
            json.dump(boveda, f)
            
        st.success(f"✅ ¡Éxito, **{atleta_matchkey}**! Tus entrenamientos están sincronizados con el Club TYM.")
        st.balloons()
        
        if st.button("Sincronizar otro atleta"):
            st.query_params.clear()
            st.rerun()
    else:
        st.error("Hubo un error de comunicación con Strava. Intenta de nuevo.")
