import streamlit as st
import requests
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="Piloto Strava Lite", layout="centered")

# --- 1. PON TUS CREDENCIALES AQUÍ ---
CLIENT_ID = "TU_CLIENT_ID" # Ejemplo: "123456" (Déjalo entre comillas)
CLIENT_SECRET = "TU_CLIENT_SECRET" # Ejemplo: "abc123def456..." (Entre comillas)

# --- 2. PON TU URL DE STREAMLIT AQUÍ ---
# ¡Importante! Debe tener https:// al principio y terminar con una barra /
REDIRECT_URI = "https://metrikm-piloto.streamlit.app/"

st.title("🚴‍♂️ Prueba Piloto: API de Strava")

# --- 3. LEER LA URL PARA BUSCAR EL CÓDIGO DE PERMISO ---
codigo_autorizacion = st.query_params.get("code")

# --- 4. BOTÓN DE CONEXIÓN ---
if not codigo_autorizacion:
    st.markdown("### Paso 1: Autorizar a Metri KM")
    st.info("Haz clic en el botón de abajo para ir a Strava y dar permiso para leer tus actividades.")
    
    # Este link es la "puerta web" hacia Strava
    auth_url = f"https://www.strava.com/oauth/authorize?client_id={CLIENT_ID}&response_type=code&redirect_uri={REDIRECT_URI}&scope=activity:read_all"
    
    st.link_button("🔗 Conectar con mi cuenta de Strava", auth_url)

# --- 5. EXTRACCIÓN DE DATOS ---
else:
    st.success("✅ ¡Autorización recibida por Strava!")
    st.markdown("### Paso 2: Extrayendo tus datos...")
    
    # Cambiamos el código temporal por la "Llave Maestra" (Access Token)
    url_token = "https://www.strava.com/oauth/token"
    payload = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "code": codigo_autorizacion,
        "grant_type": "authorization_code"
    }
    
    respuesta_token = requests.post(url_token, data=payload).json()
    
    if "access_token" in respuesta_token:
        access_token = respuesta_token["access_token"]
        atleta = respuesta_token["athlete"]["firstname"]
        
        st.write(f"👋 Hola, **{atleta}**. Llave maestra obtenida con éxito.")
        
        # Vamos a pedir las últimas 10 actividades
        url_actividades = "https://www.strava.com/api/v3/athlete/activities?per_page=10"
        headers = {"Authorization": f"Bearer {access_token}"}
        
        respuesta_actividades = requests.get(url_actividades, headers=headers).json()
        
        # Transformamos los datos crudos a una tabla bonita para que los leas
        lista_limpia = []
        for act in respuesta_actividades:
            fecha_corta = act['start_date_local'].split("T")[0]
            # Convertimos segundos a minutos netos
            minutos = int(act['moving_time'] / 60) 
            
            lista_limpia.append({
                "Fecha": fecha_corta,
                "Deporte": act['type'],
                "Nombre": act['name'],
                "Minutos Netos": minutos,
                "Distancia (Mts)": act['distance']
            })
            
        df_actividades = pd.DataFrame(lista_limpia)
        
        st.markdown("### 📊 Tus últimos 10 entrenamientos reales:")
        st.dataframe(df_actividades, use_container_width=True)
        
        st.success("¡Prueba Piloto Completada! Ya sabemos cómo leer los datos desde los servidores de Strava.")
        
        if st.button("Empezar de nuevo"):
            st.query_params.clear()
            st.rerun()
            
    else:
        st.error("Hubo un error al cambiar el código. Verifica tu Client ID y Secret.")
        st.write(respuesta_token)
