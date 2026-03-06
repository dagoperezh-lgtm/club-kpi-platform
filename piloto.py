import streamlit as st
import requests
import pandas as pd

st.set_page_config(page_title="Piloto Strava Lite", layout="centered")

# =====================================================================
# 1. TUS CREDENCIALES (Reemplaza con tus datos reales)
# =====================================================================
CLIENT_ID = "162131"       # Ejemplo: "123456"
CLIENT_SECRET = "f827c4de29d7334330b43fdd04a99d900df566c2"   # Ejemplo: "abc123def..."

# =====================================================================
# 2. TU URL PÚBLICA DE STREAMLIT (Debe llevar https:// y terminar en /)
# =====================================================================
REDIRECT_URI = "metrikm-piloto.streamlit.app"  # <-- PON TU URL REAL AQUÍ

st.title("🚴‍♂️ Prueba Piloto: API de Strava")

# Leemos la barra de direcciones para ver si Strava ya nos mandó de vuelta
codigo_autorizacion = st.query_params.get("code")

if not codigo_autorizacion:
    st.markdown("### Paso 1: Conectar con Strava")
    st.info("Haz clic en el botón para autorizar la extracción de datos.")
    
    # Este es el link que le dice a Strava a dónde debe regresar
    auth_url = f"https://www.strava.com/oauth/authorize?client_id={CLIENT_ID}&response_type=code&redirect_uri={REDIRECT_URI}&scope=activity:read_all"
    
    st.link_button("🔗 Autorizar en Strava", auth_url)

else:
    st.success("✅ ¡Código de autorización recibido!")
    st.markdown("### Paso 2: Descargando tus entrenamientos...")
    
    # Canjeamos el código por la Llave Maestra
    url_token = "https://www.strava.com/oauth/token"
    payload = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "code": codigo_autorizacion,
        "grant_type": "authorization_code"
    }
    
    respuesta = requests.post(url_token, data=payload).json()
    
    if "access_token" in respuesta:
        access_token = respuesta["access_token"]
        atleta = respuesta["athlete"]["firstname"]
        st.write(f"👋 Hola, **{atleta}**. Conexión exitosa a los servidores de Strava.")
        
        # Pedimos las últimas 10 actividades
        url_actividades = "https://www.strava.com/api/v3/athlete/activities?per_page=10"
        headers = {"Authorization": f"Bearer {access_token}"}
        
        respuesta_actividades = requests.get(url_actividades, headers=headers).json()
        
        lista_limpia = []
        for act in respuesta_actividades:
            minutos = int(act['moving_time'] / 60)
            lista_limpia.append({
                "Fecha": act['start_date_local'].split("T")[0],
                "Deporte": act['type'],
                "Nombre": act['name'],
                "Minutos Netos": minutos,
                "Distancia (M)": act['distance']
            })
            
        st.dataframe(pd.DataFrame(lista_limpia), use_container_width=True)
        
        if st.button("🔄 Empezar de nuevo"):
            st.query_params.clear()
            st.rerun()
    else:
        st.error("Error al canjear el código con Strava.")
        st.json(respuesta)
