import streamlit as st
import pandas as pd
from datetime import datetime
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.client_context import ClientContext

# Configuración inicial de la página (Modo ancho para mejor visualización en móvil/PC)
st.set_page_config(page_title="Control de Asistencia - KCO Group", page_icon="⏱️", layout="wide")

# ==========================================
# 1. AUTENTICACIÓN Y CONEXIÓN SEGURA A SHAREPOINT
# ==========================================
@st.cache_resource
def init_sharepoint_connection():
    site_url = st.secrets["sharepoint"]["site_url"]
    client_id = st.secrets["sharepoint"]["client_id"]
    client_secret = st.secrets["sharepoint"]["client_secret"]
    
    ctx_auth = ClientContext(site_url).with_client_credentials(client_id, client_secret)
    return ctx_auth

try:
    ctx = init_sharepoint_connection()
except Exception as e:
    st.error(f"Error de conexión con SharePoint: {e}")
    st.stop()

# ==========================================
# 2. FUNCIONES DE CARGA DE DATOS EN VIVO
# ==========================================
@st.cache_data(ttl=60) # Actualiza los datos cada 60 segundos automáticamente
def load_data():
    # Cargar colaboradores vigentes
    collab_list = ctx.web.lists.get_by_title("dimColaboradores")
    collab_items = collab_list.items.get().execute_query()
    df_collab = pd.DataFrame([item.properties for item in collab_items])
    
    # Cargar marcaciones (transacciones)
    marc_list = ctx.web.lists.get_by_title("factMarcaciones")
    marc_items = marc_list.items.get().execute_query()
    df_marc = pd.DataFrame([item.properties for item in marc_items])
    
    return df_collab, df_marc

df_colaboradores, df_marcaciones = load_data()

# ==========================================
# 3. INTERFAZ VISUAL DEL DASHBOARD
# ==========================================
st.title("⏱️ KCO Group - Control de Asistencia en Tiempo Real")
st.markdown("Monitoreo operativo de asistencia, horas acumuladas y faltas semanales.")

if df_colaboradores.empty or df_marcaciones.empty:
    st.warning("Aún no hay suficientes registros en las listas de SharePoint.")
else:
    # Limpieza básica de datos y conversión de fechas
    df_marcaciones['FechaHoraMarcacion'] = pd.to_datetime(df_marcaciones['FechaHoraMarcacion'])
    df_marcaciones['FechaSolo'] = df_marcaciones['FechaHoraMarcacion'].dt.date
    
    # Filtro lateral por semana o colaborador
    st.sidebar.header("Filtros de Control")
    colaborador_filtro = st.sidebar.selectbox("Seleccionar Colaborador", ["Todos"] + list(df_colaboradores['NombreCompleto'].unique()))
    
    # Métricas generales rápidas (KPIs)
    col1, col2, col3 = st.columns(3)
    col1.metric("Colaboradores Activos", len(df_colaboradores[df_colaboradores['EsVigente'] == 'Sí']))
    col2.metric("Total Marcaciones Registradas", len(df_marcaciones))
    col3.metric("Última Sincronización", datetime.now().strftime("%d/%m/%Y %H:%M"))

    st.divider()
    
    # Tabla de transacciones en tiempo real
    st.subheader("📋 Registro Detallado de Marcaciones")
    
    # Cruce de DNI con Nombre
    df_merged = pd.merge(df_marcaciones, df_colaboradores[['DNI', 'NombreCompleto']], on='DNI', how='left')
    
    if colaborador_filtro != "Todos":
        df_merged = df_merged[df_merged['NombreCompleto'] == colaborador_filtro]
        
    st.dataframe(df_merged[['NombreCompleto', 'DNI', 'TipoMarcacion', 'FechaHoraMarcacion']], use_container_width=True)