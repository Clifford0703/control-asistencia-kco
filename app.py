import msal
import requests
import streamlit as st

# Cargar secretos desde st.secrets
TENANT_ID = st.secrets["sharepoint"]["tenant_id"]  # O tu ID de directorio (tenant)
CLIENT_ID = st.secrets["sharepoint"]["client_id"]
CLIENT_SECRET = st.secrets["sharepoint"]["client_secret"]
SITE_URL = st.secrets["sharepoint"]["site_url"]  # Ej: "https://tuempresa.sharepoint.com/sites/Tusitio"


def get_access_token():
  authority = f"https://login.microsoftonline.com/{TENANT_ID}"
  app = msal.ConfidentialClientApplication(
      CLIENT_ID,
      authority=authority,
      client_credential=CLIENT_SECRET,
  )
  result = app.acquire_token_for_client(
      scopes=["https://graph.microsoft.com/.default"]
  )
  if "access_token" in result:
    return result["access_token"]
  else:
    raise Exception(
        f"Error al obtener token de acceso: {result.get('error_description')}"
    )


def load_data():
  token = get_access_token()
  headers = {
      "Authorization": f"Bearer {token}",
      "Accept": "application/json",
  }

  # Ejemplo para consultar una lista de SharePoint mediante Microsoft Graph API
  # 1. Obtener el ID del sitio web de SharePoint
  site_domain = SITE_URL.replace("https://", "").split("/")[0]
  site_path = "/".join(SITE_URL.replace("https://", "").split("/")[1:])
  site_info_url = (
      f"https://graph.microsoft.com/v1.0/sites/{site_domain}:/{site_path}"
  )

  res = requests.get(site_info_url, headers=headers)
  site_id = res.json().get("id")

  # 2. Obtener los ítems de tu lista (reemplaza 'Colaboradores' por el nombre de tu lista)
  lists_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists"
  # ... Lógica para extraer los datos de las listas y cargarlos en tus DataFrames ...
