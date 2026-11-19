import msal
import pandas as pd
import requests
import streamlit as st

st.set_page_config(
    page_title="Asistencia KCO",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# Barra superior 100% transparente y flotante + diseño móvil compacto
st.markdown(
    """
<style>
    header[data-testid="stHeader"] {
        background: transparent !important;
        backdrop-filter: none !important;
        -webkit-backdrop-filter: none !important;
        border: none !important;
        pointer-events: none;
    }

    header[data-testid="stHeader"] * {
        pointer-events: auto;
    }

    .block-container {
        padding-top: 0.85rem !important;
        padding-bottom: 1.2rem !important;
        padding-left: 0.8rem !important;
        padding-right: 0.8rem !important;
        max-width: 900px;
    }

    .app-title {
        font-size: 1.25rem;
        font-weight: 700;
        color: #0f172a;
        line-height: 1.2;
        margin: 0;
        padding-right: 110px;
    }

    .app-subtitle {
        font-size: 0.75rem;
        color: #64748b;
        margin-top: 2px;
        margin-bottom: 0.65rem;
    }

    h3 {
        font-size: 1.02rem !important;
        margin-top: 0.75rem !important;
        margin-bottom: 0.35rem !important;
    }

    [data-testid="stMetric"] {
        background-color: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 8px 10px;
    }

    [data-testid="stMetricLabel"] {
        font-size: 0.75rem !important;
    }

    [data-testid="stMetricValue"] {
        font-size: 1.15rem !important;
    }

    .stSelectbox {
        margin-bottom: -0.4rem;
    }
</style>
""",
    unsafe_allow_html=True,
)

# Cargar credenciales desde st.secrets
TENANT_ID = st.secrets["sharepoint"]["tenant_id"]
CLIENT_ID = st.secrets["sharepoint"]["client_id"]
CLIENT_SECRET = st.secrets["sharepoint"]["client_secret"]
SITE_URL = st.secrets["sharepoint"]["site_url"].rstrip("/")


@st.cache_data(ttl=120)
def fetch_sharepoint_data():
  authority = f"https://login.microsoftonline.com/{TENANT_ID}"
  app = msal.ConfidentialClientApplication(
      CLIENT_ID,
      authority=authority,
      client_credential=CLIENT_SECRET,
  )
  token_res = app.acquire_token_for_client(
      scopes=["https://graph.microsoft.com/.default"]
  )
  if "access_token" not in token_res:
    raise RuntimeError(
        f"Error de autenticación: {token_res.get('error_description')}"
    )

  headers = {
      "Authorization": f"Bearer {token_res['access_token']}",
      "Accept": "application/json",
  }

  clean_url = SITE_URL.replace("https://", "")
  domain = clean_url.split("/")[0]
  path = "/".join(clean_url.split("/")[1:])
  site_resp = requests.get(
      f"https://graph.microsoft.com/v1.0/sites/{domain}:/{path}",
      headers=headers,
  )
  site_resp.raise_for_status()
  site_id = site_resp.json()["id"]

  def get_list_items(list_name):
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_name}/items?expand=fields&$top=999"
    rows = []
    while url:
      r = requests.get(url, headers=headers)
      r.raise_for_status()
      payload = r.json()
      for item in payload.get("value", []):
        rows.append(item.get("fields", {}))
      url = payload.get("@odata.nextLink")
    return pd.DataFrame(rows)

  return get_list_items("dimColaboradores"), get_list_items("factMarcaciones")


def find_col(df, candidates):
  for c in candidates:
    if c in df.columns:
      return c
  for col in df.columns:
    for c in candidates:
      if col.lower().startswith(c.lower()[:12]):
        return col
  return None


# Encabezado limpio a todo lo ancho
st.markdown(
    """
    <div class="app-title">Asistencia Semanal — KCO Group</div>
    <div class="app-subtitle">Desarrollado por Alfredo Huaracha Mendoza</div>
    """,
    unsafe_allow_html=True,
)

try:
  df_col_raw, df_mar_raw = fetch_sharepoint_data()
except Exception as e:
  st.error(f"Error al conectar con SharePoint: {e}")
  st.stop()

if df_col_raw.empty or df_mar_raw.empty:
  st.warning("No hay registros disponibles en las listas.")
  st.stop()

# Mapear dimColaboradores
c_dni = find_col(df_col_raw, ["DNI"])
c_nom = find_col(df_col_raw, ["NombreCompleto"])
c_reg = find_col(df_col_raw, ["RegimenLaboral"])
c_hrs = find_col(
    df_col_raw, ["HorasSemanalesObjetivo", "HorasSemanales", "HorasSemanalesO"]
)
c_vig = find_col(df_col_raw, ["EsVigente"])

df_col = df_col_raw[[
    c for c in [c_dni, c_nom, c_reg, c_hrs, c_vig] if c
]].copy()
df_col = df_col.rename(
    columns={
        k: v
        for k, v in {
            c_dni: "DNI",
            c_nom: "NombreCompleto",
            c_reg: "Régimen",
            c_hrs: "Obj",
            c_vig: "EsVigente",
        }.items()
        if k
    }
)
df_col["DNI"] = df_col["DNI"].astype(str).str.strip()
if "Obj" in df_col.columns:
  df_col["Obj"] = pd.to_numeric(df_col["Obj"], errors="coerce").fillna(0)
else:
  df_col["Obj"] = 0

if "EsVigente" in df_col.columns:
  df_col = df_col[
      df_col["EsVigente"].astype(str).str.lower().isin(["si", "sí", "true", "1"])
  ]
df_col = df_col.drop_duplicates(subset=["DNI"])


def short_name(full_name):
  parts = str(full_name).strip().split()
  if len(parts) >= 3:
    return f"{parts[0]} {parts[-2]}"
  return " ".join(parts)


df_col["Colaborador"] = df_col["NombreCompleto"].apply(short_name)

# Mapear factMarcaciones
m_dni = find_col(df_mar_raw, ["DNI"])
m_tipo = find_col(df_mar_raw, ["TipoMarcacion"])
m_fh = find_col(df_mar_raw, ["FechaHoraMarcacion", "FechaHoraMarcac"])
m_obs = find_col(df_mar_raw, ["Observaciones"])

df_mar = df_mar_raw[[c for c in [m_dni, m_tipo, m_fh, m_obs] if c]].copy()
df_mar = df_mar.rename(
    columns={
        k: v
        for k, v in {
            m_dni: "DNI",
            m_tipo: "Tipo",
            m_fh: "FechaHora_Raw",
            m_obs: "Obs",
        }.items()
