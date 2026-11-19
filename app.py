import msal
import pandas as pd
import requests
import streamlit as st

# Configuración de página (tema claro y ancho completo)
st.set_page_config(
    page_title="Control de Asistencia - KCO Group",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# Cargar credenciales desde st.secrets
TENANT_ID = st.secrets["sharepoint"]["tenant_id"]
CLIENT_ID = st.secrets["sharepoint"]["client_id"]
CLIENT_SECRET = st.secrets["sharepoint"]["client_secret"]
SITE_URL = st.secrets["sharepoint"]["site_url"].rstrip("/")


@st.cache_data(ttl=120)
def fetch_sharepoint_data():
  # 1. Obtener token con MSAL
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
        f"Error de autenticación Azure: {token_res.get('error_description')}"
    )

  headers = {
      "Authorization": f"Bearer {token_res['access_token']}",
      "Accept": "application/json",
  }

  # 2. Resolver Site ID desde la URL de SharePoint
  clean_url = SITE_URL.replace("https://", "")
  domain = clean_url.split("/")[0]
  path = "/".join(clean_url.split("/")[1:])
  site_resp = requests.get(
      f"https://graph.microsoft.com/v1.0/sites/{domain}:/{path}",
      headers=headers,
  )
  site_resp.raise_for_status()
  site_id = site_resp.json()["id"]

  # 3. Función auxiliar para descargar todos los ítems de una lista (con paginación)
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

  df_col = get_list_items("dimColaboradores")
  df_mar = get_list_items("factMarcaciones")
  return df_col, df_mar


def find_col(df, candidates):
  for c in candidates:
    if c in df.columns:
      return c
  for col in df.columns:
    for c in candidates:
      if col.lower().startswith(c.lower()[:12]):
        return col
  return None


st.title("Control de Asistencia Semanal — KCO Group")

col_btn, col_info = st.columns([1, 5])
with col_btn:
  if st.button("Actualizar datos"):
    st.cache_data.clear()
    st.rerun()

try:
  df_col_raw, df_mar_raw = fetch_sharepoint_data()
except Exception as e:
  st.error(f"No se pudo conectar con SharePoint: {e}")
  st.stop()

if df_col_raw.empty or df_mar_raw.empty:
  st.warning(
      "Una de las listas está vacía en SharePoint. Verifica que existan"
      " registros."
  )
  st.stop()

# Normalizar columnas de dimColaboradores
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
rename_col = {}
if c_dni:
  rename_col[c_dni] = "DNI"
if c_nom:
  rename_col[c_nom] = "Colaborador"
if c_reg:
  rename_col[c_reg] = "Régimen"
if c_hrs:
  rename_col[c_hrs] = "Horas Objetivo"
if c_vig:
  rename_col[c_vig] = "EsVigente"
df_col = df_col.rename(columns=rename_col)
df_col["DNI"] = df_col["DNI"].astype(str).str.strip()

if "EsVigente" in df_col.columns:
  df_col = df_col[
      df_col["EsVigente"].astype(str).str.lower().isin(["si", "sí", "true", "1"])
  ]
df_col = df_col.drop_duplicates(subset=["DNI"])

# Normalizar columnas de factMarcaciones
m_dni = find_col(df_mar_raw, ["DNI"])
m_tipo = find_col(df_mar_raw, ["TipoMarcacion"])
m_fh = find_col(df_mar_raw, ["FechaHoraMarcacion", "FechaHoraMarcac"])
m_obs = find_col(df_mar_raw, ["Observaciones"])

df_mar = df_mar_raw[[c for c in [m_dni, m_tipo, m_fh, m_obs] if c]].copy()
rename_mar = {}
if m_dni:
  rename_mar[m_dni] = "DNI"
if m_tipo:
  rename_mar[m_tipo] = "Tipo"
if m_fh:
  rename_mar[m_fh] = "FechaHora_Raw"
if m_obs:
  rename_mar[m_obs] = "Observaciones"
df_mar = df_mar.rename(columns=rename_mar)
df_mar["DNI"] = df_mar["DNI"].astype(str).str.strip()

# Convertir fecha/hora a zona horaria de Lima
dt = pd.to_datetime(df_mar["FechaHora_Raw"], errors="coerce", utc=True)
df_mar["FechaHora"] = dt.dt.tz_convert("America/Lima").dt.tz_localize(None)
df_mar = df_mar.dropna(subset=["FechaHora"]).copy()

# Calcular inicio de semana (Lunes) y etiqueta semanal
df_mar["InicioSemana"] = df_mar["FechaHora"].dt.to_period("W-SUN").dt.start_time
df_mar["Semana"] = (
    df_mar["InicioSemana"].dt.strftime("%d/%m/%Y")
    + " al "
    + (df_mar["InicioSemana"] + pd.Timedelta(days=6)).dt.strftime("%d/%m/%Y")
)
df_mar["Fecha"] = df_mar["FechaHora"].dt.date
df_mar["Hora"] = df_mar["FechaHora"].dt.strftime("%H:%M")

# Cruzar marcaciones con colaboradores
df_merged = df_mar.merge(df_col, on="DNI", how="left")
df_merged["Colaborador"] = df_merged["Colaborador"].fillna(
    "DNI " + df_merged["DNI"]
)

# Filtros superiores
semanas_df = (
    df_merged[["InicioSemana", "Semana"]]
    .drop_duplicates()
    .sort_values("InicioSemana", ascending=False)
)
lista_semanas = semanas_df["Semana"].tolist()

f1, f2 = st.columns([2, 2])
with f1:
  semana_sel = st.selectbox("Seleccionar Semana (Lunes a Domingo)", lista_semanas)
with f2:
  colaboradores_disp = ["Todos"] + sorted(
      df_col["Colaborador"].dropna().unique().tolist()
  )
  colab_sel = st.selectbox("Filtrar Colaborador", colaboradores_disp)

df_sem = df_merged[df_merged["Semana"] == semana_sel].copy()
if colab_sel != "Todos":
  df_sem = df_sem[df_sem["Colaborador"] == colab_sel]

# Calcular horas trabajadas por día (emparejando Ingreso y Salida por DNI y Fecha)
registros_diarios = []
for (dni, fecha), grupo in df_sem.sort_values("FechaHora").groupby(
    ["DNI", "Fecha"]
):
  ingresos = grupo[grupo["Tipo"].str.lower() == "ingreso"]["FechaHora"].tolist()
  salidas = grupo[grupo["Tipo"].str.lower() == "salida"]["FechaHora"].tolist()

  horas_dia = 0.0
  pares = min(len(ingresos), len(salidas))
  for i in range(pares):
    diff = (salidas[i] - ingresos[i]).total_seconds() / 3600.0
    if diff > 0:
      horas_dia += diff

  estado = (
      "Completo"
      if len(ingresos) == len(salidas) and len(ingresos) > 0
      else "Pendiente salida"
  )
  registros_diarios.append({
      "DNI": dni,
      "Colaborador": grupo["Colaborador"].iloc[0],
      "Régimen": (
          grupo["Régimen"].iloc[0] if "Régimen" in grupo.columns else "-"
      ),
      "Horas Objetivo": (
          grupo["Horas Objetivo"].iloc[0]
          if "Horas Objetivo" in grupo.columns
          else 0
      ),
      "Fecha": fecha,
      "Primer Ingreso": ingresos[0].strftime("%H:%M") if ingresos else "-",
      "Última Salida": salidas[-1].strftime("%H:%M") if salidas else "-",
      "Horas Día": round(horas_dia, 2),
      "Estado": estado,
  })

df_diario = pd.DataFrame(registros_diarios)

# KPIs de la semana
k1, k2, k3 = st.columns(3)
k1.metric("Total Marcaciones", len(df_sem))
k2.metric(
    "Colaboradores con Registro",
    df_sem["DNI"].nunique(),
    f"de {len(df_col)} vigentes",
)
k3.metric(
    "Horas Acumuladas (con salida)",
    f"{df_diario['Horas Día'].sum():.1f} h" if not df_diario.empty else "0.0 h",
)

st.divider()

# 1. Resumen Semanal por Colaborador
st.subheader("Resumen Semanal por Colaborador")
if not df_diario.empty:
  resumen = (
      df_diario.groupby(
          ["DNI", "Colaborador", "Régimen", "Horas Objetivo"], as_index=False
      )
      .agg(
          Días_Asistidos=("Fecha", "nunique"),
          Horas_Acumuladas=("Horas Día", "sum"),
          Marcaciones_Incompletas=(
              "Estado",
              lambda s: (s == "Pendiente salida").sum(),
          ),
      )
      .rename(
          columns={
              "Días_Asistidos": "Días Asistidos",
              "Horas_Acumuladas": "Horas Acumuladas",
              "Marcaciones_Incompletas": "Días sin Salida",
          }
      )
  )
  st.dataframe(resumen, use_container_width=True, hide_index=True)

# 2. Matriz Diaria de la Semana (Lunes a Domingo)
st.subheader("Detalle Diario de la Semana")
if not df_diario.empty:
  df_diario["Día"] = pd.to_datetime(df_diario["Fecha"]).dt.strftime(
      "%a %d/%m"
  )
  df_diario["Resumen Día"] = (
      df_diario["Primer Ingreso"]
      + " - "
      + df_diario["Última Salida"]
      + " ("
      + df_diario["Horas Día"].astype(str)
      + "h)"
  )
  matriz = df_diario.pivot_table(
      index=["Colaborador", "Régimen"],
      columns="Día",
      values="Resumen Día",
      aggfunc="first",
      fill_value="—",
  ).reset_index()
  st.dataframe(matriz, use_container_width=True, hide_index=True)

# 3. Log Crudo de Marcaciones de la Semana
with st.expander("Ver historial detallado de marcaciones de la semana"):
  cols_show = ["Fecha", "Hora", "Colaborador", "DNI", "Tipo"]
  if "Observaciones" in df_sem.columns:
    cols_show.append("Observaciones")
  st.dataframe(
      df_sem.sort_values("FechaHora", ascending=False)[cols_show],
      use_container_width=True,
      hide_index=True,
  )
