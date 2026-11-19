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
        if k
    }
)
df_mar["DNI"] = df_mar["DNI"].astype(str).str.strip()

# Zona horaria Lima
dt = pd.to_datetime(df_mar["FechaHora_Raw"], errors="coerce", utc=True)
df_mar["FechaHora"] = dt.dt.tz_convert("America/Lima").dt.tz_localize(None)
df_mar = df_mar.dropna(subset=["FechaHora"]).copy()

# Etiqueta semanal compacta
df_mar["InicioSemana"] = df_mar["FechaHora"].dt.to_period("W-SUN").dt.start_time
df_mar["Semana"] = (
    df_mar["InicioSemana"].dt.strftime("%d/%m")
    + " - "
    + (df_mar["InicioSemana"] + pd.Timedelta(days=6)).dt.strftime("%d/%m/%y")
)
df_mar["Fecha"] = df_mar["FechaHora"].dt.date
df_mar["Hora"] = df_mar["FechaHora"].dt.strftime("%H:%M")

df_merged = df_mar.merge(df_col, on="DNI", how="left")
df_merged["Colaborador"] = df_merged["Colaborador"].fillna(
    "DNI " + df_merged["DNI"]
)
df_merged["Obj"] = df_merged["Obj"].fillna(0)

# Fila única de filtros + botón de actualización a la derecha
semanas_df = (
    df_merged[["InicioSemana", "Semana"]]
    .drop_duplicates()
    .sort_values("InicioSemana", ascending=False)
)
f1, f2, f3 = st.columns([4, 4, 1], vertical_alignment="center")
with f1:
  semana_sel = st.selectbox(
      "Semana", semanas_df["Semana"].tolist(), label_visibility="collapsed"
  )
with f2:
  colabs = ["Todos"] + sorted(
      df_col["Colaborador"].dropna().unique().tolist()
  )
  colab_sel = st.selectbox(
      "Colaborador", colabs, label_visibility="collapsed"
  )
with f3:
  if st.button(
      "↻", help="Actualizar datos en vivo", use_container_width=True
  ):
    st.cache_data.clear()
    st.rerun()

df_sem = df_merged[df_merged["Semana"] == semana_sel].copy()
if colab_sel != "Todos":
  df_sem = df_sem[df_sem["Colaborador"] == colab_sel]

# Cálculo diario emparejando Ingreso y Salida
registros = []
for (dni, fecha), grupo in df_sem.sort_values("FechaHora").groupby(
    ["DNI", "Fecha"]
):
  ingresos = grupo[grupo["Tipo"].str.lower() == "ingreso"]["FechaHora"].tolist()
  salidas = grupo[grupo["Tipo"].str.lower() == "salida"]["FechaHora"].tolist()

  horas = 0.0
  for i in range(min(len(ingresos), len(salidas))):
    diff = (salidas[i] - ingresos[i]).total_seconds() / 3600.0
    if diff > 0:
      horas += diff

  incompleto = not (len(ingresos) == len(salidas) and len(ingresos) > 0)
  registros.append({
      "DNI": dni,
      "Colaborador": grupo["Colaborador"].iloc[0],
      "Obj": float(grupo["Obj"].iloc[0]),
      "Fecha": fecha,
      "In": ingresos[0].strftime("%H:%M") if ingresos else "--:--",
      "Out": salidas[-1].strftime("%H:%M") if salidas else "--:--",
      "Hrs": round(horas, 2),
      "Pend": 1 if incompleto else 0,
  })

df_diario = pd.DataFrame(registros)

# Resumen semanal por colaborador y cálculo de horas extra
if not df_diario.empty:
  resumen_base = df_diario.groupby(["Colaborador", "Obj"], as_index=False).agg(
      Días=("Fecha", "nunique"),
      Hrs=("Hrs", "sum"),
      Pend=("Pend", "sum"),
  )
  resumen_base["Extra"] = resumen_base.apply(
      lambda r: max(0.0, r["Hrs"] - r["Obj"]) if r["Obj"] > 0 else 0.0, axis=1
  )
  total_horas = resumen_base["Hrs"].sum()
  total_extra = resumen_base["Extra"].sum()
  total_pend = int(resumen_base["Pend"].sum())
else:
  resumen_base = pd.DataFrame()
  total_horas = 0.0
  total_extra = 0.0
  total_pend = 0

# 3 métricas operativas compactas
k1, k2, k3 = st.columns(3)
k1.metric("Horas Totales", f"{total_horas:.1f}h")
k2.metric("Horas Extra", f"{total_extra:.1f}h")
k3.metric("Sin Salida", total_pend)

# Vista 1: Resumen Semanal
st.subheader("Resumen Semanal")
if not resumen_base.empty:
  resumen = resumen_base.assign(
      Avance=lambda d: d["Hrs"].round(1).astype(str)
      + "/"
      + d["Obj"].astype(int).astype(str)
      + "h",
      H_Extra=lambda d: d["Extra"].apply(
          lambda x: f"+{x:.1f}h" if x > 0 else "—"
      ),
  )[["Colaborador", "Días", "Avance", "H_Extra", "Pend"]].rename(
      columns={"H_Extra": "Extra", "Pend": "Sin Sal."}
  )
  st.dataframe(resumen, use_container_width=True, hide_index=True)

# Vista 2: Detalle Diario
st.subheader("Detalle Diario")
if not df_diario.empty:
  dias_es = ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"]
  dt_f = pd.to_datetime(df_diario["Fecha"])
  df_diario["Día"] = [
      f"{dias_es[d.weekday()]} {d.strftime('%d/%m')}" for d in dt_f
  ]
  df_diario["Turno"] = df_diario["In"] + "–" + df_diario["Out"]
  df_diario["Hrs_Txt"] = df_diario["Hrs"].round(1).astype(str) + "h"

  tabla_diaria = df_diario.sort_values(
      ["Fecha", "Colaborador"], ascending=[False, True]
  )[["Día", "Colaborador", "Turno", "Hrs_Txt"]].rename(
      columns={"Hrs_Txt": "Hrs"}
  )
  st.dataframe(tabla_diaria, use_container_width=True, hide_index=True)

# Log crudo oculto en expander
with st.expander("Historial de marcaciones"):
  cols = ["Fecha", "Hora", "Colaborador", "Tipo"]
  if "Obs" in df_sem.columns:
    cols.append("Obs")
  st.dataframe(
      df_sem.sort_values("FechaHora", ascending=False)[cols],
      use_container_width=True,
      hide_index=True,
  )
