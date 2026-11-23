from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
import msal
import pandas as pd
import requests
import streamlit as st

st.set_page_config(
    page_title="Asistencia y Planilla - KCO Group",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# Barra superior 100% transparente y flotante + UI ejecutiva compacta para móvil
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
        margin-bottom: 0.55rem;
    }

    h3 {
        font-size: 1.02rem !important;
        margin-top: 0.75rem !important;
        margin-bottom: 0.35rem !important;
        color: #0f172a;
    }

    [data-testid="stMetric"] {
        background-color: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 8px 10px;
    }

    [data-testid="stMetricLabel"] {
        font-size: 0.74rem !important;
        color: #475569 !important;
    }

    [data-testid="stMetricValue"] {
        font-size: 1.15rem !important;
        font-weight: 700 !important;
    }

    .stSelectbox {
        margin-bottom: -0.4rem;
    }

    button[data-baseweb="tab"] {
        padding-top: 0.35rem !important;
        padding-bottom: 0.35rem !important;
        font-size: 0.88rem !important;
        font-weight: 600 !important;
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


# Encabezado ejecutivo
st.markdown(
    """
    <div class="app-title">Control Asistencia y Planilla — KCO Group</div>
    <div class="app-subtitle">Desarrollado por Alfredo Huaracha Mendoza</div>
    """,
    unsafe_allow_html=True,
)

try:
  df_col_raw, df_mar_raw = fetch_sharepoint_data()
except Exception as e:
  st.error(f"Error al conectar con SharePoint: {e}")
  st.stop()

if df_col_raw.empty:
  st.warning("No hay colaboradores registrados en SharePoint.")
  st.stop()

# 1. Procesar plantilla (dimColaboradores)
c_dni = find_col(df_col_raw, ["DNI"])
c_nom = find_col(df_col_raw, ["NombreCompleto"])
c_reg = find_col(df_col_raw, ["RegimenLaboral"])
c_hrs = find_col(
    df_col_raw, ["HorasSemanalesObjetivo", "HorasSemanales", "HorasSemanalesO"]
)
c_vig = find_col(df_col_raw, ["EsVigente"])
c_tar = find_col(
    df_col_raw, ["TarifaHora", "PagoHora", "SueldoHora", "CostoHora"]
)

cols_col = [c for c in [c_dni, c_nom, c_reg, c_hrs, c_vig, c_tar] if c]
df_col = df_col_raw[cols_col].copy()
rename_map = {
    c_dni: "DNI",
    c_nom: "NombreCompleto",
    c_reg: "Régimen",
    c_hrs: "Meta",
    c_vig: "EsVigente",
}
if c_tar:
  rename_map[c_tar] = "TarifaBase"

df_col = df_col.rename(columns={k: v for k, v in rename_map.items() if k})
df_col["DNI"] = df_col["DNI"].astype(str).str.strip()
df_col["Régimen"] = (
    df_col["Régimen"].astype(str).str.strip()
    if "Régimen" in df_col.columns
    else ""
)
df_col["Meta"] = (
    pd.to_numeric(df_col["Meta"], errors="coerce").fillna(0.0)
    if "Meta" in df_col.columns
    else 0.0
)
df_col["TarifaBase"] = (
    pd.to_numeric(df_col["TarifaBase"], errors="coerce").fillna(0.0)
    if "TarifaBase" in df_col.columns
    else 0.0
)

if "EsVigente" in df_col.columns:
  df_col = df_col[
      df_col["EsVigente"].astype(str).str.lower().isin(["si", "sí", "true", "1"])
  ]
df_col = df_col.drop_duplicates(subset=["DNI"]).copy()


def short_name(full_name):
  parts = str(full_name).strip().split()
  if len(parts) >= 3:
    return f"{parts[0]} {parts[-2]}"
  return " ".join(parts)


df_col["Colaborador"] = df_col["NombreCompleto"].apply(short_name)

# 2. Procesar marcaciones (factMarcaciones)
if not df_mar_raw.empty:
  m_dni = find_col(df_mar_raw, ["DNI"])
  m_tipo = find_col(df_mar_raw, ["TipoMarcacion"])
  m_fh = find_col(df_mar_raw, ["FechaHoraMarcacion", "FechaHoraMarcac"])

  df_mar = df_mar_raw[[c for c in [m_dni, m_tipo, m_fh] if c]].copy()
  df_mar = df_mar.rename(
      columns={
          k: v
          for k, v in {
              m_dni: "DNI",
              m_tipo: "Tipo",
              m_fh: "FechaHora_Raw",
          }.items()
          if k
      }
  )
  df_mar["DNI"] = df_mar["DNI"].astype(str).str.strip()

  dt = pd.to_datetime(df_mar["FechaHora_Raw"], errors="coerce", utc=True)
  df_mar["FechaHora"] = dt.dt.tz_convert("America/Lima").dt.tz_localize(None)
  df_mar = df_mar.dropna(subset=["FechaHora"]).copy()
  df_mar["InicioSemana"] = (
      df_mar["FechaHora"].dt.to_period("W-SUN").dt.start_time.dt.date
  )
  df_mar["Fecha"] = df_mar["FechaHora"].dt.date
else:
  df_mar = pd.DataFrame(
      columns=["DNI", "Tipo", "FechaHora", "InicioSemana", "Fecha"]
  )

# Filtrar plantilla estrictamente a quienes tienen historial de marcaciones
dnis_con_marcaciones = set(df_mar["DNI"].unique())
if dnis_con_marcaciones:
  df_col = df_col[df_col["DNI"].isin(dnis_con_marcaciones)].copy()

# 3. Construir catálogo de semanas garantizando la semana actual de Lima como predeterminada
hoy_lima = datetime.now(ZoneInfo("America/Lima")).date()
lunes_actual = hoy_lima - timedelta(days=hoy_lima.weekday())

inicios_semana = set(df_mar["InicioSemana"].dropna().unique())
inicios_semana.add(lunes_actual)
inicios_ordenados = sorted(inicios_semana, reverse=True)


def label_semana(inicio):
  fin = inicio + timedelta(days=6)
  return f"{inicio.strftime('%d/%m')} - {fin.strftime('%d/%m/%y')}"


semana_map = {label_semana(ini): ini for ini in inicios_ordenados}
lista_semanas = list(semana_map.keys())
idx_default = (
    lista_semanas.index(label_semana(lunes_actual))
    if label_semana(lunes_actual) in lista_semanas
    else 0
)

# Filtros globales en una sola fila
f1, f2, f3 = st.columns([4, 4, 1], vertical_alignment="center")
with f1:
  semana_sel = st.selectbox(
      "Semana",
      lista_semanas,
      index=idx_default,
      label_visibility="collapsed",
  )

inicio_sem_sel = semana_map[semana_sel]
fin_sem_sel = inicio_sem_sel + timedelta(days=6)

df_sem_mar = df_mar[df_mar["InicioSemana"] == inicio_sem_sel].copy()
dnis_semana = set(df_sem_mar["DNI"].unique())
df_plantilla_sem = (
    df_col[df_col["DNI"].isin(dnis_semana)].copy()
    if dnis_semana
    else df_col.iloc[0:0].copy()
)

with f2:
  colabs_sem = ["Todos"] + sorted(
      df_plantilla_sem["Colaborador"].dropna().unique().tolist()
  )
  colab_sel = st.selectbox(
      "Colaborador", colabs_sem, label_visibility="collapsed"
  )
with f3:
  if st.button(
      "↻", help="Actualizar datos en vivo", use_container_width=True
  ):
    st.cache_data.clear()
    st.rerun()

if colab_sel != "Todos":
  df_plantilla_sem = df_plantilla_sem[
      df_plantilla_sem["Colaborador"] == colab_sel
  ]
  df_sem_mar = df_sem_mar[df_sem_mar["DNI"].isin(df_plantilla_sem["DNI"])]

regimen_map = df_col.set_index("DNI")["Régimen"].to_dict()
meta_map = df_col.set_index("DNI")["Meta"].to_dict()


# Función reutilizable de cálculo diario (Escenario A: descuenta 1.0h refrigerio a Full Time > 5h)
def calcular_jornadas(df_marcaciones_sub):
  if df_marcaciones_sub.empty:
    return pd.DataFrame(
        columns=["DNI", "Fecha", "In", "Out", "Ref", "Hrs", "Pend"]
    )
  registros = []
  for (dni, fecha), grupo in df_marcaciones_sub.sort_values(
      "FechaHora"
  ).groupby(["DNI", "Fecha"]):
    ingresos = grupo[grupo["Tipo"].str.lower() == "ingreso"][
        "FechaHora"
    ].tolist()
    salidas = grupo[grupo["Tipo"].str.lower() == "salida"]["FechaHora"].tolist()

    horas_brutas = 0.0
    pares_validos = min(len(ingresos), len(salidas))
    for i in range(pares_validos):
      diff = (salidas[i] - ingresos[i]).total_seconds() / 3600.0
      if diff > 0:
        horas_brutas += diff

    regimen_actual = str(regimen_map.get(dni, "")).lower()
    es_full_time = ("full" in regimen_actual) or (meta_map.get(dni, 0) >= 40)

    descuento_refrigerio = 0.0
    if es_full_time and pares_validos == 1 and horas_brutas > 5.0:
      descuento_refrigerio = 1.0

    horas_netas = max(0.0, horas_brutas - descuento_refrigerio)
    incompleto = not (len(ingresos) == len(salidas) and len(ingresos) > 0)

    registros.append({
        "DNI": dni,
        "Fecha": fecha,
        "In": ingresos[0].strftime("%H:%M") if ingresos else "--:--",
        "Out": salidas[-1].strftime("%H:%M") if salidas else "--:--",
        "Ref": "-1h" if descuento_refrigerio > 0 else "—",
        "Hrs": round(horas_netas, 2),
        "Pend": 1 if incompleto else 0,
    })
  return pd.DataFrame(registros)


df_diario = calcular_jornadas(df_sem_mar)

# Consolidar Balance Semanal
if not df_diario.empty:
  agg_sem = df_diario.groupby("DNI", as_index=False).agg(
      Días=("Fecha", "nunique"),
      Trabajadas=("Hrs", "sum"),
      Sin_Salida=("Pend", "sum"),
  )
  df_resumen = df_plantilla_sem.merge(agg_sem, on="DNI", how="inner")
else:
  df_resumen = pd.DataFrame(
      columns=[
          "DNI",
          "Colaborador",
          "NombreCompleto",
          "Régimen",
          "Meta",
          "TarifaBase",
          "Días",
          "Trabajadas",
          "Sin_Salida",
      ]
  )

df_resumen["Trabajadas"] = df_resumen["Trabajadas"].fillna(0.0).round(2)
df_resumen["Sin_Salida"] = df_resumen["Sin_Salida"].fillna(0).astype(int)
df_resumen["Diferencia"] = (df_resumen["Trabajadas"] - df_resumen["Meta"]).round(
    2
)
df_resumen["Extra"] = df_resumen["Diferencia"].apply(lambda x: max(0.0, x))
df_resumen["Faltante"] = df_resumen["Diferencia"].apply(lambda x: max(0.0, -x))


def format_balance(row):
  diff = row["Diferencia"]
  alerta = " ⚠️" if row["Sin_Salida"] > 0 else ""
  if diff > 0:
    return f"+{diff:.1f}h (Extra){alerta}"
  elif diff < 0:
    return f"-{abs(diff):.1f}h (Falta){alerta}"
  return f"0.0h (OK){alerta}"


df_resumen["Balance"] = df_resumen.apply(format_balance, axis=1)

# 4. Pestañas principales: Asistencia vs. Planilla
tab_asistencia, tab_planilla = st.tabs(
    ["Control Asistencia", "Planilla Semanal"]
)

with tab_asistencia:
  k1, k2, k3 = st.columns(3)
  k1.metric("Trabajadas", f"{df_resumen['Trabajadas'].sum():.1f}h")
  k2.metric("Horas Extra", f"+{df_resumen['Extra'].sum():.1f}h")
  k3.metric("Por Completar", f"-{df_resumen['Faltante'].sum():.1f}h")

  st.subheader("Control Semanal por Colaborador")
  if not df_resumen.empty:
    tabla_principal = df_resumen.sort_values("Colaborador")[
        ["Colaborador", "Trabajadas", "Meta", "Balance"]
    ].copy()
    tabla_principal["Trabajadas"] = tabla_principal["Trabajadas"].apply(
        lambda x: f"{x:.1f}h"
    )
    tabla_principal["Meta"] = tabla_principal["Meta"].apply(
        lambda x: f"{int(x)}h" if x == int(x) else f"{x:.1f}h"
    )
    st.dataframe(tabla_principal, use_container_width=True, hide_index=True)

    st.caption(
        "Nota: A jornadas continuas Full Time (>5h) se les descuenta"
        " automáticamente 1.0h de refrigerio."
    )
    if df_resumen["Sin_Salida"].sum() > 0:
      st.caption(
          "⚠️ Indica jornadas abiertas sin marcación de salida registrada"
          " (horas parciales)."
      )
  else:
    st.info(f"Aún no hay marcaciones registradas en la semana ({semana_sel}).")

  with st.expander("Ver detalle diario de turnos"):
    if not df_diario.empty:
      df_diario_view = df_diario.merge(
          df_col[["DNI", "Colaborador"]], on="DNI", how="left"
      )
      dias_es = ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"]
      dt_f = pd.to_datetime(df_diario_view["Fecha"])
      df_diario_view["Día"] = [
          f"{dias_es[d.weekday()]} {d.strftime('%d/%m')}" for d in dt_f
      ]
      df_diario_view["Turno"] = (
          df_diario_view["In"] + "–" + df_diario_view["Out"]
      )
      df_diario_view["Hrs"] = df_diario_view["Hrs"].round(1).astype(str) + "h"

      st.dataframe(
          df_diario_view.sort_values(
              ["Fecha", "Colaborador"], ascending=[False, True]
          )[["Día", "Colaborador", "Turno", "Ref", "Hrs"]],
          use_container_width=True,
          hide_index=True,
      )
    else:
      st.info("Sin turnos registrados en este periodo.")

with tab_planilla:
  # Segmentador de fechas propio para cálculo de planilla (sincronizado con la semana elegida por defecto)
  rango_fechas = st.date_input(
      "Rango de fechas para cálculo de planilla",
      value=(inicio_sem_sel, fin_sem_sel),
      format="DD/MM/YYYY",
      key=f"rango_planilla_{semana_sel}",
  )

  if isinstance(rango_fechas, tuple) and len(rango_fechas) == 2:
    f_inicio, f_fin = rango_fechas
  elif isinstance(rango_fechas, tuple) and len(rango_fechas) == 1:
    f_inicio = f_fin = rango_fechas[0]
  else:
    f_inicio = f_fin = rango_fechas

  df_mar_plan = df_mar[
      (df_mar["Fecha"] >= f_inicio) & (df_mar["Fecha"] <= f_fin)
  ].copy()
  if colab_sel != "Todos":
    dnis_filtro = df_col[df_col["Colaborador"] == colab_sel]["DNI"].tolist()
    df_mar_plan = df_mar_plan[df_mar_plan["DNI"].isin(dnis_filtro)]

  df_diario_plan = calcular_jornadas(df_mar_plan)

  if df_diario_plan.empty:
    st.info(
        "No hay marcaciones registradas entre el"
        f" {f_inicio.strftime('%d/%m/%Y')} y el {f_fin.strftime('%d/%m/%Y')}."
    )
  else:
    # Escalar meta proporcionalmente si el rango abarca más o menos de 7 días
    dias_rango = max(1, (f_fin - f_inicio).days + 1)
    factor_semanas = dias_rango / 7.0

    agg_plan = df_diario_plan.groupby("DNI", as_index=False).agg(
        Días=("Fecha", "nunique"),
        Trabajadas=("Hrs", "sum"),
        Sin_Salida=("Pend", "sum"),
    )
    df_plan_base = df_col.merge(agg_plan, on="DNI", how="inner")
    df_plan_base["MetaPeriodo"] = (
        df_plan_base["Meta"] * factor_semanas
    ).round(2)
    df_plan_base["Regulares"] = df_plan_base.apply(
        lambda r: min(r["Trabajadas"], r["MetaPeriodo"])
        if r["MetaPeriodo"] > 0
        else r["Trabajadas"],
        axis=1,
    ).round(2)
    df_plan_base["Extra"] = df_plan_base.apply(
        lambda r: max(0.0, r["Trabajadas"] - r["MetaPeriodo"])
        if r["MetaPeriodo"] > 0
        else 0.0,
        axis=1,
    ).round(2)

    p1, p2 = st.columns(2)
    with p1:
      tarifa_default = st.number_input(
          "Tarifa base por hora (S/)",
          min_value=0.0,
          value=7.50,
          step=0.50,
          format="%.2f",
          help="Se aplica si no hay tarifa definida en SharePoint.",
      )
    with p2:
      factor_he = st.number_input(
          "Factor Hora Extra (ej. 1.25 = +25%)",
          min_value=1.0,
          value=1.25,
          step=0.05,
          format="%.2f",
      )

    df_calc = df_plan_base.sort_values("Colaborador")[
        [
            "DNI",
            "Colaborador",
            "Régimen",
            "Días",
            "Regulares",
            "Extra",
            "TarifaBase",
            "Sin_Salida",
        ]
    ].copy()
    df_calc["S/ Hora"] = df_calc["TarifaBase"].apply(
        lambda x: float(x) if x > 0 else float(tarifa_default)
    )

    st.caption("Edita **S/ Hora** en la tabla para ajustar tarifas:")
    df_editado = st.data_editor(
        df_calc[["Colaborador", "Días", "Regulares", "Extra", "S/ Hora"]],
        column_config={
            "Colaborador": st.column_config.TextColumn(disabled=True),
            "Días": st.column_config.NumberColumn(disabled=True, format="%d"),
            "Regulares": st.column_config.NumberColumn(
                "Hrs Reg.", disabled=True, format="%.1f"
            ),
            "Extra": st.column_config.NumberColumn(
                "Hrs Ext.", disabled=True, format="%.1f"
            ),
            "S/ Hora": st.column_config.NumberColumn(
                "S/ Hora", min_value=0.0, step=0.5, format="S/ %.2f"
            ),
        },
        use_container_width=True,
        hide_index=True,
        key=f"editor_planilla_{f_inicio}_{f_fin}_{colab_sel}",
    )

    df_calc["S/ Hora"] = df_editado["S/ Hora"].fillna(0.0)
    df_calc["Pago Regular"] = (
        df_calc["Regulares"] * df_calc["S/ Hora"]
    ).round(2)
    df_calc["Pago Extra"] = (
        df_calc["Extra"] * df_calc["S/ Hora"] * factor_he
    ).round(2)
    df_calc["Total S/"] = (
        df_calc["Pago Regular"] + df_calc["Pago Extra"]
    ).round(2)

    m1, m2, m3 = st.columns(3)
    m1.metric("Pago Regular", f"S/ {df_calc['Pago Regular'].sum():,.2f}")
    m2.metric("Pago H. Extra", f"S/ {df_calc['Pago Extra'].sum():,.2f}")
    m3.metric("Total Planilla", f"S/ {df_calc['Total S/'].sum():,.2f}")

    st.subheader("Liquidación del Periodo")
    tabla_pago = df_calc[
        [
            "Colaborador",
            "Regulares",
            "Extra",
            "Pago Regular",
            "Pago Extra",
            "Total S/",
        ]
    ].copy()
    tabla_pago["Regulares"] = tabla_pago["Regulares"].apply(
        lambda x: f"{x:.1f}h"
    )
    tabla_pago["Extra"] = tabla_pago["Extra"].apply(
        lambda x: f"+{x:.1f}h" if x > 0 else "0.0h"
    )
    tabla_pago["Pago Regular"] = tabla_pago["Pago Regular"].apply(
        lambda x: f"S/ {x:,.2f}"
    )
    tabla_pago["Pago Extra"] = tabla_pago["Pago Extra"].apply(
        lambda x: f"S/ {x:,.2f}"
    )
    tabla_pago["Total S/"] = tabla_pago["Total S/"].apply(
        lambda x: f"S/ {x:,.2f}"
    )

    st.dataframe(tabla_pago, use_container_width=True, hide_index=True)

    if df_calc["Sin_Salida"].sum() > 0:
      st.warning(
          "Hay marcaciones sin salida registrada en este rango; verifica los"
          " turnos antes de cerrar el pago."
      )

    csv_export = df_calc[[
        "DNI",
        "Colaborador",
        "Régimen",
        "Días",
        "Regulares",
        "Extra",
        "S/ Hora",
        "Pago Regular",
        "Pago Extra",
        "Total S/",
    ]].to_csv(index=False)

    st.download_button(
        label="Descargar Planilla (CSV)",
        data=csv_export.encode("utf-8-sig"),
        file_name=f"planilla_kco_{f_inicio.strftime('%Y%m%d')}_{f_fin.strftime('%Y%m%d')}.csv",
        mime="text/csv",
        use_container_width=True,
    )
