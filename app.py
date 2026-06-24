import os
import re
from io import BytesIO
from pathlib import Path
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt

st.set_page_config(page_title="CEMEX | Legalizaciones Cúcuta", layout="wide")

# =========================
# ESTILO
# =========================
st.markdown("""
<style>
.main {background-color:#f7f8fa;}
.block-container {padding-top:1.2rem; padding-bottom:2rem;}
[data-testid="stMetricValue"] {font-size: 1.55rem; font-weight: 800;}
[data-testid="stMetricLabel"] {font-size: 0.9rem; color:#404040;}
.card {
    background:white; border:1px solid #e6e8eb; border-radius:14px;
    padding:15px; box-shadow:0 1px 4px rgba(0,0,0,.05); height:100%;
}
.small {font-size: 0.86rem; color:#555;}
.kpi-title {font-size:0.82rem;color:#555;margin-bottom:4px;}
.kpi-value {font-size:1.45rem;font-weight:800;color:#111;}
.kpi-sub {font-size:0.78rem;color:#777;}
.alerta {background:#fff4e5;border-left:5px solid #f59e0b;padding:10px;border-radius:8px;}
.ok {background:#eaf7ee;border-left:5px solid #22c55e;padding:10px;border-radius:8px;}
.bad {background:#fdecec;border-left:5px solid #ef4444;padding:10px;border-radius:8px;}
</style>
""", unsafe_allow_html=True)

# =========================
# UTILIDADES
# =========================
def money(x):
    try:
        if pd.isna(x):
            return "$ 0"
        return "$ {:,.0f}".format(float(x)).replace(",", ".")
    except Exception:
        return "$ 0"

def pct(x):
    try:
        if pd.isna(x):
            return "0,00%"
        return f"{float(x)*100:,.2f}%".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "0,00%"

def clean_col(c):
    c = str(c).strip().upper()
    c = re.sub(r"\s+", " ", c)
    c = c.replace("Á","A").replace("É","E").replace("Í","I").replace("Ó","O").replace("Ú","U").replace("Ñ","N")
    return c

def normalize_text(s):
    if pd.isna(s):
        return ""
    return re.sub(r"\s+", " ", str(s).strip().upper())

def to_number(s):
    if pd.isna(s):
        return 0.0
    if isinstance(s, (int, float, np.number)):
        return float(s)
    txt = str(s).strip()
    txt = re.sub(r"[^0-9,.-]", "", txt)
    if txt == "":
        return 0.0
    # Manejo Colombia: 1.234.567,89
    if "," in txt and "." in txt:
        txt = txt.replace(".", "").replace(",", ".")
    elif "," in txt and "." not in txt:
        txt = txt.replace(",", ".")
    try:
        return float(txt)
    except Exception:
        return 0.0

def to_date(s):
    if pd.isna(s):
        return pd.NaT
    return pd.to_datetime(s, errors="coerce", dayfirst=True)

def find_header_row(raw, required_words):
    """Busca fila encabezado por palabras clave."""
    for i in range(min(len(raw), 20)):
        vals = [clean_col(v) for v in list(raw.iloc[i].values)]
        joined = " | ".join(vals)
        hits = sum(1 for w in required_words if w in joined)
        if hits >= max(2, len(required_words)//2):
            return i
    return 0

def read_sheet_flexible(file, sheet_name, required_words):
    raw = pd.read_excel(file, sheet_name=sheet_name, header=None, engine="openpyxl")
    header = find_header_row(raw, required_words)
    df = pd.read_excel(file, sheet_name=sheet_name, header=header, engine="openpyxl")
    df = df.dropna(how="all")
    df.columns = [clean_col(c) for c in df.columns]
    df = df.loc[:, ~pd.Series(df.columns).duplicated().values]
    return df

def standardize(df, tipo):
    if df is None or df.empty:
        return pd.DataFrame()
    out = df.copy()
    out.columns = [clean_col(c) for c in out.columns]

    # Quitar filas de encabezado repetidas
    if "FECHA" in out.columns:
        out = out[out["FECHA"].astype(str).str.upper().str.strip() != "FECHA"]

    mapping = {
        "FECHA":"FECHA", "HORA":"HORA", "CC":"CC", "CONDUCTOR":"CONDUCTOR", "PLACA":"PLACA",
        "ORIGEN":"ORIGEN", "DESTINO":"DESTINO", "RUTA":"RUTA", "CANTIDAD USUARIO":"USUARIOS",
        "FUNCIONARIOS":"FUNCIONARIO", "FUNCIONARIO":"FUNCIONARIO", "TURNO":"TURNO", "INGRESO":"INGRESO",
        "JORNADA":"JORNADA", "REMESA":"REMESA", "VALOR":"VALOR", "TARIFA":"VALOR",
        "OBSERVACIONES":"OBSERVACIONES", "QUIEN SOLICITA":"QUIEN SOLICITA"
    }
    for src, dst in mapping.items():
        if src in out.columns and dst not in out.columns:
            out[dst] = out[src]

    base_cols = ["FECHA","HORA","CC","CONDUCTOR","PLACA","ORIGEN","DESTINO","RUTA","USUARIOS","FUNCIONARIO","TURNO","INGRESO","JORNADA","REMESA","VALOR","OBSERVACIONES","QUIEN SOLICITA"]
    for c in base_cols:
        if c not in out.columns:
            out[c] = np.nan
    out = out[base_cols]
    out["TIPO_SERVICIO"] = tipo
    out["FECHA"] = out["FECHA"].apply(to_date)
    out["VALOR"] = out["VALOR"].apply(to_number)
    out["REMESA"] = out["REMESA"].apply(normalize_text)
    out["PLACA"] = out["PLACA"].apply(normalize_text)
    out["CONDUCTOR"] = out["CONDUCTOR"].apply(normalize_text)
    out["CC"] = out["CC"].apply(lambda x: "" if pd.isna(x) else re.sub(r"\.0$", "", str(x).strip()))
    for c in ["ORIGEN","DESTINO","RUTA","JORNADA","TURNO","OBSERVACIONES","QUIEN SOLICITA","FUNCIONARIO"]:
        out[c] = out[c].apply(normalize_text)
    out["DIA"] = out["FECHA"].dt.date
    out["DIA_SEMANA"] = out["FECHA"].dt.day_name(locale=None)
    out["MES"] = out["FECHA"].dt.to_period("M").astype(str)
    out["TIENE_REMESA"] = out["REMESA"].str.len() > 0
    out["TIENE_PLACA"] = out["PLACA"].str.len() > 0
    out["TIENE_CONDUCTOR"] = out["CONDUCTOR"].str.len() > 0
    out["TIENE_CC"] = out["CC"].str.len() > 0
    out["TIENE_VALOR"] = out["VALOR"] > 0
    out["FECHA_VALIDA"] = out["FECHA"].notna()
    return out

@st.cache_data(show_spinner="Leyendo archivo Excel...")
def read_workbook(file_bytes_or_path):
    xls = pd.ExcelFile(file_bytes_or_path, engine="openpyxl")
    sheets = xls.sheet_names
    rec = pd.DataFrame(); adi = pd.DataFrame(); resumen = pd.DataFrame(); turnos = pd.DataFrame()
    for sh in sheets:
        nsh = clean_col(sh)
        if "RECORR" in nsh:
            rec = read_sheet_flexible(file_bytes_or_path, sh, ["FECHA", "CONDUCTOR", "PLACA", "REMESA", "VALOR"])
        elif "ADIC" in nsh:
            adi = read_sheet_flexible(file_bytes_or_path, sh, ["FECHA", "CONDUCTOR", "PLACA", "REMESA", "TARIFA"])
        elif "RESUM" in nsh:
            resumen = read_sheet_flexible(file_bytes_or_path, sh, ["CIUDAD", "PLANTA", "RUTA", "TARIFA"])
        elif "TURNO" in nsh:
            turnos = read_sheet_flexible(file_bytes_or_path, sh, ["RUTA", "TURNO", "JORNADA"])
    return rec, adi, resumen, turnos, sheets

def bar_chart(df, x, y, title, top=15, horizontal=True):
    data = df.copy().head(top)
    if data.empty:
        st.info("Sin datos para graficar.")
        return
    fig, ax = plt.subplots(figsize=(9, max(3, len(data)*0.35) if horizontal else 4))
    if horizontal:
        data = data.sort_values(y, ascending=True)
        ax.barh(data[x].astype(str), data[y])
        ax.set_xlabel(y)
    else:
        ax.bar(data[x].astype(str), data[y])
        ax.tick_params(axis='x', rotation=45)
    ax.set_title(title)
    ax.grid(axis="x" if horizontal else "y", alpha=.25)
    st.pyplot(fig, clear_figure=True)

def kpi(label, value, sub=""):
    st.markdown(f"""
    <div class='card'>
      <div class='kpi-title'>{label}</div>
      <div class='kpi-value'>{value}</div>
      <div class='kpi-sub'>{sub}</div>
    </div>
    """, unsafe_allow_html=True)

def make_excel_download(df_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for name, df in df_dict.items():
            safe = name[:31]
            df.to_excel(writer, sheet_name=safe, index=False)
    output.seek(0)
    return output

# =========================
# CARGA ARCHIVO
# =========================
st.title("📊 CEMEX Cúcuta | Informe de Legalizaciones")
st.caption("Dashboard ejecutivo para control de remesas, soportes, pendientes, duplicados y cierre operativo.")

with st.sidebar:
    st.header("📁 Archivo")
    uploaded = st.file_uploader("Carga el Excel consolidado CEMEX", type=["xlsx"])
    local_files = sorted([str(p) for p in Path(".").glob("*.xlsx")])
    default_name = "CONSOLIDADO CUCUTA JUNIO 2026_V23.xlsx"
    st.caption("Si estás en Streamlit Cloud, sube el archivo desde este botón. Así evitamos el error FileNotFoundError.")

if uploaded is not None:
    file_source = BytesIO(uploaded.getvalue())
    source_name = uploaded.name
elif os.path.exists(default_name):
    file_source = default_name
    source_name = default_name
elif local_files:
    file_source = local_files[0]
    source_name = local_files[0]
else:
    st.error("No encontré el archivo Excel. Súbelo con el botón lateral o carga el archivo al repositorio GitHub.")
    st.stop()

try:
    rec_raw, adi_raw, resumen_raw, turnos_raw, sheets = read_workbook(file_source)
except Exception as e:
    st.error("No pude leer el archivo. Verifica que sea .xlsx y que requirements.txt tenga openpyxl.")
    st.exception(e)
    st.stop()

rec = standardize(rec_raw, "PROGRAMADO")
adi = standardize(adi_raw, "ADICIONAL")
df = pd.concat([rec, adi], ignore_index=True)
df = df.dropna(how="all")

if df.empty:
    st.error("El archivo se leyó, pero no encontré datos en Recorridos o Adicionales.")
    st.stop()

# =========================
# FILTROS
# =========================
with st.sidebar:
    st.success(f"Archivo activo: {source_name}")
    st.write("Hojas detectadas:", ", ".join(sheets))
    tipos = st.multiselect("Tipo de servicio", sorted(df["TIPO_SERVICIO"].dropna().unique()), default=sorted(df["TIPO_SERVICIO"].dropna().unique()))
    jornadas = st.multiselect("Jornada", sorted([x for x in df["JORNADA"].dropna().unique() if x]), default=sorted([x for x in df["JORNADA"].dropna().unique() if x]))
    placas = st.multiselect("Placa", sorted([x for x in df["PLACA"].dropna().unique() if x])[:300])
    conductores = st.multiselect("Conductor", sorted([x for x in df["CONDUCTOR"].dropna().unique() if x])[:300])
    fecha_min = df["FECHA"].min()
    fecha_max = df["FECHA"].max()
    if pd.notna(fecha_min) and pd.notna(fecha_max):
        rango = st.date_input("Rango de fechas", value=(fecha_min.date(), fecha_max.date()))
    else:
        rango = None

f = df.copy()
if tipos:
    f = f[f["TIPO_SERVICIO"].isin(tipos)]
if jornadas:
    f = f[f["JORNADA"].isin(jornadas)]
if placas:
    f = f[f["PLACA"].isin(placas)]
if conductores:
    f = f[f["CONDUCTOR"].isin(conductores)]
if rango and isinstance(rango, tuple) and len(rango) == 2:
    ini, fin = pd.to_datetime(rango[0]), pd.to_datetime(rango[1])
    f = f[(f["FECHA"].isna()) | ((f["FECHA"] >= ini) & (f["FECHA"] <= fin))]

# =========================
# KPIS
# =========================
servicios = len(f)
programados = int((f["TIPO_SERVICIO"] == "PROGRAMADO").sum())
adicionales = int((f["TIPO_SERVICIO"] == "ADICIONAL").sum())
valor_total = f["VALOR"].sum()
remesas_ok = int(f["TIENE_REMESA"].sum())
remesas_pend = int((~f["TIENE_REMESA"]).sum())
legalizacion = remesas_ok / servicios if servicios else 0
valor_pendiente = f.loc[~f["TIENE_REMESA"], "VALOR"].sum()
valor_legalizado = f.loc[f["TIENE_REMESA"], "VALOR"].sum()
duplicadas = int(f[f["REMESA"].ne("")].duplicated("REMESA", keep=False).sum())
fechas_malas = int((~f["FECHA_VALIDA"]).sum())
placas_pend = int((~f["TIENE_PLACA"]).sum())
conduct_pend = int((~f["TIENE_CONDUCTOR"]).sum())
cc_pend = int((~f["TIENE_CC"]).sum())
valor_pend = int((~f["TIENE_VALOR"]).sum())

st.subheader("1) Tarjetas ejecutivas de legalización")
cols = st.columns(5)
with cols[0]: kpi("Servicios totales", f"{servicios:,}".replace(",","."), "Programados + adicionales")
with cols[1]: kpi("Programados", f"{programados:,}".replace(",","."), pct(programados/servicios if servicios else 0))
with cols[2]: kpi("Adicionales", f"{adicionales:,}".replace(",","."), pct(adicionales/servicios if servicios else 0))
with cols[3]: kpi("Valor legalizable", money(valor_total), "Base filtrada")
with cols[4]: kpi("% legalización", pct(legalizacion), f"{remesas_ok} con remesa")

cols = st.columns(5)
with cols[0]: kpi("Valor legalizado", money(valor_legalizado), "Con remesa")
with cols[1]: kpi("Valor pendiente", money(valor_pendiente), "Sin remesa")
with cols[2]: kpi("Remesas pendientes", f"{remesas_pend:,}".replace(",","."), "Cierre requerido")
with cols[3]: kpi("Remesas duplicadas", f"{duplicadas:,}".replace(",","."), "Revisar antes de facturar")
with cols[4]: kpi("Fechas inválidas", f"{fechas_malas:,}".replace(",","."), "Calidad de base")

# =========================
# SEMÁFORO ANALISTA
# =========================
st.subheader("2) Diagnóstico de cierre")
if legalizacion >= 0.95 and remesas_pend == 0 and duplicadas == 0:
    st.markdown("<div class='ok'><b>Estado:</b> Cierre sano. La base está lista para validación final y facturación.</div>", unsafe_allow_html=True)
elif legalizacion >= 0.85:
    st.markdown("<div class='alerta'><b>Estado:</b> Cierre en riesgo medio. Priorizar remesas pendientes y duplicadas.</div>", unsafe_allow_html=True)
else:
    st.markdown("<div class='bad'><b>Estado:</b> Cierre crítico. El porcentaje de legalización es bajo y requiere gestión inmediata.</div>", unsafe_allow_html=True)

c1, c2, c3, c4 = st.columns(4)
with c1: kpi("Placas pendientes", f"{placas_pend:,}".replace(",","."), "Control documental")
with c2: kpi("Conductores pendientes", f"{conduct_pend:,}".replace(",","."), "Responsable del soporte")
with c3: kpi("CC pendientes", f"{cc_pend:,}".replace(",","."), "Identificación conductor")
with c4: kpi("Valores pendientes", f"{valor_pend:,}".replace(",","."), "Tarifa no registrada")

# =========================
# ANÁLISIS
# =========================
st.subheader("3) Análisis operativo y financiero")
tab1, tab2, tab3, tab4, tab5 = st.tabs(["📅 Diario", "🚗 Placas", "👤 Conductores", "🧾 Remesas", "⚠️ Calidad"])

with tab1:
    diario = f.groupby("DIA", dropna=False).agg(
        SERVICIOS=("TIPO_SERVICIO", "size"),
        VALOR=("VALOR", "sum"),
        CON_REMESA=("TIENE_REMESA", "sum")
    ).reset_index()
    diario["% LEGALIZACION"] = diario["CON_REMESA"] / diario["SERVICIOS"]
    st.dataframe(diario, use_container_width=True)
    chart = diario.dropna(subset=["DIA"]).copy()
    chart["DIA"] = chart["DIA"].astype(str)
    bar_chart(chart.sort_values("DIA"), "DIA", "SERVICIOS", "Servicios por día", top=40, horizontal=False)

with tab2:
    placas_df = f.groupby("PLACA").agg(SERVICIOS=("PLACA","size"), VALOR=("VALOR","sum"), REMESAS_OK=("TIENE_REMESA","sum")).reset_index()
    placas_df = placas_df[placas_df["PLACA"] != ""].sort_values("SERVICIOS", ascending=False)
    placas_df["% LEGALIZACION"] = placas_df["REMESAS_OK"] / placas_df["SERVICIOS"]
    st.dataframe(placas_df, use_container_width=True)
    bar_chart(placas_df, "PLACA", "SERVICIOS", "Top placas por cantidad de servicios")

with tab3:
    cond_df = f.groupby("CONDUCTOR").agg(SERVICIOS=("CONDUCTOR","size"), VALOR=("VALOR","sum"), REMESAS_OK=("TIENE_REMESA","sum")).reset_index()
    cond_df = cond_df[cond_df["CONDUCTOR"] != ""].sort_values("SERVICIOS", ascending=False)
    cond_df["% LEGALIZACION"] = cond_df["REMESAS_OK"] / cond_df["SERVICIOS"]
    st.dataframe(cond_df, use_container_width=True)
    bar_chart(cond_df, "CONDUCTOR", "SERVICIOS", "Top conductores por cantidad de servicios")

with tab4:
    dup = f[(f["REMESA"] != "") & (f.duplicated("REMESA", keep=False))].sort_values("REMESA")
    pendientes = f[~f["TIENE_REMESA"]].copy()
    st.markdown("**Remesas duplicadas**")
    st.dataframe(dup, use_container_width=True)
    st.markdown("**Servicios pendientes de remesa**")
    st.dataframe(pendientes, use_container_width=True)

with tab5:
    calidad = pd.DataFrame({
        "ALERTA": ["Sin remesa", "Remesa duplicada", "Sin fecha válida", "Sin placa", "Sin conductor", "Sin CC", "Sin valor"],
        "CONTEO": [remesas_pend, duplicadas, fechas_malas, placas_pend, conduct_pend, cc_pend, valor_pend]
    }).sort_values("CONTEO", ascending=False)
    st.dataframe(calidad, use_container_width=True)
    bar_chart(calidad, "ALERTA", "CONTEO", "Pareto de alertas de legalización", top=10)

# =========================
# MATRIZ DE GESTIÓN
# =========================
st.subheader("4) Matriz de gestión para cierre")
gestion = f.copy()
gestion["ESTADO_LEGALIZACION"] = np.where(gestion["TIENE_REMESA"], "LEGALIZADO", "PENDIENTE REMESA")
gestion["ALERTA"] = ""
gestion.loc[~gestion["TIENE_REMESA"], "ALERTA"] += "SIN REMESA | "
gestion.loc[gestion["REMESA"].ne("") & gestion.duplicated("REMESA", keep=False), "ALERTA"] += "REMESA DUPLICADA | "
gestion.loc[~gestion["FECHA_VALIDA"], "ALERTA"] += "FECHA INVALIDA | "
gestion.loc[~gestion["TIENE_PLACA"], "ALERTA"] += "SIN PLACA | "
gestion.loc[~gestion["TIENE_CONDUCTOR"], "ALERTA"] += "SIN CONDUCTOR | "
gestion.loc[~gestion["TIENE_CC"], "ALERTA"] += "SIN CC | "
gestion.loc[~gestion["TIENE_VALOR"], "ALERTA"] += "SIN VALOR | "
gestion["ALERTA"] = gestion["ALERTA"].str.rstrip(" | ")
cols_show = ["ESTADO_LEGALIZACION","ALERTA","TIPO_SERVICIO","FECHA","HORA","REMESA","PLACA","CONDUCTOR","CC","ORIGEN","DESTINO","RUTA","JORNADA","TURNO","VALOR","OBSERVACIONES","QUIEN SOLICITA"]
st.dataframe(gestion[cols_show], use_container_width=True, height=420)

resumen_kpis = pd.DataFrame({
    "KPI": ["Servicios totales", "Programados", "Adicionales", "Valor legalizable", "Valor legalizado", "Valor pendiente", "% Legalizacion", "Remesas pendientes", "Remesas duplicadas"],
    "VALOR": [servicios, programados, adicionales, valor_total, valor_legalizado, valor_pendiente, legalizacion, remesas_pend, duplicadas]
})
excel = make_excel_download({
    "KPIS": resumen_kpis,
    "MATRIZ_GESTION": gestion[cols_show],
    "DIARIO": diario if 'diario' in locals() else pd.DataFrame(),
})
st.download_button("⬇️ Descargar matriz de gestión en Excel", data=excel, file_name="informe_legalizaciones_cemex_cucuta.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.caption("Desarrollado para control de legalizaciones CEMEX: remesa, valor, conductor, placa, fecha, duplicados y cierre operativo.")
