# app_cemex_legalizaciones.py
# Dashboard Streamlit - Informe de Legalizaciones CEMEX
# Ejecutar: streamlit run app_cemex_legalizaciones.py

from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(
    page_title="CEMEX | Informe de Legalizaciones",
    page_icon="📋",
    layout="wide",
)

DEFAULT_FILE = Path("CONSOLIDADO CUCUTA JUNIO 2026_V23.xlsx")

# -----------------------------
# ESTILOS
# -----------------------------
st.markdown(
    """
    <style>
    .main .block-container {padding-top: 1.5rem; padding-bottom: 2rem;}
    .kpi-card {
        border: 1px solid #e8e8e8; border-radius: 14px; padding: 16px 18px;
        background: #ffffff; box-shadow: 0 1px 4px rgba(0,0,0,0.06); min-height: 112px;
    }
    .kpi-title {font-size: 0.82rem; color: #555; margin-bottom: 6px; font-weight: 600;}
    .kpi-value {font-size: 1.65rem; color: #111; font-weight: 800; line-height: 1.15;}
    .kpi-sub {font-size: 0.78rem; color: #666; margin-top: 6px;}
    .alert-ok {color: #116b2e; font-weight: 700;}
    .alert-warn {color: #b35c00; font-weight: 700;}
    .alert-risk {color: #a60000; font-weight: 700;}
    </style>
    """,
    unsafe_allow_html=True,
)

# -----------------------------
# FUNCIONES BASE
# -----------------------------
def clean_col(col: object) -> str:
    col = str(col).strip().upper()
    col = re.sub(r"\s+", " ", col)
    return col


def money(x: float) -> str:
    try:
        return "$ {:,.0f}".format(float(x)).replace(",", ".")
    except Exception:
        return "$ 0"


def pct(x: float) -> str:
    if pd.isna(x) or np.isinf(x):
        return "0,00%"
    return f"{x:.2%}".replace(".", ",")


def kpi_card(title: str, value: str, subtitle: str = ""):
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-title">{title}</div>
            <div class="kpi-value">{value}</div>
            <div class="kpi-sub">{subtitle}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def first_existing(df: pd.DataFrame, candidates: list[str]) -> str | None:
    cols = {clean_col(c): c for c in df.columns}
    for c in candidates:
        if clean_col(c) in cols:
            return cols[clean_col(c)]
    return None


def to_num(s: pd.Series) -> pd.Series:
    return pd.to_numeric(
        s.astype(str).str.replace("$", "", regex=False).str.replace(".", "", regex=False).str.replace(",", ".", regex=False),
        errors="coerce",
    )


def normalize_text(s: pd.Series) -> pd.Series:
    return s.fillna("").astype(str).str.strip().str.upper().replace({"NAN": "", "NONE": "", "NULL": ""})


@st.cache_data(show_spinner=False)
def read_workbook(file) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, Dict[str, str]]:
    """Lee el libro CEMEX. Recorridos tiene encabezado en fila 5 del Excel, por eso header=4."""
    xls = pd.ExcelFile(file)
    sheets = xls.sheet_names

    rec = pd.read_excel(file, sheet_name="Recorridos", header=4) if "Recorridos" in sheets else pd.DataFrame()
    adi = pd.read_excel(file, sheet_name="Adicionales") if "Adicionales" in sheets else pd.DataFrame()
    res = pd.read_excel(file, sheet_name="Resumen", header=1) if "Resumen" in sheets else pd.DataFrame()
    turnos = pd.read_excel(file, sheet_name="TURNOS") if "TURNOS" in sheets else pd.DataFrame()

    fuente = {"hojas": ", ".join(sheets)}
    return rec, adi, res, turnos, fuente


def prepare_operacion(rec: pd.DataFrame, adi: pd.DataFrame) -> pd.DataFrame:
    frames = []

    if not rec.empty:
        df = rec.copy()
        df.columns = [clean_col(c) for c in df.columns]
        if "#" in df.columns:
            df = df[pd.to_numeric(df["#"], errors="coerce").notna()].copy()
        df["TIPO_SERVICIO"] = "RECORRIDO PROGRAMADO"
        if "VALOR" in df.columns:
            df["VALOR_SERVICIO"] = to_num(df["VALOR"])
        else:
            df["VALOR_SERVICIO"] = 0
        frames.append(df)

    if not adi.empty:
        df = adi.copy()
        df.columns = [clean_col(c) for c in df.columns]
        if "#" in df.columns:
            df = df[pd.to_numeric(df["#"], errors="coerce").notna()].copy()
        df = df[df.get("FECHA", pd.Series(index=df.index)).notna()].copy()
        df["TIPO_SERVICIO"] = "ADICIONAL / URBANO"
        if "TARIFA" in df.columns:
            df["VALOR_SERVICIO"] = to_num(df["TARIFA"])
        else:
            df["VALOR_SERVICIO"] = 0
        if "HORA" not in df.columns:
            df["HORA"] = ""
        if "RUTA" not in df.columns:
            df["RUTA"] = "ADICIONAL / URBANO"
        if "CANTIDAD USUARIO" not in df.columns:
            df["CANTIDAD USUARIO"] = np.nan
        frames.append(df)

    if not frames:
        return pd.DataFrame()

    all_cols = sorted(set().union(*[set(f.columns) for f in frames]))
    base = pd.concat([f.reindex(columns=all_cols) for f in frames], ignore_index=True)

    for c in ["FECHA", "REMESA", "PLACA", "CONDUCTOR", "CC", "ORIGEN", "DESTINO", "RUTA", "JORNADA", "TURNO", "INGRESO", "OBSERVACIONES", "QUIEN SOLICITA"]:
        if c not in base.columns:
            base[c] = ""

    base["FECHA"] = pd.to_datetime(base["FECHA"], errors="coerce", dayfirst=True)
    base["DIA"] = base["FECHA"].dt.day
    base["DIA_SEMANA"] = base["FECHA"].dt.day_name(locale="es_ES") if hasattr(base["FECHA"].dt, "day_name") else base["FECHA"].dt.day_name()
    base["SEMANA"] = base["FECHA"].dt.isocalendar().week.astype("Int64")
    base["MES"] = base["FECHA"].dt.strftime("%Y-%m")
    base["REMESA_TXT"] = normalize_text(base["REMESA"])
    base["PLACA_TXT"] = normalize_text(base["PLACA"])
    base["CONDUCTOR_TXT"] = normalize_text(base["CONDUCTOR"])
    base["CC_TXT"] = normalize_text(base["CC"])
    base["VALOR_SERVICIO"] = pd.to_numeric(base["VALOR_SERVICIO"], errors="coerce").fillna(0)

    obligatorios = ["FECHA", "REMESA_TXT", "PLACA_TXT", "CONDUCTOR_TXT", "CC_TXT", "VALOR_SERVICIO"]
    base["FALTA_FECHA"] = base["FECHA"].isna()
    base["FALTA_REMESA"] = base["REMESA_TXT"].eq("")
    base["FALTA_PLACA"] = base["PLACA_TXT"].eq("")
    base["FALTA_CONDUCTOR"] = base["CONDUCTOR_TXT"].eq("")
    base["FALTA_CC"] = base["CC_TXT"].eq("")
    base["FALTA_VALOR"] = base["VALOR_SERVICIO"].le(0)
    base["CAMPOS_FALTANTES"] = base[["FALTA_FECHA", "FALTA_REMESA", "FALTA_PLACA", "FALTA_CONDUCTOR", "FALTA_CC", "FALTA_VALOR"]].sum(axis=1)
    base["ESTADO_LEGALIZACION"] = np.where(base["CAMPOS_FALTANTES"].eq(0), "LEGALIZADO", "PENDIENTE / REVISAR")
    base["REMESA_DUPLICADA"] = base["REMESA_TXT"].ne("") & base.duplicated("REMESA_TXT", keep=False)
    return base


def agg_count_value(df: pd.DataFrame, group: list[str]) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()
    return (
        df.groupby(group, dropna=False)
        .agg(SERVICIOS=("TIPO_SERVICIO", "size"), VALOR=("VALOR_SERVICIO", "sum"), USUARIOS=("CANTIDAD USUARIO", "sum"))
        .reset_index()
        .sort_values("SERVICIOS", ascending=False)
    )


# -----------------------------
# CARGA
# -----------------------------
st.title("📋 Informe de Legalizaciones | Cliente CEMEX")
st.caption("Análisis operativo, documental y de legalización para servicios programados, urbanos y adicionales.")

uploaded = st.sidebar.file_uploader("Cargar archivo Excel CEMEX", type=["xlsx", "xls"])
file_source = uploaded if uploaded is not None else DEFAULT_FILE

try:
    rec_raw, adi_raw, resumen_raw, turnos_raw, fuente = read_workbook(file_source)
except Exception as e:
    st.error("No pude leer el archivo. Verifica que esté en la misma carpeta del app o vuelve a cargarlo desde el panel lateral.")
    st.exception(e)
    st.stop()

base = prepare_operacion(rec_raw, adi_raw)
if base.empty:
    st.warning("No se encontraron datos en las hojas Recorridos o Adicionales.")
    st.stop()

# -----------------------------
# FILTROS
# -----------------------------
st.sidebar.header("Filtros")
min_date, max_date = base["FECHA"].min(), base["FECHA"].max()
if pd.notna(min_date) and pd.notna(max_date):
    rango = st.sidebar.date_input("Rango de fechas", value=(min_date.date(), max_date.date()))
else:
    rango = None

tipos = st.sidebar.multiselect("Tipo de servicio", sorted(base["TIPO_SERVICIO"].dropna().unique()), default=sorted(base["TIPO_SERVICIO"].dropna().unique()))
jornadas = st.sidebar.multiselect("Jornada", sorted([x for x in base["JORNADA"].dropna().unique() if str(x).strip() != ""]), default=sorted([x for x in base["JORNADA"].dropna().unique() if str(x).strip() != ""]))
estados = st.sidebar.multiselect("Estado legalización", sorted(base["ESTADO_LEGALIZACION"].unique()), default=sorted(base["ESTADO_LEGALIZACION"].unique()))

f = base.copy()
if rango and isinstance(rango, tuple) and len(rango) == 2:
    f = f[(f["FECHA"].dt.date >= rango[0]) & (f["FECHA"].dt.date <= rango[1])]
if tipos:
    f = f[f["TIPO_SERVICIO"].isin(tipos)]
if jornadas:
    f = f[(f["JORNADA"].isin(jornadas)) | (f["JORNADA"].astype(str).str.strip().eq(""))]
if estados:
    f = f[f["ESTADO_LEGALIZACION"].isin(estados)]

# -----------------------------
# KPIS PRINCIPALES
# -----------------------------
total_serv = len(f)
valor_total = f["VALOR_SERVICIO"].sum()
legalizados = int((f["ESTADO_LEGALIZACION"] == "LEGALIZADO").sum())
pendientes = int((f["ESTADO_LEGALIZACION"] != "LEGALIZADO").sum())
porc_legal = legalizados / total_serv if total_serv else 0
remesas_unicas = f.loc[f["REMESA_TXT"].ne(""), "REMESA_TXT"].nunique()
dias_operados = f["FECHA"].dt.date.nunique()
adicionales = int((f["TIPO_SERVICIO"] == "ADICIONAL / URBANO").sum())
programados = int((f["TIPO_SERVICIO"] == "RECORRIDO PROGRAMADO").sum())

st.subheader("1. Tarjetas de conteo ejecutivo")
c1, c2, c3, c4, c5 = st.columns(5)
with c1: kpi_card("Servicios totales", f"{total_serv:,.0f}".replace(",", "."), f"Programados: {programados} | Adic.: {adicionales}")
with c2: kpi_card("Valor legalizable", money(valor_total), "Suma de VALOR/TARIFA")
with c3: kpi_card("% legalizado", pct(porc_legal), f"{legalizados} servicios completos")
with c4: kpi_card("Pendientes/revisar", f"{pendientes:,.0f}".replace(",", "."), "Falta remesa, fecha, placa, conductor, CC o valor")
with c5: kpi_card("Remesas únicas", f"{remesas_unicas:,.0f}".replace(",", "."), f"Días operados: {dias_operados}")

c6, c7, c8, c9, c10 = st.columns(5)
with c6: kpi_card("Sin remesa", f"{int(f['FALTA_REMESA'].sum()):,.0f}".replace(",", "."), "Crítico para legalización")
with c7: kpi_card("Remesa duplicada", f"{int(f['REMESA_DUPLICADA'].sum()):,.0f}".replace(",", "."), "Revisar posibles cruces")
with c8: kpi_card("Sin placa", f"{int(f['FALTA_PLACA'].sum()):,.0f}".replace(",", "."), "Validación documental")
with c9: kpi_card("Sin conductor/CC", f"{int((f['FALTA_CONDUCTOR'] | f['FALTA_CC']).sum()):,.0f}".replace(",", "."), "Validación conductor")
with c10: kpi_card("Sin valor", f"{int(f['FALTA_VALOR'].sum()):,.0f}".replace(",", "."), "No legalizable financieramente")

# -----------------------------
# ALERTA ANALÍTICA
# -----------------------------
st.subheader("2. Diagnóstico de legalizaciones")
if pendientes == 0 and int(f["REMESA_DUPLICADA"].sum()) == 0:
    st.markdown("<span class='alert-ok'>✅ La base filtrada está completa para legalización: no presenta faltantes críticos ni remesas duplicadas.</span>", unsafe_allow_html=True)
else:
    st.markdown(
        f"<span class='alert-warn'>⚠️ Se deben revisar {pendientes} servicios con campos incompletos y {int(f['REMESA_DUPLICADA'].sum())} registros con remesa duplicada.</span>",
        unsafe_allow_html=True,
    )

faltantes = pd.DataFrame({
    "Campo crítico": ["Fecha", "Remesa", "Placa", "Conductor", "CC", "Valor/Tarifa"],
    "Registros con novedad": [int(f["FALTA_FECHA"].sum()), int(f["FALTA_REMESA"].sum()), int(f["FALTA_PLACA"].sum()), int(f["FALTA_CONDUCTOR"].sum()), int(f["FALTA_CC"].sum()), int(f["FALTA_VALOR"].sum())],
})
left, right = st.columns([1, 2])
with left:
    st.dataframe(faltantes, use_container_width=True, hide_index=True)
with right:
    fig = px.bar(faltantes, x="Campo crítico", y="Registros con novedad", text="Registros con novedad", title="Campos pendientes para legalización")
    fig.update_layout(height=360, yaxis_title="Registros", xaxis_title="")
    st.plotly_chart(fig, use_container_width=True)

# -----------------------------
# VISUALES OPERATIVOS
# -----------------------------
st.subheader("3. Distribución operativa y financiera")
tab1, tab2, tab3, tab4 = st.tabs(["Por fecha", "Por ruta/jornada", "Por conductor/placa", "Adicionales"])

with tab1:
    diario = agg_count_value(f, ["FECHA", "TIPO_SERVICIO"])
    if not diario.empty:
        diario["FECHA_TXT"] = diario["FECHA"].dt.strftime("%d/%m/%Y")
        fig = px.bar(diario, x="FECHA_TXT", y="SERVICIOS", color="TIPO_SERVICIO", title="Servicios por día")
        fig.update_layout(height=420, xaxis_title="Fecha", yaxis_title="Servicios")
        st.plotly_chart(fig, use_container_width=True)
        diario_valor = f.groupby("FECHA", dropna=False).agg(SERVICIOS=("TIPO_SERVICIO", "size"), VALOR=("VALOR_SERVICIO", "sum"), REMESAS=("REMESA_TXT", "nunique")).reset_index()
        diario_valor["FECHA"] = diario_valor["FECHA"].dt.strftime("%d/%m/%Y")
        diario_valor["VALOR"] = diario_valor["VALOR"].map(money)
        st.dataframe(diario_valor, use_container_width=True, hide_index=True)

with tab2:
    col_a, col_b = st.columns(2)
    rutas = agg_count_value(f, ["RUTA"])
    jornadas_df = agg_count_value(f, ["JORNADA"])
    with col_a:
        fig = px.bar(rutas.head(15), x="SERVICIOS", y="RUTA", orientation="h", text="SERVICIOS", title="Top rutas por cantidad de servicios")
        fig.update_layout(height=420, yaxis_title="", xaxis_title="Servicios")
        st.plotly_chart(fig, use_container_width=True)
    with col_b:
        fig = px.pie(jornadas_df, names="JORNADA", values="SERVICIOS", title="Participación por jornada")
        fig.update_layout(height=420)
        st.plotly_chart(fig, use_container_width=True)
    rutas_show = rutas.copy()
    rutas_show["VALOR"] = rutas_show["VALOR"].map(money)
    st.dataframe(rutas_show, use_container_width=True, hide_index=True)

with tab3:
    col_a, col_b = st.columns(2)
    conductores = agg_count_value(f[f["CONDUCTOR_TXT"].ne("")], ["CONDUCTOR"])
    placas = agg_count_value(f[f["PLACA_TXT"].ne("")], ["PLACA"])
    with col_a:
        fig = px.bar(conductores.head(12), x="SERVICIOS", y="CONDUCTOR", orientation="h", text="SERVICIOS", title="Servicios por conductor")
        fig.update_layout(height=450, yaxis_title="", xaxis_title="Servicios")
        st.plotly_chart(fig, use_container_width=True)
    with col_b:
        fig = px.bar(placas.head(12), x="SERVICIOS", y="PLACA", orientation="h", text="SERVICIOS", title="Servicios por placa")
        fig.update_layout(height=450, yaxis_title="", xaxis_title="Servicios")
        st.plotly_chart(fig, use_container_width=True)

with tab4:
    adi_f = f[f["TIPO_SERVICIO"] == "ADICIONAL / URBANO"].copy()
    if adi_f.empty:
        st.info("No hay adicionales en el filtro seleccionado.")
    else:
        solicitantes = agg_count_value(adi_f, ["QUIEN SOLICITA"])
        solicitantes_show = solicitantes.copy()
        solicitantes_show["VALOR"] = solicitantes_show["VALOR"].map(money)
        col_a, col_b = st.columns(2)
        with col_a:
            fig = px.bar(solicitantes, x="SERVICIOS", y="QUIEN SOLICITA", orientation="h", text="SERVICIOS", title="Adicionales por solicitante")
            fig.update_layout(height=420, yaxis_title="", xaxis_title="Servicios")
            st.plotly_chart(fig, use_container_width=True)
        with col_b:
            obs = agg_count_value(adi_f, ["OBSERVACIONES"])
            fig = px.bar(obs.head(10), x="SERVICIOS", y="OBSERVACIONES", orientation="h", text="SERVICIOS", title="Adicionales por observación/tipo")
            fig.update_layout(height=420, yaxis_title="", xaxis_title="Servicios")
            st.plotly_chart(fig, use_container_width=True)
        st.dataframe(solicitantes_show, use_container_width=True, hide_index=True)

# -----------------------------
# TABLA DE NOVEDADES PARA GESTIÓN
# -----------------------------
st.subheader("4. Matriz de gestión para cierre y legalización")
cols_show = ["TIPO_SERVICIO", "FECHA", "REMESA", "PLACA", "CONDUCTOR", "CC", "ORIGEN", "DESTINO", "RUTA", "JORNADA", "TURNO", "VALOR_SERVICIO", "ESTADO_LEGALIZACION", "CAMPOS_FALTANTES", "REMESA_DUPLICADA", "OBSERVACIONES", "QUIEN SOLICITA"]
cols_show = [c for c in cols_show if c in f.columns]
gestion = f[cols_show].copy()
if "FECHA" in gestion.columns:
    gestion["FECHA"] = pd.to_datetime(gestion["FECHA"], errors="coerce").dt.strftime("%d/%m/%Y")
if "VALOR_SERVICIO" in gestion.columns:
    gestion["VALOR_SERVICIO"] = gestion["VALOR_SERVICIO"].map(money)

solo_novedades = st.toggle("Ver solo pendientes / duplicados", value=True)
if solo_novedades:
    idx = (f["ESTADO_LEGALIZACION"] != "LEGALIZADO") | (f["REMESA_DUPLICADA"])
    gestion = gestion.loc[idx]
st.dataframe(gestion, use_container_width=True, hide_index=True)

# Descargar Excel de gestión
export = f.copy()
export["FECHA"] = pd.to_datetime(export["FECHA"], errors="coerce").dt.strftime("%d/%m/%Y")
bytes_out = None
try:
    import io
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        export.to_excel(writer, sheet_name="Base_Gestion_Legalizacion", index=False)
        faltantes.to_excel(writer, sheet_name="Resumen_Faltantes", index=False)
        agg_count_value(f, ["RUTA"]).to_excel(writer, sheet_name="Resumen_Rutas", index=False)
        agg_count_value(f, ["CONDUCTOR"]).to_excel(writer, sheet_name="Resumen_Conductores", index=False)
    bytes_out = buffer.getvalue()
except Exception:
    bytes_out = None

if bytes_out:
    st.download_button(
        "⬇️ Descargar matriz de gestión en Excel",
        data=bytes_out,
        file_name="matriz_legalizacion_cemex.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# -----------------------------
# COMPARATIVO CONTRA HOJA RESUMEN
# -----------------------------
st.subheader("5. Validación contra hoja Resumen")
if not resumen_raw.empty:
    st.caption("La hoja Resumen del archivo sirve como referencia, pero este dashboard audita directamente Recorridos + Adicionales.")
    st.dataframe(resumen_raw, use_container_width=True, hide_index=True)
    st.info(f"Total calculado desde las bases operativas: {total_serv} servicios | {money(valor_total)}")
else:
    st.info("No se encontró hoja Resumen para comparar.")

st.caption("Informe diseñado con enfoque de legalizaciones: completitud de remesa, trazabilidad de conductor/placa, valor legalizable, duplicidades y control de adicionales.")
