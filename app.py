# ============================================================
# Título: Visor Hidrológico CAMELS-PE en Streamlit
# Conversión base desde Shiny/R a Python/Streamlit
# - Mapa interactivo con ríos, estaciones y cuencas
# - Parámetros morfométricos
# - Índices climáticos y firmas hidrológicas desde Google Drive
# - Series temporales por gauge_id desde Google Drive
# - Descarga de Excel y SHP ZIP
# - Estaciones y cuencas desde GPKG local
# - Nombres de estaciones/cuencas en MAYÚSCULAS
# ============================================================

from __future__ import annotations

import io
import os
import zipfile
import tempfile
from pathlib import Path
from typing import Optional, Iterable

import folium
import geopandas as gpd
import pandas as pd
import plotly.express as px
import streamlit as st
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from streamlit_folium import st_folium

# ============================================================
# CONFIGURACIÓN GENERAL
# ============================================================

st.set_page_config(
    page_title="Visor Hidrológico CAMELS-PE",
    layout="wide",
    initial_sidebar_state="expanded",
)

SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]

# Carpeta donde están:
# landcover_attributes.csv
# climatic_indices.csv
# hydrological_signatures.csv
DRIVE_ATRIBUTOS_ID = "1g6gk3yhO_rXD2ydRS5Y6TBhzcv9GA9_J"

# Carpeta by_catchment donde están los CSV PE_<gauge_id>.csv
DRIVE_TIMESERIES_ID = "1rax6zd5iygqHK8tR1lzWcPMs6JrrobMe"

# Rutas locales dentro del repositorio
RUTA_GAUGES = Path("shp/camels_pe_gauges.gpkg")
LAYER_GAUGES = "camels_pe_gauges"

RUTA_CATCHMENTS = Path("shp/camels_pe_catchments.gpkg")
LAYER_CATCHMENTS = "camels_pe_catchments"

RUTA_LOGO = Path("logo.jpg")
RUTA_MANUAL = Path("Manual_de_uso.pdf")

# Nombres de columnas de series temporales
RUTA_TMAX = "tmax"
RUTA_TMED = "tmean"
RUTA_TMIN = "tmin"
RUTA_PREC = "prec"
RUTA_QOBS = "flow_obs"
RUTA_QSIM = "flow_sim"
RUTA_EVAP = "pet"
RUTA_SRAD = "srad"
RUTA_VPR = "vprp"

VARIABLES = [
    "Temperatura",
    "Caudales",
    "Precipitacion",
    "Evapotranspiracion",
    "Radiacion solar",
    "Presion de Vapor",
]

FRECUENCIAS = ["Diario", "Mensual", "Anual"]

NA_VALUES = ["", "NA", "NaN", "-9999", "-999.0", "-9999.0"]

# ============================================================
# FUNCIONES DE NORMALIZACIÓN
# ============================================================


def normalizar_gauge_pe(gauge_id) -> Optional[str]:
    if gauge_id is None or pd.isna(gauge_id):
        return None
    codigo = str(gauge_id).strip().upper()
    codigo = codigo.removeprefix("PE_")
    if codigo in {"", "NA", "NAN", "NONE"}:
        return None
    return f"PE_{codigo}"


def normalizar_codigo_cuenca(gauge_id) -> str:
    gid = normalizar_gauge_pe(gauge_id)
    if gid is None:
        return ""
    return gid.replace("PE_", "")


def obtener_columna_texto(df: pd.DataFrame, columnas: Iterable[str], fallback: str = "") -> pd.Series:
    for col in columnas:
        if col in df.columns:
            return df[col].astype(str)
    return pd.Series([fallback] * len(df), index=df.index)


def limpiar_nombres_columnas(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip().lower() for c in df.columns]
    return df


def tabla_param_val(diccionario: dict) -> pd.DataFrame:
    filas = []
    for parametro, valor in diccionario.items():
        if valor is None or pd.isna(valor):
            valor_fmt = ""
        elif isinstance(valor, (float, int)):
            valor_fmt = f"{valor:.3f}"
        else:
            valor_fmt = str(valor)
        filas.append({"Parámetro": parametro, "Valor": valor_fmt})
    return pd.DataFrame(filas)

# ============================================================
# GOOGLE DRIVE
# ============================================================


@st.cache_resource(show_spinner=False)
def obtener_servicio_drive():
    """Crea el cliente de Google Drive.

    En Streamlit Cloud se recomienda usar st.secrets.
    Para pruebas locales también puede usarse un archivo service_account.json,
    pero ese archivo NO debe subirse a GitHub.
    """
    if "gcp_service_account" in st.secrets:
        info = dict(st.secrets["gcp_service_account"])
        creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
    else:
        ruta_json = Path("service_account.json")
        if not ruta_json.exists():
            raise FileNotFoundError(
                "No se encontró la credencial. Usa .streamlit/secrets.toml "
                "o un archivo local service_account.json no subido a GitHub."
            )
        creds = service_account.Credentials.from_service_account_file(str(ruta_json), scopes=SCOPES)

    return build("drive", "v3", credentials=creds)


@st.cache_data(ttl=900, show_spinner=False)
def buscar_archivo_drive(folder_id: str, nombre_archivo: str) -> Optional[str]:
    service = obtener_servicio_drive()
    nombre_esc = nombre_archivo.replace("'", "\\'")
    query = (
        f"'{folder_id}' in parents and "
        f"name = '{nombre_esc}' and "
        "trashed = false"
    )
    resp = service.files().list(q=query, fields="files(id, name)", pageSize=10).execute()
    archivos = resp.get("files", [])
    if not archivos:
        return None
    return archivos[0]["id"]


@st.cache_data(ttl=900, show_spinner=False)
def descargar_archivo_drive(file_id: str) -> bytes:
    service = obtener_servicio_drive()
    request = service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    fh.seek(0)
    return fh.read()


def leer_csv_drive_directo(folder_id: str, nombre_archivo: str, detener_si_no_existe: bool = True) -> Optional[pd.DataFrame]:
    file_id = buscar_archivo_drive(folder_id, nombre_archivo)
    if file_id is None:
        mensaje = f"No se encontró el archivo en Google Drive: {nombre_archivo}"
        if detener_si_no_existe:
            raise FileNotFoundError(mensaje)
        st.warning(mensaje)
        return None

    contenido = descargar_archivo_drive(file_id)
    return pd.read_csv(io.BytesIO(contenido), na_values=NA_VALUES)


@st.cache_data(ttl=900, show_spinner=False)
def leer_landcover_drive() -> pd.DataFrame:
    df = leer_csv_drive_directo(DRIVE_ATRIBUTOS_ID, "landcover_attributes.csv")
    df = limpiar_nombres_columnas(df)
    if "gauge_id" not in df.columns:
        raise ValueError("El archivo landcover_attributes.csv no tiene la columna gauge_id.")
    df["gauge_id"] = df["gauge_id"].apply(normalizar_gauge_pe)
    return df


@st.cache_data(ttl=900, show_spinner=False)
def leer_indices_drive() -> pd.DataFrame:
    df = leer_csv_drive_directo(DRIVE_ATRIBUTOS_ID, "climatic_indices.csv")
    df = limpiar_nombres_columnas(df)
    if "gauge_id" not in df.columns:
        raise ValueError("El archivo climatic_indices.csv no tiene la columna gauge_id.")
    df["gauge_id"] = df["gauge_id"].apply(normalizar_gauge_pe)
    return df


@st.cache_data(ttl=900, show_spinner=False)
def leer_firmas_drive() -> pd.DataFrame:
    df = leer_csv_drive_directo(DRIVE_ATRIBUTOS_ID, "hydrological_signatures.csv")
    df = limpiar_nombres_columnas(df)
    if "gauge_id" not in df.columns:
        raise ValueError("El archivo hydrological_signatures.csv no tiene la columna gauge_id.")
    df["gauge_id"] = df["gauge_id"].apply(normalizar_gauge_pe)
    return df


@st.cache_data(ttl=900, show_spinner=False)
def leer_topografia_drive() -> pd.DataFrame:
    df = leer_csv_drive_directo(DRIVE_ATRIBUTOS_ID, "topographic_attributes.csv")
    df = limpiar_nombres_columnas(df)
    if "gauge_id" not in df.columns:
        raise ValueError("El archivo topographic_attributes.csv no tiene la columna gauge_id.")
    df["gauge_id"] = df["gauge_id"].apply(normalizar_gauge_pe)
    return df


@st.cache_data(ttl=900, show_spinner=False)
def leer_geologia_drive() -> pd.DataFrame:
    df = leer_csv_drive_directo(DRIVE_ATRIBUTOS_ID, "geologic_attributes.csv")
    df = limpiar_nombres_columnas(df)
    if "gauge_id" not in df.columns:
        raise ValueError("El archivo geologic_attributes.csv no tiene la columna gauge_id.")
    df["gauge_id"] = df["gauge_id"].apply(normalizar_gauge_pe)
    return df


@st.cache_data(ttl=900, show_spinner=False)
def leer_suelos_drive() -> pd.DataFrame:
    df = leer_csv_drive_directo(DRIVE_ATRIBUTOS_ID, "soil_attributes.csv")
    df = limpiar_nombres_columnas(df)
    if "gauge_id" not in df.columns:
        raise ValueError("El archivo soil_attributes.csv no tiene la columna gauge_id.")
    df["gauge_id"] = df["gauge_id"].apply(normalizar_gauge_pe)
    return df


@st.cache_data(ttl=900, show_spinner=False)
def leer_csv_timeseries_drive(gauge_id: str) -> Optional[pd.DataFrame]:
    codigo = normalizar_codigo_cuenca(gauge_id)
    nombre_csv = f"PE_{codigo}.csv"
    return leer_csv_drive_directo(DRIVE_TIMESERIES_ID, nombre_csv, detener_si_no_existe=False)

# ============================================================
# LECTURA DE GPKG / SHP LOCAL
# ============================================================


@st.cache_data(show_spinner="Cargando datos geoespaciales...")
def cargar_geodatos():
    if not RUTA_GAUGES.exists():
        raise FileNotFoundError(f"No existe: {RUTA_GAUGES}")
    if not RUTA_CATCHMENTS.exists():
        raise FileNotFoundError(f"No existe: {RUTA_CATCHMENTS}")
    estaciones = gpd.read_file(RUTA_GAUGES, layer=LAYER_GAUGES).to_crs(4326)
    cuencas = gpd.read_file(RUTA_CATCHMENTS, layer=LAYER_CATCHMENTS).to_crs(4326)

    estaciones["geometry"] = estaciones.geometry.make_valid()
    cuencas["geometry"] = cuencas.geometry.make_valid()

    estaciones["gauge_id"] = estaciones["gauge_id"].apply(normalizar_gauge_pe)
    cuencas["gauge_id"] = cuencas["gauge_id"].apply(normalizar_gauge_pe)

    estaciones["Estacion"] = obtener_columna_texto(
        estaciones, ["name", "Estacion", "estacion", "nombre"], ""
    ).str.strip().str.upper()
    cuencas["Estacion"] = obtener_columna_texto(
        cuencas, ["name", "Estacion", "estacion", "nombre"], ""
    ).str.strip().str.upper()

    estaciones.loc[estaciones["Estacion"].isin(["", "NA", "NAN"]), "Estacion"] = estaciones["gauge_id"]
    cuencas.loc[cuencas["Estacion"].isin(["", "NA", "NAN"]), "Estacion"] = cuencas["gauge_id"]

    # Región y Zonal desde el GPKG corregido.
    # IMPORTANTE: no usar name_cat como Region, porque name_cat contiene nombres de cuenca.
    columnas_lower = {str(col).strip().lower(): col for col in estaciones.columns}

    if "region" not in columnas_lower:
        raise ValueError(
            "El GPKG de gauges no tiene el campo Region. "
            "Usa camels_pe_gauges_con_region_zonal.gpkg o agrega ese campo antes de correr el visor."
        )

    if "zonal" not in columnas_lower:
        raise ValueError(
            "El GPKG de gauges no tiene el campo Zonal. "
            "Usa camels_pe_gauges_con_region_zonal.gpkg o agrega ese campo antes de correr el visor."
        )

    col_region = columnas_lower["region"]
    col_zonal = columnas_lower["zonal"]

    estaciones["Region"] = estaciones[col_region].astype(str).str.strip().str.upper()
    estaciones["Zonal"] = estaciones[col_zonal].astype(str).str.strip().str.upper()

    estaciones.loc[
        estaciones["Region"].isna()
        | estaciones["Region"].isin(["", "NA", "NAN", "NONE"]),
        "Region",
    ] = "SIN REGIÓN"

    estaciones.loc[
        estaciones["Zonal"].isna()
        | estaciones["Zonal"].isin(["", "NA", "NAN", "NONE"]),
        "Zonal",
    ] = "SIN ZONAL"

    estaciones = estaciones.sort_values(["Region", "Estacion", "gauge_id"]).reset_index(drop=True)

    return estaciones, cuencas

# ============================================================
# SERIES TEMPORALES
# ============================================================


def parsear_fecha_timeseries(serie: pd.Series) -> pd.Series:
    if pd.api.types.is_datetime64_any_dtype(serie):
        return pd.to_datetime(serie).dt.date

    if pd.api.types.is_numeric_dtype(serie):
        return pd.to_datetime(serie, origin="1899-12-30", unit="D", errors="coerce").dt.date

    texto = serie.astype(str).str.strip()
    fecha = pd.to_datetime(texto, errors="coerce", dayfirst=False)
    faltantes = fecha.isna()
    if faltantes.any():
        fecha.loc[faltantes] = pd.to_datetime(texto.loc[faltantes], errors="coerce", dayfirst=True)
    return fecha.dt.date


def extraer_serie_desde_df(
    dat: Optional[pd.DataFrame],
    columna_objetivo: str,
    fecha_ini,
    fecha_fin,
) -> Optional[pd.DataFrame]:
    if dat is None or dat.empty:
        return None

    df = limpiar_nombres_columnas(dat)
    columna_objetivo = columna_objetivo.strip().lower()

    if "date" in df.columns:
        col_fecha = "date"
    elif "fecha" in df.columns:
        col_fecha = "fecha"
    else:
        return None

    if columna_objetivo not in df.columns:
        return None

    out = pd.DataFrame({
        "Fecha": pd.to_datetime(parsear_fecha_timeseries(df[col_fecha])),
        "Valor": pd.to_numeric(df[columna_objetivo], errors="coerce"),
    })

    fecha_ini = pd.to_datetime(fecha_ini)
    fecha_fin = pd.to_datetime(fecha_fin)

    out = out.dropna(subset=["Fecha"])
    out = out[(out["Fecha"] >= fecha_ini) & (out["Fecha"] <= fecha_fin)]
    return out


def renombrar_serie_safe(df: Optional[pd.DataFrame], nombre_col: str) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame({"Fecha": pd.to_datetime([]), nombre_col: pd.Series(dtype="float")})
    return df.rename(columns={"Valor": nombre_col})


def agregar_serie(df: Optional[pd.DataFrame], frecuencia: str, modo: str = "mean") -> Optional[pd.DataFrame]:
    if df is None or df.empty:
        return df
    if frecuencia == "Diario":
        return df

    regla = "MS" if frecuencia == "Mensual" else "YS"
    df2 = df.copy()
    df2["Fecha"] = pd.to_datetime(df2["Fecha"])
    df2 = df2.set_index("Fecha")

    if modo == "sum":
        out = df2.resample(regla)["Valor"].sum(min_count=1).reset_index()
    else:
        out = df2.resample(regla)["Valor"].mean().reset_index()
    return out


def agregar_cortes_por_vacios(df: pd.DataFrame, frecuencia: str) -> pd.DataFrame:
    """Evita que Plotly una con una línea recta periodos sin datos."""
    if df is None or df.empty or "Fecha" not in df.columns:
        return df

    datos = df.copy()
    datos["Fecha"] = pd.to_datetime(datos["Fecha"])

    if "Variable" not in datos.columns:
        datos["Variable"] = "Serie"

    max_gap = {
        "Diario": pd.Timedelta(days=2),
        "Mensual": pd.Timedelta(days=40),
        "Anual": pd.Timedelta(days=370),
    }.get(frecuencia, pd.Timedelta(days=2))

    partes = []
    for variable, grupo in datos.groupby("Variable", dropna=False):
        grupo = grupo.sort_values("Fecha").copy()
        grupo["_gap"] = grupo["Fecha"].diff().gt(max_gap).fillna(False)
        grupo["_segmento_id"] = grupo["_gap"].cumsum()
        grupo["Segmento"] = grupo["Variable"].astype(str) + "_" + grupo["_segmento_id"].astype(str)
        partes.append(grupo.drop(columns=["_gap", "_segmento_id"]))

    return pd.concat(partes, ignore_index=True)


def tabla_desde_excel_por_gauge(df_excel: pd.DataFrame, gid: str, titulo_no_registro: str) -> pd.DataFrame:
    gid_pe = normalizar_gauge_pe(gid)
    fila = df_excel[df_excel["gauge_id"] == gid_pe]
    if fila.empty:
        return pd.DataFrame({
            "Parámetro": [titulo_no_registro],
            "Valor": [f"No hay registro para el gauge_id {gid_pe}"],
        })

    fila2 = fila.drop(columns=["gauge_id"], errors="ignore").iloc[0]
    return tabla_param_val(fila2.to_dict())

# ============================================================
# CÁLCULOS MORFOMÉTRICOS
# ============================================================


def proyectar_metrico(gdf: gpd.GeoDataFrame) -> gpd.GeoDataFrame:
    if gdf.empty:
        return gdf
    try:
        crs_metrico = gdf.estimate_utm_crs()
        if crs_metrico is None:
            crs_metrico = "EPSG:32718"
    except Exception:
        crs_metrico = "EPSG:32718"
    return gdf.to_crs(crs_metrico)


def tabla_topografica_por_gauge(df_topografia: pd.DataFrame, gid: str) -> pd.DataFrame:
    gid_pe = normalizar_gauge_pe(gid)
    fila = df_topografia[df_topografia["gauge_id"] == gid_pe]

    if fila.empty:
        return pd.DataFrame({
            "Parámetro": ["Parámetros morfométricos"],
            "Valor": [f"No hay registro para el gauge_id {gid_pe}"],
        })

    fila = fila.iloc[0]

    nombres = {
        "area": "Área",
        "perimeter": "Perímetro",
        "elev_min": "Elevación mínima",
        "elev_max": "Elevación máxima",
        "elev_mean": "Elevación media",
        "elev_median": "Elevación mediana",
        "slope_mean": "Pendiente media",
    }

    valores = {}
    for columna, nombre in nombres.items():
        if columna in df_topografia.columns:
            valores[nombre] = pd.to_numeric(fila[columna], errors="coerce")

    return tabla_param_val(valores)


def preparar_cobertura_landcover(df_landcover: pd.DataFrame, gid: str) -> Optional[pd.DataFrame]:
    gid_pe = normalizar_gauge_pe(gid)
    fila = df_landcover[df_landcover["gauge_id"] == gid_pe]
    if fila.empty:
        return None

    columnas_cobertura = [
        "agricul_perc",
        "forest_perc",
        "non_vaget_perc",
        "non_veget_perc",
        "non_woody_perc",
        "water_perc",
        "non_forest_perc",
        "non_identi_perc",
    ]
    columnas_cobertura = [c for c in columnas_cobertura if c in fila.columns]
    if not columnas_cobertura:
        return None

    nombres = {
        "agricul_perc": "Agrícola",
        "forest_perc": "Bosque",
        "non_vaget_perc": "Sin vegetación",
        "non_veget_perc": "Sin vegetación",
        "non_woody_perc": "Vegetación no leñosa",
        "water_perc": "Agua",
        "non_forest_perc": "No bosque",
        "non_identi_perc": "No identificado",
    }

    tabla = (
        fila[columnas_cobertura]
        .melt(var_name="variable", value_name="porcentaje")
        .assign(
            Categoria=lambda x: x["variable"].map(nombres),
            porcentaje=lambda x: pd.to_numeric(x["porcentaje"], errors="coerce"),
        )
        .dropna(subset=["porcentaje"])
    )
    tabla = tabla[tabla["porcentaje"] > 0]
    tabla = tabla[["Categoria", "porcentaje"]].sort_values("porcentaje", ascending=False)
    return tabla if not tabla.empty else None


def preparar_geologia(df_geologia: pd.DataFrame, gid: str) -> Optional[pd.DataFrame]:
    gid_pe = normalizar_gauge_pe(gid)
    fila = df_geologia[df_geologia["gauge_id"] == gid_pe]
    if fila.empty:
        return None

    columnas = {
        "geol_class_1st_per": "Geología clase principal",
        "geol_class_2nd_per": "Geología clase secundaria",
        "inter_volca_rocks_pe": "Rocas volcánicas intermedias",
        "geol_porosity": "Porosidad geológica",
        "geol_permeability": "Permeabilidad geológica",
    }
    columnas = {col: nombre for col, nombre in columnas.items() if col in fila.columns}
    if not columnas:
        return None

    tabla = (
        fila[list(columnas.keys())]
        .melt(var_name="variable", value_name="valor")
        .assign(
            Categoria=lambda x: x["variable"].map(columnas),
            valor=lambda x: pd.to_numeric(x["valor"], errors="coerce"),
        )
        .dropna(subset=["valor"])
    )
    tabla = tabla[["Categoria", "valor"]].sort_values("valor", ascending=False)
    return tabla if not tabla.empty else None


def preparar_suelos(df_suelos: pd.DataFrame, gid: str) -> Optional[pd.DataFrame]:
    gid_pe = normalizar_gauge_pe(gid)
    fila = df_suelos[df_suelos["gauge_id"] == gid_pe]
    if fila.empty:
        return None

    columnas = [
        "inceptisols_perc",
        "entisols_perc",
        "alfisols_perc",
        "ultisols_perc",
        "aridisols_perc",
        "gelisols_perc",
        "oxisols_perc",
        "mollisols_perc",
        "vertisols_perc",
    ]
    columnas = [c for c in columnas if c in fila.columns]
    if not columnas:
        return None

    nombres = {
        "inceptisols_perc": "Inceptisols",
        "entisols_perc": "Entisols",
        "alfisols_perc": "Alfisols",
        "ultisols_perc": "Ultisols",
        "aridisols_perc": "Aridisols",
        "gelisols_perc": "Gelisols",
        "oxisols_perc": "Oxisols",
        "mollisols_perc": "Mollisols",
        "vertisols_perc": "Vertisols",
    }

    tabla = (
        fila[columnas]
        .melt(var_name="variable", value_name="porcentaje")
        .assign(
            Categoria=lambda x: x["variable"].map(nombres),
            porcentaje=lambda x: pd.to_numeric(x["porcentaje"], errors="coerce"),
        )
        .dropna(subset=["porcentaje"])
    )
    tabla = tabla[tabla["porcentaje"] > 0]
    tabla = tabla[["Categoria", "porcentaje"]].sort_values("porcentaje", ascending=False)
    return tabla if not tabla.empty else None


def grafico_barras_atributos(datos: Optional[pd.DataFrame], titulo: str, columna_valor: str, color: str):
    if datos is None or datos.empty:
        return None

    fig = px.bar(
        datos,
        x=columna_valor,
        y="Categoria",
        orientation="h",
        text=datos[columna_valor].map(lambda x: f"{x:.2f}"),
        labels={columna_valor: "Valor", "Categoria": ""},
        title=titulo,
    )
    fig.update_traces(marker_color=color, textposition="outside")
    fig.update_layout(
        height=380,
        yaxis={"categoryorder": "total ascending"},
        margin=dict(l=10, r=10, t=45, b=10),
        showlegend=False,
    )
    return fig

# ============================================================
# MAPA
# ============================================================


def limpiar_poligonos_para_mapa(gdf: gpd.GeoDataFrame) -> gpd.GeoDataFrame:
    """Limpia geometrías antes de enviarlas a Folium.

    Algunas cuencas pueden convertirse en GeometryCollection después de make_valid().
    Folium puede fallar con esas geometrías porque no siempre tienen la clave
    coordinates en el GeoJSON. Por eso se conservan solo Polygon/MultiPolygon.
    """
    if gdf is None or gdf.empty:
        return gpd.GeoDataFrame(columns=["geometry"], geometry="geometry", crs="EPSG:4326")

    salida = gdf.copy()
    salida = salida[salida.geometry.notna()].copy()
    salida = salida[~salida.geometry.is_empty].copy()

    if salida.empty:
        return salida

    salida["geometry"] = salida.geometry.make_valid()
    salida = salida.explode(index_parts=False).reset_index(drop=True)
    salida = salida[salida.geometry.geom_type.isin(["Polygon", "MultiPolygon"])].copy()
    salida = salida[salida.geometry.notna()].copy()
    salida = salida[~salida.geometry.is_empty].copy()

    return salida


def crear_mapa(estaciones: gpd.GeoDataFrame, cuenca: gpd.GeoDataFrame) -> folium.Map:
    cuenca_mapa = limpiar_poligonos_para_mapa(cuenca)

    if not cuenca_mapa.empty:
        centroide = cuenca_mapa.geometry.union_all().centroid
        centro = [centroide.y, centroide.x]
        zoom = 9
    else:
        centro = [-9.19, -75.02]
        zoom = 5

    mapa = folium.Map(
        location=centro,
        zoom_start=zoom,
        tiles="OpenStreetMap",
        control_scale=True,
    )

    if not cuenca_mapa.empty:
        campos_tooltip = [c for c in ["Estacion", "gauge_id"] if c in cuenca_mapa.columns]
        tooltip = folium.GeoJsonTooltip(fields=campos_tooltip) if campos_tooltip else None

        folium.GeoJson(
            cuenca_mapa.to_json(),
            name="Cuenca",
            tooltip=tooltip,
            style_function=lambda _: {
                "color": "black",
                "weight": 2,
                "fillColor": "orange",
                "fillOpacity": 0.4,
            },
        ).add_to(mapa)

        xmin, ymin, xmax, ymax = cuenca_mapa.total_bounds
        mapa.fit_bounds([[ymin, xmin], [ymax, xmax]])

    estaciones_pts = estaciones.copy()
    for _, row in estaciones_pts.iterrows():
        if row.geometry is None or row.geometry.is_empty:
            continue
        folium.CircleMarker(
            location=[row.geometry.y, row.geometry.x],
            radius=5,
            color="blue",
            fill=True,
            fill_color="blue",
            fill_opacity=0.9,
            tooltip=f"{row['Estacion']} ({row['gauge_id']})",
        ).add_to(mapa)

    return mapa

# ============================================================
# GRÁFICOS Y DESCARGAS
# ============================================================


def construir_datos_grafico(dat_drive: pd.DataFrame, tipo: str, frecuencia: str, fecha_ini, fecha_fin):
    if tipo == "Temperatura":
        tmin = extraer_serie_desde_df(dat_drive, RUTA_TMIN, fecha_ini, fecha_fin)
        tmed = extraer_serie_desde_df(dat_drive, RUTA_TMED, fecha_ini, fecha_fin)
        tmax = extraer_serie_desde_df(dat_drive, RUTA_TMAX, fecha_ini, fecha_fin)

        df = renombrar_serie_safe(tmin, "Tmin")
        df = df.merge(renombrar_serie_safe(tmed, "Tmed"), on="Fecha", how="outer")
        df = df.merge(renombrar_serie_safe(tmax, "Tmax"), on="Fecha", how="outer")

        if frecuencia != "Diario":
            regla = "MS" if frecuencia == "Mensual" else "YS"
            df = df.set_index("Fecha").resample(regla).mean().reset_index()

        largo = df.melt(id_vars="Fecha", var_name="Variable", value_name="Valor")
        largo = largo.dropna(subset=["Valor"])
        largo = agregar_cortes_por_vacios(largo, frecuencia)
        return largo, "line", "°C"

    if tipo == "Caudales":
        qobs = extraer_serie_desde_df(dat_drive, RUTA_QOBS, fecha_ini, fecha_fin)
        qsim = extraer_serie_desde_df(dat_drive, RUTA_QSIM, fecha_ini, fecha_fin)

        df = renombrar_serie_safe(qobs, "Observado")
        df = df.merge(renombrar_serie_safe(qsim, "Simulado"), on="Fecha", how="outer")

        if frecuencia != "Diario":
            regla = "MS" if frecuencia == "Mensual" else "YS"
            df = df.set_index("Fecha").resample(regla).mean().reset_index()

        largo = df.melt(id_vars="Fecha", var_name="Variable", value_name="Valor")
        largo = largo.dropna(subset=["Valor"])
        largo = agregar_cortes_por_vacios(largo, frecuencia)
        return largo, "line", "mm/día"

    config = {
        "Precipitacion": (RUTA_PREC, "sum", "bar", "mm"),
        "Evapotranspiracion": (RUTA_EVAP, "mean", "line", "mm/día"),
        "Radiacion solar": (RUTA_SRAD, "sum", "line", "MJ/m²"),
        "Presion de Vapor": (RUTA_VPR, "mean", "line", "hPa"),
    }
    columna, modo, tipo_grafico, unidad = config[tipo]
    serie = extraer_serie_desde_df(dat_drive, columna, fecha_ini, fecha_fin)
    serie = agregar_serie(serie, frecuencia, modo=modo)
    if serie is None:
        return pd.DataFrame(), tipo_grafico, unidad
    serie = serie.dropna(subset=["Valor"])
    if tipo_grafico == "line":
        serie = agregar_cortes_por_vacios(serie, frecuencia)
    return serie, tipo_grafico, unidad


def generar_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Datos")
    output.seek(0)
    return output.read()


def generar_zip_shapefile(cuenca: gpd.GeoDataFrame, gauge_id: str) -> bytes:
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir_path = Path(tmpdir)
        nombre_base = f"Cuenca_{normalizar_codigo_cuenca(gauge_id)}"
        ruta_shp = tmpdir_path / f"{nombre_base}.shp"
        cuenca.to_file(ruta_shp, driver="ESRI Shapefile")

        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for archivo in tmpdir_path.glob(f"{nombre_base}.*"):
                zf.write(archivo, arcname=archivo.name)
        buffer.seek(0)
        return buffer.read()

# ============================================================
# APP STREAMLIT
# ============================================================

st.title("Visor Hidrológico Del Peru")

try:
    estaciones, cuencas_gpkg = cargar_geodatos()
except Exception as e:
    st.error(f"Error al cargar datos geoespaciales: {e}")
    st.stop()

with st.sidebar:
    if RUTA_LOGO.exists():
        st.image(str(RUTA_LOGO), use_container_width=True)

    regiones = ["Todas las regiones"] + sorted(estaciones["Region"].dropna().unique().tolist())
    region_input = st.selectbox("Región:", regiones)

    if region_input == "Todas las regiones":
        estaciones_filtradas = estaciones.copy()
    else:
        estaciones_filtradas = estaciones[estaciones["Region"] == region_input].copy()

    # Orden real por Región y luego por Estación.
    # Esto evita que el selector se ordene por nombre de cuenca cuando existe name_cat.
    estaciones_filtradas = estaciones_filtradas.sort_values(["Region", "Estacion", "gauge_id"])
    opciones_gauge = estaciones_filtradas["gauge_id"].dropna().tolist()

    def etiqueta_gauge(gid: str) -> str:
        fila = estaciones[estaciones["gauge_id"] == gid]
        if fila.empty:
            return gid
        estacion = fila.iloc[0]["Estacion"]
        return f"{estacion} - {gid}"

    if not opciones_gauge:
        st.warning("No hay estaciones para la región seleccionada.")
        st.stop()

    cuenca_input = st.selectbox("Estación:", opciones_gauge, format_func=etiqueta_gauge)

    tipo_dato = st.selectbox("Variable:", VARIABLES)
    frecuencia = st.selectbox("Frecuencia:", FRECUENCIAS)

    rango_fechas = st.date_input(
        "Rango de fechas:",
        value=(pd.to_datetime("1981-01-01"), pd.to_datetime("2025-12-31")),
        min_value=pd.to_datetime("1981-01-01"),
        max_value=pd.to_datetime("2025-12-31"),
        format="DD/MM/YYYY",
    )


    if RUTA_MANUAL.exists():
        st.download_button(
            "📘 Descargar manual de uso",
            data=RUTA_MANUAL.read_bytes(),
            file_name="Manual_de_uso.pdf",
            mime="application/pdf",
        )

    st.markdown("### Descargas")
    descargar_datos_slot = st.empty()
    descargar_cuenca_slot = st.empty()


if isinstance(rango_fechas, tuple) and len(rango_fechas) == 2:
    fecha_ini, fecha_fin = rango_fechas
else:
    fecha_ini, fecha_fin = pd.to_datetime("1981-01-01"), pd.to_datetime("2025-12-31")

cuenca_sel = cuencas_gpkg[cuencas_gpkg["gauge_id"] == normalizar_gauge_pe(cuenca_input)].copy()
if cuenca_sel.empty:
    st.error(f"No se encontró cuenca en el GPKG para: {cuenca_input}")
    st.stop()

cuenca_sel["Estacion"] = cuenca_sel["Estacion"].astype(str).str.strip().str.upper()
cuenca_sel.loc[cuenca_sel["Estacion"].isin(["", "NA", "NAN"]), "Estacion"] = normalizar_gauge_pe(cuenca_input)

col_mapa, col_tablas = st.columns([1.35, 1], gap="large")

with col_mapa:
    st.subheader("Mapa interactivo")
    mapa = crear_mapa(estaciones, cuenca_sel)
    st_folium(mapa, height=620, use_container_width=True)

    st.subheader("Atributos de la cuenca")

    try:
        landcover = leer_landcover_drive()
        cobertura = preparar_cobertura_landcover(landcover, cuenca_input)
        fig_cov = grafico_barras_atributos(cobertura, "Cobertura de suelo (%)", "porcentaje", "#2E7D32")
        if fig_cov is None:
            st.info("Sin datos de cobertura.")
        else:
            st.plotly_chart(fig_cov, use_container_width=True)
    except Exception as e:
        st.warning(f"No se pudo leer landcover desde Drive: {e}")

    try:
        geologia = leer_geologia_drive()
        datos_geo = preparar_geologia(geologia, cuenca_input)
        fig_geo = grafico_barras_atributos(datos_geo, "Atributos geológicos", "valor", "#6D4C41")
        if fig_geo is None:
            st.info("Sin datos geológicos.")
        else:
            st.plotly_chart(fig_geo, use_container_width=True)
    except Exception as e:
        st.warning(f"No se pudo leer geologic desde Drive: {e}")

    try:
        suelos = leer_suelos_drive()
        datos_suelo = preparar_suelos(suelos, cuenca_input)
        fig_suelo = grafico_barras_atributos(datos_suelo, "Atributos de suelo (%)", "porcentaje", "#1565C0")
        if fig_suelo is None:
            st.info("Sin datos de suelo.")
        else:
            st.plotly_chart(fig_suelo, use_container_width=True)
    except Exception as e:
        st.warning(f"No se pudo leer soil desde Drive: {e}")

    st.subheader("Serie temporal")
    try:
        dat_drive = leer_csv_timeseries_drive(cuenca_input)
    except Exception as e:
        st.warning(f"No se pudo leer la serie temporal desde Drive: {e}")
        dat_drive = None

    if dat_drive is None:
        st.info("No se encontró el CSV de series temporales para este gauge_id.")
        datos_grafico = pd.DataFrame()
    else:
        datos_grafico, tipo_grafico, unidad = construir_datos_grafico(
            dat_drive, tipo_dato, frecuencia, fecha_ini, fecha_fin
        )

        if datos_grafico.empty:
            st.info("No hay datos para la variable, estación o rango de fechas seleccionado.")
        else:
            titulo = f"{tipo_dato} ({frecuencia}) - Estación: {etiqueta_gauge(cuenca_input)}"
            if "Variable" in datos_grafico.columns:
                colores_series = {
                    "Observado": "red",
                    "Simulado": "blue",
                    "Tmin": "blue",
                    "Tmed": "red",
                    "Tmax": "orange",
                }
                fig = px.line(
                    datos_grafico,
                    x="Fecha",
                    y="Valor",
                    color="Variable",
                    line_group="Segmento" if "Segmento" in datos_grafico.columns else None,
                    color_discrete_map=colores_series,
                    title=titulo,
                    labels={"Valor": unidad, "Fecha": "Fecha"},
                )
            elif tipo_grafico == "bar":
                fig = px.bar(
                    datos_grafico,
                    x="Fecha",
                    y="Valor",
                    title=titulo,
                    labels={"Valor": unidad, "Fecha": "Fecha"},
                )
            else:
                fig = px.line(
                    datos_grafico,
                    x="Fecha",
                    y="Valor",
                    line_group="Segmento" if "Segmento" in datos_grafico.columns else None,
                    title=titulo,
                    labels={"Valor": unidad, "Fecha": "Fecha"},
                )
                fig.update_traces(line=dict(color="blue"))
            fig.update_layout(height=360, margin=dict(l=10, r=10, t=55, b=10))
            st.plotly_chart(fig, use_container_width=True)

with col_tablas:
    st.subheader("Parámetros morfométricos")
    try:
        topografia = leer_topografia_drive()
        tabla_parametros = tabla_topografica_por_gauge(topografia, cuenca_input)
        st.dataframe(tabla_parametros, use_container_width=True, hide_index=True)
    except Exception as e:
        st.warning(f"No se pudieron leer los parámetros morfométricos desde Drive: {e}")

    with st.expander("Índices climáticos", expanded=True):
        try:
            indices = leer_indices_drive()
            tabla_indices = tabla_desde_excel_por_gauge(indices, cuenca_input, "Índices climáticos")
            st.dataframe(tabla_indices, use_container_width=True, hide_index=True)
        except Exception as e:
            st.warning(f"No se pudo leer índices desde Drive: {e}")

    with st.expander("Firmas hidrológicas", expanded=True):
        try:
            firmas = leer_firmas_drive()
            tabla_firmas = tabla_desde_excel_por_gauge(firmas, cuenca_input, "Firmas hidrológicas")
            st.dataframe(tabla_firmas, use_container_width=True, hide_index=True)
        except Exception as e:
            st.warning(f"No se pudo leer firmas desde Drive: {e}")

with descargar_datos_slot:
    if dat_drive is not None and not datos_grafico.empty:
        st.download_button(
            "📥 Descargar datos en Excel",
            data=generar_excel_bytes(datos_grafico),
            file_name=f"Datos_{cuenca_input}_{tipo_dato}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    else:
        st.button("📥 Descargar datos en Excel", disabled=True, use_container_width=True)

with descargar_cuenca_slot:
    try:
        zip_bytes = generar_zip_shapefile(cuenca_sel, cuenca_input)
        st.download_button(
            "🗺️ Descargar cuenca (SHP ZIP)",
            data=zip_bytes,
            file_name=f"Cuenca_{normalizar_codigo_cuenca(cuenca_input)}.zip",
            mime="application/zip",
            use_container_width=True,
        )
    except Exception as e:
        st.warning(f"No se pudo preparar la descarga SHP: {e}")

