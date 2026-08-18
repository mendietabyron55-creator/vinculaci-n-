import base64
import re
import unicodedata
import pandas as pd
import numpy as np
import plotly.express as px
import streamlit as st
from pathlib import Path

# ----------------------------------------------------------------------------
# Configuración general
# ----------------------------------------------------------------------------
st.set_page_config(
    page_title="Observatorio - Movilidad Humana",
    page_icon="🌎",
    layout="wide",
)

DATA_PATH = Path(__file__).parent / "REPORTE_CONSOLIDADO_MOVILIDAD_HUMANA_FINAL.xlsx"
DATA_CACHE_PATH = DATA_PATH.with_suffix(".cache.pkl")
DATA_CACHE_VERSION = 5

# Rutas de los logos (ajusta el nombre/extensión si es necesario)
LOGO1_PATH = Path(__file__).parent / "logo1.jpeg"   # Fundación Mensajeros de la Paz
LOGO2_PATH = Path(__file__).parent / "logo2.jpeg"   # Tec.Azuay

SI_NO_COLS = [
    "atencion_emergente", "kit_aseo", "kit_salud", "kit_escolar",
    "enfermedad_catastrofica", "tiene_discapacidad", "embarazo", "estudiando",
    "atencion_trabajo_social", "atencion_psicologica", "atencion_legal",
    "serv_salud", "serv_educacion", "serv_junta_cantonal",
    "serv_reunificacion_familiar", "serv_eti", "serv_acogimiento_institucional",
    "serv_apoyo_custodia_familiar", "serv_discapacidades", "serv_adulto_mayor",
    "serv_cdi", "serv_cnh", "part_talleres_capacitacion",
    "part_talleres_sensibilizacion", "part_encuentros_comunitarios",
    "part_talleres_nna", "part_redes_comunitarias",
]

FRIENDLY_NAMES = {
    "tiene_discapacidad": "Tiene discapacidad",
    "embarazo": "Embarazo",
    "enfermedad_catastrofica": "Enfermedad catastrófica",
    "atencion_trabajo_social": "Trabajo social",
    "atencion_psicologica": "Atención psicológica",
    "atencion_legal": "Atención legal",
    "serv_salud": "Salud",
    "serv_educacion": "Educación",
    "serv_junta_cantonal": "Junta cantonal",
    "serv_reunificacion_familiar": "Reunificación familiar",
    "serv_eti": "ETI",
    "serv_acogimiento_institucional": "Acogimiento institucional",
    "serv_apoyo_custodia_familiar": "Apoyo y custodia familiar",
    "serv_discapacidades": "Discapacidades",
    "serv_adulto_mayor": "Adulto mayor",
    "serv_cdi": "CDI",
    "serv_cnh": "CNH",
    "kit_aseo": "Kit de aseo",
    "kit_salud": "Kit de salud",
    "kit_escolar": "Kit escolar",
    "part_talleres_capacitacion": "Talleres de capacitación",
    "part_talleres_sensibilizacion": "Talleres de sensibilización",
    "part_encuentros_comunitarios": "Encuentros comunitarios",
    "part_talleres_nna": "Talleres lúdicos NNA",
    "part_redes_comunitarias": "Inserción a redes comunitarias",
}

# Tarjetas de categoría que se muestran en la portada (icono opcional + texto)
PORTADA_CARDS = [
    "Perfil de la Población",
    "Vulnerabilidades",
    "Intervenciones Técnicas",
    "Situación Migratoria",
    "Asistencia Humanitaria",
    "Integración Comunitaria",
]

# ----------------------------------------------------------------------------
# Paleta de color de las gráficas
# ----------------------------------------------------------------------------
# 8 tonos categóricos en orden fijo, validados para que sean distinguibles bajo
# daltonismo y tengan contraste suficiente. El orden nunca cambia: cada valor
# de una dimensión siempre recibe el mismo color, sin importar el ranking o el
# filtro activo (evita "repintar" categorías cuando cambian los datos).
COLOR_CATEGORICAL = [
    "#2a78d6", "#eb6834", "#1baf7a", "#eda100",
    "#e87ba4", "#008300", "#4a3aa7", "#e34948",
]
COLOR_PRIMARY = COLOR_CATEGORICAL[0]

# Rango de edad es una variable ordinal (bandas etarias): un solo tono de
# claro a oscuro, del grupo más joven al más longevo, en vez de colores
# categóricos sin relación entre sí.
COLOR_RANGO_EDAD = {
    "NN": "#86b6ef",
    "ADOLESCENTE": "#3987e5",
    "ADULTO": "#1c5cab",
    "ADULTO MAYOR": "#0d366b",
}
COLOR_GENERO = {
    "FEMENINO": COLOR_CATEGORICAL[0],
    "MASCULINO": COLOR_CATEGORICAL[1],
    "LGTBIQ+": COLOR_CATEGORICAL[2],
}
COLOR_SITUACION_MIGRATORIA = {
    "IRREGULAR": COLOR_CATEGORICAL[0],
    "REGULAR": COLOR_CATEGORICAL[1],
    "NO APLICA": COLOR_CATEGORICAL[2],
}
COLOR_FORMA_INGRESO = {
    "POR PASO IRREGULAR": COLOR_CATEGORICAL[0],
    "POR PASO REGULAR": COLOR_CATEGORICAL[1],
    "NO APLICA": COLOR_CATEGORICAL[2],
}
COLOR_SITUACION_MOVILIDAD = {
    "VOCACION DE PERMANENCIA": COLOR_CATEGORICAL[0],
    "EN TRANSITO": COLOR_CATEGORICAL[1],
    "NO APLICA": COLOR_CATEGORICAL[2],
}

# plot_si_bars se usa en 4 secciones distintas (Vulnerabilidades, Intervenciones
# Tecnicas, Asistencia Humanitaria, Integracion Comunitaria). Dentro de cada
# grafica los indicadores son nominales (que servicio, no un ranking), asi que
# llevan un solo tono; para que cada seccion se distinga de las demas, cada una
# usa un slot distinto de la paleta, siempre en el mismo orden fijo. Perfil de
# la Poblacion y Situacion Migratoria no tienen un color propio en sus graficas
# (son multicolor), asi que para la tarjeta de portada usan dos slots libres.
COLOR_SECCIONES = {
    "Perfil de la Población": COLOR_CATEGORICAL[5],
    "Vulnerabilidades": COLOR_CATEGORICAL[0],
    "Intervenciones Técnicas": COLOR_CATEGORICAL[1],
    "Situación Migratoria": COLOR_CATEGORICAL[6],
    "Asistencia Humanitaria": COLOR_CATEGORICAL[2],
    "Integración Comunitaria": COLOR_CATEGORICAL[3],
}

# Icono por seccion, mostrado en las tarjetas de la portada.
PORTADA_ICONS = {
    "Perfil de la Población": "👥",
    "Vulnerabilidades": "⚠️",
    "Intervenciones Técnicas": "🩺",
    "Situación Migratoria": "🛂",
    "Asistencia Humanitaria": "🎁",
    "Integración Comunitaria": "🤝",
}

CHART_FONT = "system-ui, -apple-system, 'Segoe UI', sans-serif"


def _style_chart(fig, legend: bool = False) -> None:
    """Aplica el mismo fondo, tipografía y líneas de apoyo a cualquier figura."""
    tiene_titulo = bool(fig.layout.title and fig.layout.title.text)
    fig.update_layout(
        plot_bgcolor="#fcfcfb",
        paper_bgcolor="#fcfcfb",
        font=dict(family=CHART_FONT, color="#0b0b0b", size=13),
        margin=dict(t=55 if tiene_titulo else 20, l=10, r=20, b=10),
        showlegend=legend,
        legend=dict(bgcolor="rgba(0,0,0,0)", title_text=""),
        bargap=0.35,
    )
    # title_font solo se fija si hay titulo: en algunas versiones de Plotly.js
    # fijar el font de un titulo vacio lo renderiza como el texto literal
    # "undefined" en vez de no mostrar nada.
    if tiene_titulo:
        fig.update_layout(title_font=dict(size=16, color="#0b0b0b"))
    fig.update_xaxes(gridcolor="#e1e0d9", linecolor="#c3c2b7", tickfont=dict(color="#898781"), zeroline=False)
    fig.update_yaxes(gridcolor="#e1e0d9", linecolor="#c3c2b7", tickfont=dict(color="#898781"), zeroline=False)


def _headroom(max_value: float, factor: float = 1.15) -> float:
    """Deja espacio suficiente para que la etiqueta de valor no se recorte."""
    return max(max_value * factor, 1)


def load_data() -> pd.DataFrame:
    if not DATA_PATH.exists():
        st.error(f"No se encontró el archivo de datos: {DATA_PATH.name}")
        st.stop()

    cached = _load_prepared_cache()
    if cached is not None:
        return cached

    try:
        df = _read_source_excel()
    except ImportError:
        st.error("Falta instalar 'openpyxl' para leer el archivo Excel. Ejecuta: pip install openpyxl")
        st.stop()

    prepared = _prepare_excel_data(df)
    _write_prepared_cache(prepared)
    return prepared


def si_pct(series: pd.Series) -> float:
    """% de 'SI' sobre los valores no nulos de una columna SI/NO."""
    s = series.dropna()
    if len(s) == 0:
        return 0.0
    return 100 * (s == "SI").sum() / len(s)


def _img_to_base64(path: Path) -> str | None:
    """Convierte una imagen local a base64 para poder incrustarla en HTML."""
    if not path.exists():
        return None
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode()


MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}

MESES_NUM = {v.upper(): k for k, v in MESES_ES.items()}

# Solo se leen del Excel las columnas listadas aquí. Cualquier columna que no
# aparezca (dirección, teléfonos, correos, documentos de identificación,
# contactos de madre/padre/tutor, código único de NNA, etc.) se descarta antes
# de cargarse en memoria: nunca llega a la aplicación.
#
# NOMBRES/APELLIDOS/FECHA_DE_NACIMIENTO son la única excepción: se leen solo
# para poder agrupar los registros de una misma persona atendida en distintos
# meses (ver _derive_persona_id). _prepare_excel_data las descarta antes de
# devolver los datos, así que nunca se guardan en la caché ni se muestran en
# ninguna pantalla de la aplicación.
COLUMN_MAP = {
    "ZONA": "zona",
    "PROVINCIA": "provincia",
    "DISTRITO": "distrito",
    "CIUDAD": "ciudad",
    "NOMBRES": "nombres_raw",
    "APELLIDOS": "apellidos_raw",
    "FECHA_DE_NACIMIENTO": "fecha_nac_raw",
    "EDAD_ANOS": "edad_anios",
    "EDAD_MESES": "edad_meses",
    "RANGO_DE_EDAD": "rango_edad",
    "SEXO": "sexo",
    "GENERO": "genero",
    "NACIONALIDAD": "nacionalidad",
    "ETNIA": "etnia",
    "SITUACION_DE_MOVILIDAD": "situacion_movilidad",
    "FORMA_DE_INGRESO_AL_ECUADOR": "forma_ingreso",
    "ATENCION_EMERGENTE": "atencion_emergente",
    "KIT_DE_ASEO": "kit_aseo",
    "KIT_DE_SALUD": "kit_salud",
    "KIT_ESCOLAR": "kit_escolar",
    "TIENE_ENFERMEDAD_CATASTROFICA": "enfermedad_catastrofica",
    "TIENE_DISCAPACIDAD": "tiene_discapacidad",
    "EMBARAZO_SOLO_PARA_MUJERES": "embarazo",
    "ACTUALMENTE_ESTA_ESTUDIANDO": "estudiando",
    "ATENCION_DE_TRABAJO_SOCIAL": "atencion_trabajo_social",
    "ATENCION_PSICOLOGICA": "atencion_psicologica",
    "ATENCION_LEGAL": "atencion_legal",
    "SALUD": "serv_salud",
    "EDUCACION": "serv_educacion",
    "JUNTA_CANTONAL": "serv_junta_cantonal",
    "REUNIFICACION_FAMILIAR": "serv_reunificacion_familiar",
    "ETI": "serv_eti",
    "ACOGIMINETO_INSTITUCIONAL": "serv_acogimiento_institucional",
    "ACOGIMIENTO_INSTITUCIONAL": "serv_acogimiento_institucional",
    "APOYO_Y_CUSTODIA_FAMILIAR": "serv_apoyo_custodia_familiar",
    "DISCAPACIDADES": "serv_discapacidades",
    "ADULTO_MAYOR": "serv_adulto_mayor",
    "CDI": "serv_cdi",
    "CNH": "serv_cnh",
    "PARTICIPACION_A_TALLERES_DE_CAPACITACION": "part_talleres_capacitacion",
    "PARTICIPACION_A_TALLERES_DE_SENSIBILIZACION": "part_talleres_sensibilizacion",
    "PARTICIPACION_EN_ENCUENTROS_COMUNITARIOS": "part_encuentros_comunitarios",
    "PARTICIPACION_EN_TALLERES_Y_ESPACIOS_LUDICO_DE_RECREACION_DE_NNA": "part_talleres_nna",
    "INSERCION_A_REDES_COMUNITARIAS": "part_redes_comunitarias",
    "SITUACION_MIGRATORIA": "situacion_migratoria",
    "MES": "mes_raw",
    "ANO": "anio_raw",
}


def _norm_text(value) -> str:
    text = "" if pd.isna(value) else str(value)
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"\s+", " ", text)
    return text.strip().upper()


def _norm_col(value) -> str:
    text = _norm_text(value)
    text = re.sub(r"[^A-Z0-9]+", "_", text)
    return text.strip("_")


def _source_signature() -> dict:
    stat = DATA_PATH.stat()
    return {"mtime": stat.st_mtime, "size": stat.st_size}


def _load_prepared_cache() -> pd.DataFrame | None:
    if not DATA_CACHE_PATH.exists():
        return None

    try:
        payload = pd.read_pickle(DATA_CACHE_PATH)
    except Exception:
        return None

    if not isinstance(payload, dict):
        return None
    if payload.get("version") != DATA_CACHE_VERSION:
        return None
    if payload.get("source") != _source_signature():
        return None

    df = payload.get("data")
    return df if isinstance(df, pd.DataFrame) else None


def _write_prepared_cache(df: pd.DataFrame) -> None:
    payload = {
        "version": DATA_CACHE_VERSION,
        "source": _source_signature(),
        "data": df,
    }
    try:
        pd.to_pickle(payload, DATA_CACHE_PATH)
    except Exception:
        pass


def _read_source_excel() -> pd.DataFrame:
    return pd.read_excel(
        DATA_PATH,
        dtype=str,
        usecols=lambda col: _norm_col(col) in COLUMN_MAP,
    )


def _derive_periodo(df: pd.DataFrame) -> pd.Series:
    """Deriva el periodo mensual desde las columnas MES/AÑO del reporte.

    Se usan estas columnas (y no el nombre del archivo de origen) porque el
    nombre del archivo refleja cuándo se consolidó el reporte, no el mes real
    del registro (p.ej. hay filas de un sheet 'ENERO' dentro de un archivo
    llamado 'Abril 2023_...xlsx').
    """
    periodo = pd.Series(pd.NaT, index=df.index, dtype="datetime64[ns]")
    if "mes_raw" not in df.columns or "anio_raw" not in df.columns:
        return periodo

    mes_num = df["mes_raw"].map(_norm_text).map(MESES_NUM)
    anio_num = pd.to_numeric(df["anio_raw"], errors="coerce")
    valid = mes_num.notna() & anio_num.notna()
    if valid.any():
        fechas = (
            anio_num[valid].astype(int).astype(str) + "-"
            + mes_num[valid].astype(int).astype(str).str.zfill(2) + "-01"
        )
        periodo.loc[valid] = pd.to_datetime(fechas, format="%Y-%m-%d")
    return periodo


def _derive_persona_id(df: pd.DataFrame) -> pd.Series:
    """Agrupa las filas que corresponden a la misma persona atendida en
    distintos meses, usando nombre + apellido + fecha de nacimiento
    normalizados como llave. Es la única función que toca esas tres columnas:
    el resultado es un código anónimo (0, 1, 2, ...) sin relación reversible
    obvia con el nombre; las columnas originales se descartan justo después,
    en _prepare_excel_data.
    """
    needed = {"nombres_raw", "apellidos_raw", "fecha_nac_raw"}
    if not needed.issubset(df.columns):
        return pd.Series(range(len(df)), index=df.index)

    key = (
        df["nombres_raw"].map(_norm_text) + "|"
        + df["apellidos_raw"].map(_norm_text) + "|"
        + df["fecha_nac_raw"].astype(str).map(_norm_text)
    )
    codes, _ = pd.factorize(key)
    return pd.Series(codes, index=df.index)


def _clean_yes_no_columns(df: pd.DataFrame) -> pd.DataFrame:
    for col in SI_NO_COLS:
        if col in df.columns:
            df[col] = df[col].map(_norm_text).replace({"": np.nan, "N/D": np.nan})
    return df


def _clean_dimension_columns(df: pd.DataFrame) -> pd.DataFrame:
    dims = [
        "zona", "provincia", "distrito", "ciudad", "rango_edad", "sexo", "genero",
        "nacionalidad", "etnia", "situacion_movilidad", "forma_ingreso",
        "situacion_migratoria",
    ]
    for col in dims:
        if col in df.columns:
            df[col] = df[col].map(_norm_text).replace({"": np.nan, "N/D": np.nan})
    if "zona" in df.columns:
        df["zona"] = df["zona"].str.replace(r"\s+", " ", regex=True).str.strip()
    return df


def _prepare_excel_data(df: pd.DataFrame) -> pd.DataFrame:
    df = df.rename(columns={col: COLUMN_MAP.get(_norm_col(col), _norm_col(col).lower()) for col in df.columns})

    df["periodo"] = _derive_periodo(df)
    df["persona_id"] = _derive_persona_id(df)
    df = df.drop(columns=["nombres_raw", "apellidos_raw", "fecha_nac_raw"], errors="ignore")

    if "edad_anios" in df.columns:
        df["edad_anios"] = df["edad_anios"].astype(str).str.extract(r"(\d+)", expand=False).astype(float)

    df = _clean_dimension_columns(df)
    df = _clean_yes_no_columns(df)
    df = df.dropna(subset=["periodo"])
    if df.empty:
        st.error("No se pudo identificar el periodo mensual a partir de las columnas 'MES' y 'AÑO' del Excel.")
        st.stop()
    df["anio"] = df["periodo"].dt.year.astype("Int64")
    df["mes_num"] = df["periodo"].dt.month.astype("Int64")

    expected_cols = [
        "periodo", "anio", "mes_num", "persona_id", "zona", "provincia", "distrito", "ciudad",
        "rango_edad", "sexo", "genero", "nacionalidad", "etnia",
        "situacion_movilidad", "forma_ingreso", "situacion_migratoria", "edad_anios",
        *SI_NO_COLS,
    ]
    for col in expected_cols:
        if col not in df.columns:
            df[col] = np.nan
    return df[expected_cols]


def formato_mes_anio(ts: pd.Timestamp) -> str:
    return f"{MESES_ES[ts.month]} {ts.year}"


# ----------------------------------------------------------------------------
# Portada / pantalla de bienvenida
# ----------------------------------------------------------------------------
if "ingresado" not in st.session_state:
    st.session_state["ingresado"] = False
if "seccion" not in st.session_state:
    st.session_state["seccion"] = PORTADA_CARDS[0]


def mostrar_portada():
    df_preview = load_data()
    rango_fechas = f"{formato_mes_anio(df_preview['periodo'].min())} - {formato_mes_anio(df_preview['periodo'].max())}"

    # Borde de color por tarjeta, distinto para cada una de las 6 (2 filas x 3
    # columnas, mismo orden que PORTADA_CARDS), usando el mismo color que cada
    # seccion ya usa en sus graficas (ver COLOR_SECCIONES).
    #
    # Cada fila de st.columns(3) queda envuelta en su PROPIO stLayoutWrapper
    # (por eso stHorizontalBlock:nth-of-type(N) no sirve para distinguir filas:
    # cada una es "la 1ra de su tipo" dentro de su propio wrapper). Los
    # wrapper si son hermanos directos entre si, así que se usan para
    # identificar la fila; stColumn si son hermanos directos dentro de cada
    # fila, así que sirven para identificar la columna.
    tarjeta_css_reglas = []
    for row_start in range(0, len(PORTADA_CARDS), 3):
        wrapper_pos = 2 + (row_start // 3)  # 1ra fila = wrapper #2, 2da = #3
        for col_num, card in enumerate(PORTADA_CARDS[row_start:row_start + 3], start=1):
            color = COLOR_SECCIONES.get(card, COLOR_PRIMARY)
            tarjeta_css_reglas.append(
                f'[data-testid="stLayoutWrapper"]:nth-of-type({wrapper_pos}) '
                f'[data-testid="stColumn"]:nth-of-type({col_num}) button {{ '
                f'border-left: 4px solid {color}; }}'
            )
    tarjeta_css = "\n        ".join(tarjeta_css_reglas)

    logo1_b64 = _img_to_base64(LOGO1_PATH)
    logo2_b64 = _img_to_base64(LOGO2_PATH)

    logo1_html = (
        f'<img src="data:image/jpeg;base64,{logo1_b64}" class="portada-logo" />'
        if logo1_b64 else ""
    )
    logo2_html = (
        f'<img src="data:image/jpeg;base64,{logo2_b64}" class="portada-logo" />'
        if logo2_b64 else ""
    )

    st.markdown(
        f"""
        <style>
        [data-testid="stSidebar"] {{display: none;}}
        html, body, [data-testid="stApp"], [data-testid="stHeader"], [data-testid="stMain"] {{
            background: #f9f9f7;
        }}
        [data-testid="stToolbar"] {{
            display: none;
        }}
        [data-testid="stCaptionContainer"] {{
            color: #898781;
        }}
        .portada-wrap {{
            display: flex;
            flex-direction: column;
            align-items: center;
            text-align: center;
            padding-top: 3vh;
        }}
        .portada-logos {{
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 40px;
            margin-bottom: 1.8rem;
        }}
        .portada-logo {{
            height: 100px;
            object-fit: contain;
        }}
        .portada-titulo-row {{
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 22px;
            margin-bottom: 1.6rem;
        }}
        .portada-titulo {{
            font-size: 4rem;
            font-weight: 800;
            color: #14208a;
            letter-spacing: 1px;
            margin: 0;
        }}
        .portada-barra {{
            width: 6px;
            height: 60px;
            background: #2a78d6;
            border-radius: 3px;
        }}
        .portada-info-box {{
            background: #ffffff;
            border: 1px solid #e1e0d9;
            border-radius: 8px;
            padding: 10px 28px;
            font-size: 1.15rem;
            font-weight: 700;
            color: #14208a;
            margin-bottom: 10px;
            box-shadow: 0 1px 3px rgba(11, 11, 11, 0.06);
        }}
        .portada-fecha-box {{
            background: #ffffff;
            border: 1px solid #e1e0d9;
            border-radius: 8px;
            padding: 8px 28px;
            font-size: 1.05rem;
            font-weight: 700;
            color: #e0a800;
            margin-bottom: 2.2rem;
            box-shadow: 0 1px 3px rgba(11, 11, 11, 0.06);
        }}
        div[data-testid="stButton"] > button {{
            border: 1px solid #e1e0d9;
            border-radius: 12px;
            padding: 22px 14px;
            font-size: 1.05rem;
            font-weight: 700;
            color: #4a4a4a;
            min-height: 80px;
            background: #ffffff;
            box-shadow: 0 1px 3px rgba(11, 11, 11, 0.06);
            transition: box-shadow 0.15s ease, border-color 0.15s ease,
                        color 0.15s ease, transform 0.15s ease;
        }}
        div[data-testid="stButton"] > button:hover {{
            border-color: #14208a;
            color: #14208a;
            background: #f7f8ff;
            box-shadow: 0 6px 16px rgba(20, 32, 138, 0.14);
            transform: translateY(-2px);
        }}
        div[data-testid="stButton"] > button[kind="primary"] {{
            display: none;
        }}
        /* Acento de color por tarjeta (una regla por posicion fila/columna),
           a juego con el color que cada seccion usa en sus propias graficas. */
        {tarjeta_css}
        @media (max-width: 900px) {{
            .portada-titulo {{ font-size: 2.6rem; }}
        }}
        </style>

        <div class="portada-wrap">
            <div class="portada-logos">
                {logo1_html}
                {logo2_html}
            </div>
            <div class="portada-titulo-row">
                <div class="portada-titulo">OBSERVATORIO</div>
                <div class="portada-barra"></div>
            </div>
            <div class="portada-info-box">Datos disponibles</div>
            <div class="portada-fecha-box">{rango_fechas}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    for row_start in range(0, len(PORTADA_CARDS), 3):
        cols = st.columns(3)
        for col, card in zip(cols, PORTADA_CARDS[row_start:row_start + 3]):
            with col:
                icon = PORTADA_ICONS.get(card, "")
                if st.button(f"{icon}  {card}", use_container_width=True):
                    st.session_state["seccion"] = card
                    st.session_state["ingresado"] = True
                    st.rerun()

    col_a, col_b, col_c = st.columns([1, 1, 1])
    with col_b:
        if st.button("Ingresar ➜", use_container_width=True, type="primary"):
            st.session_state["ingresado"] = True
            st.rerun()

    st.caption(
        "Fuente: Reporte Consolidado de Movilidad Humana. "
        "Este panel no expone nombres, direcciones, teléfonos, correos ni documentos de identificación."
    )


if not st.session_state["ingresado"]:
    mostrar_portada()
    st.stop()

# ----------------------------------------------------------------------------
# Carga de datos
# ----------------------------------------------------------------------------
df = load_data()

top_col1, top_col2 = st.columns([6, 1], vertical_alignment="center")
with top_col1:
    st.title("🌎 Dashboard de Movilidad Humana")
    st.caption(
        "Reporte consolidado de caracterización del Servicio de Movilidad Humana. "
        "Los datos mostrados son agregados y no incluyen información de identificación "
        "personal (nombres, direcciones, teléfonos, correos ni documentos)."
    )
with top_col2:
    if st.button("⟵ Salir"):
        st.session_state["ingresado"] = False
        st.rerun()

# ----------------------------------------------------------------------------
# Datos
# ----------------------------------------------------------------------------
f = df.copy()
seccion_actual = st.session_state.get("seccion", PORTADA_CARDS[0])
st.markdown(
    """
    <style>
    [data-testid="stSidebar"] {display: none;}
    [data-testid="collapsedControl"] {display: none;}

    /* Streamlit sigue el tema oscuro/claro del sistema operativo; con el SO en
       oscuro, html/body/stApp/stHeader se quedan negros aunque stMain sea
       claro. Se fija el mismo tono claro en todos para que no queden franjas
       oscuras arriba ni en los bordes de la pagina. */
    html, body, [data-testid="stApp"], [data-testid="stHeader"] {
        background: #f9f9f7;
    }
    /* Barra de herramientas de Streamlit (Deploy, menu, "Made with") oculta:
       esta pantalla es para presentar el reporte, no para desarrollarlo. */
    [data-testid="stToolbar"] {
        display: none;
    }

    /* Fondo de pagina vs. superficie de tarjetas (mismo par de tonos que usan
       las graficas de Plotly, para que todo se lea como un solo sistema). */
    [data-testid="stMain"] {
        background: #f9f9f7;
    }
    [data-testid="stMainBlockContainer"] {
        padding-top: 2rem;
        max-width: 1300px;
    }

    /* Titulo principal y titulo de seccion */
    [data-testid="stHeading"] h1 {
        color: #14208a;
        font-weight: 800;
    }
    [data-testid="stHeading"] h3 {
        color: #14208a;
        font-weight: 800;
        border-left: 5px solid #14208a;
        padding-left: 14px;
        margin: 6px 0 14px 0;
    }
    [data-testid="stCaptionContainer"] {
        color: #898781;
    }

    /* Tarjetas KPI */
    div[data-testid="stMetric"] {
        background: #fcfcfb;
        border: 1px solid #e1e0d9;
        border-top: 3px solid #2a78d6;
        border-radius: 10px;
        padding: 14px 16px 10px;
        box-shadow: 0 1px 3px rgba(11, 11, 11, 0.06);
        transition: box-shadow 0.15s ease;
    }
    div[data-testid="stMetric"]:hover {
        box-shadow: 0 4px 12px rgba(11, 11, 11, 0.10);
    }
    label[data-testid="stMetricLabel"] p {
        color: #52514e;
        font-size: 0.8rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.03em;
    }
    div[data-testid="stMetricValue"] {
        color: #14208a;
        font-weight: 800;
    }

    /* Boton "Salir" */
    button[data-testid="stBaseButton-secondary"] {
        background: #ffffff;
        border: 1px solid #d9d9d9;
        border-radius: 8px;
        color: #14208a;
        font-weight: 700;
        box-shadow: 0 1px 3px rgba(11, 11, 11, 0.06);
        transition: background 0.15s ease, color 0.15s ease, box-shadow 0.15s ease;
    }
    button[data-testid="stBaseButton-secondary"]:hover {
        background: #14208a;
        color: #ffffff;
        border-color: #14208a;
    }

    /* Tarjetas que envuelven cada grafica */
    div[data-testid="stPlotlyChart"] {
        background: #fcfcfb;
        border: 1px solid #e1e0d9;
        border-radius: 12px;
        padding: 6px;
        box-shadow: 0 1px 3px rgba(11, 11, 11, 0.06);
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ----------------------------------------------------------------------------
# KPIs
# ----------------------------------------------------------------------------
col1, col2, col3, col4, col5, col6 = st.columns(6)
col1.metric("Total de atenciones", f"{len(f):,}".replace(",", "."))
col2.metric("Personas únicas", f"{f['persona_id'].nunique():,}".replace(",", "."))
col3.metric("% Mujeres", f"{100 * (f['sexo'] == 'MUJER').sum() / f['sexo'].notna().sum():.1f}%" if f['sexo'].notna().sum() else "N/D")
irregular_pct = 100 * (f["situacion_migratoria"] == "IRREGULAR").sum() / f["situacion_migratoria"].notna().sum() if f["situacion_migratoria"].notna().sum() else 0
col4.metric("% Situación irregular", f"{irregular_pct:.1f}%")
nna_pct = 100 * f["rango_edad"].isin(["NN", "ADOLESCENTE"]).sum() / f["rango_edad"].notna().sum() if f["rango_edad"].notna().sum() else 0
col5.metric("% NNA (niños/adolescentes)", f"{nna_pct:.1f}%")
col6.metric("Nacionalidades distintas", f["nacionalidad"].nunique())
st.caption(
    "'Personas únicas' agrupa registros de nombre + apellido + fecha de nacimiento "
    "(normalizados); esos tres datos se usan solo para el conteo y no se guardan ni se muestran."
)

st.divider()

st.subheader(seccion_actual)


def plot_si_bars(cols: list[str], title: str, height: int = 420, color: str = COLOR_PRIMARY) -> None:
    data = []
    for col in cols:
        if col in f.columns and f[col].notna().sum() > 0:
            data.append({
                "indicador": FRIENDLY_NAMES.get(col, col),
                "pct_si": si_pct(f[col]),
                "n_valido": f[col].notna().sum(),
            })
    if not data:
        st.info("No hay datos disponibles para esta seccion.")
        return
    chart_df = pd.DataFrame(data).sort_values("pct_si", ascending=True)
    fig = px.bar(
        chart_df, x="pct_si", y="indicador", orientation="h", title=title,
        color_discrete_sequence=[color], text="pct_si",
    )
    fig.update_traces(texttemplate="%{text:.0f}%", textposition="outside", cliponaxis=False)
    fig.update_layout(height=height, xaxis_title="% SI", yaxis_title="")
    fig.update_xaxes(range=[0, _headroom(chart_df["pct_si"].max())])
    _style_chart(fig)
    st.plotly_chart(fig, use_container_width=True)


if seccion_actual == PORTADA_CARDS[0]:
    ts = f.groupby("periodo").size().reset_index(name="registros")
    fig_ts = px.line(ts, x="periodo", y="registros", markers=True, color_discrete_sequence=[COLOR_PRIMARY],
                      title="Evolución mensual")
    fig_ts.update_traces(line=dict(width=2), marker=dict(size=7))
    fig_ts.update_layout(height=380, xaxis_title="Mes", yaxis_title="Numero de registros")
    _style_chart(fig_ts)
    st.plotly_chart(fig_ts, use_container_width=True)

    c1, c2 = st.columns(2)
    with c1:
        nac = f["nacionalidad"].value_counts().head(10).reset_index()
        nac.columns = ["nacionalidad", "registros"]
        fig_nac = px.bar(nac.sort_values("registros"), x="registros", y="nacionalidad", orientation="h",
                          title="Top 10 nacionalidades", color_discrete_sequence=[COLOR_PRIMARY], text="registros")
        fig_nac.update_traces(texttemplate="%{text:,}", textposition="outside", cliponaxis=False)
        fig_nac.update_layout(height=380)
        fig_nac.update_xaxes(range=[0, _headroom(nac["registros"].max())])
        _style_chart(fig_nac)
        st.plotly_chart(fig_nac, use_container_width=True)
    with c2:
        edad = f["rango_edad"].value_counts().reset_index()
        edad.columns = ["rango_edad", "registros"]
        fig_edad = px.pie(edad, names="rango_edad", values="registros", title="Distribucion por rango de edad",
                           hole=0.4, color="rango_edad", color_discrete_map=COLOR_RANGO_EDAD)
        fig_edad.update_layout(height=380)
        _style_chart(fig_edad, legend=True)
        st.plotly_chart(fig_edad, use_container_width=True)

    c3, c4 = st.columns(2)
    with c3:
        genero = f["genero"].value_counts().reset_index()
        genero.columns = ["genero", "registros"]
        fig_genero = px.bar(genero, x="genero", y="registros", title="Distribucion por genero",
                             color="genero", color_discrete_map=COLOR_GENERO, text="registros")
        fig_genero.update_traces(texttemplate="%{text:,}", textposition="outside", cliponaxis=False)
        fig_genero.update_layout(height=350)
        fig_genero.update_yaxes(range=[0, _headroom(genero["registros"].max())])
        _style_chart(fig_genero)
        st.plotly_chart(fig_genero, use_container_width=True)
    with c4:
        ingreso = f["forma_ingreso"].value_counts().reset_index()
        ingreso.columns = ["forma_ingreso", "registros"]
        fig_ing = px.bar(ingreso, x="forma_ingreso", y="registros", title="Forma de ingreso al Ecuador",
                          color="forma_ingreso", color_discrete_map=COLOR_FORMA_INGRESO, text="registros")
        fig_ing.update_traces(texttemplate="%{text:,}", textposition="outside", cliponaxis=False)
        fig_ing.update_layout(height=350)
        fig_ing.update_yaxes(range=[0, _headroom(ingreso["registros"].max())])
        _style_chart(fig_ing)
        st.plotly_chart(fig_ing, use_container_width=True)

    if f["edad_anios"].notna().sum() > 20:
        fig_hist = px.histogram(f.dropna(subset=["edad_anios"]), x="edad_anios", nbins=30,
                                title="Distribucion de edad en anios", color_discrete_sequence=[COLOR_PRIMARY])
        fig_hist.update_layout(height=320, xaxis_title="Edad", yaxis_title="Numero de registros")
        _style_chart(fig_hist)
        st.plotly_chart(fig_hist, use_container_width=True)
    st.stop()

if seccion_actual == PORTADA_CARDS[1]:
    plot_si_bars(
        ["tiene_discapacidad", "enfermedad_catastrofica", "embarazo"],
        "Indicadores de vulnerabilidad (% SI)",
        color=COLOR_SECCIONES[PORTADA_CARDS[1]],
    )
    st.stop()

if seccion_actual == PORTADA_CARDS[2]:
    plot_si_bars(
        [
            "atencion_trabajo_social", "atencion_psicologica", "atencion_legal",
            "serv_salud", "serv_educacion", "serv_junta_cantonal",
            "serv_reunificacion_familiar", "serv_eti", "serv_acogimiento_institucional",
            "serv_apoyo_custodia_familiar", "serv_discapacidades", "serv_adulto_mayor",
            "serv_cdi", "serv_cnh",
        ],
        "Intervenciones tecnicas y servicios recibidos (% SI)",
        height=560,
        color=COLOR_SECCIONES[PORTADA_CARDS[2]],
    )
    st.stop()

if seccion_actual == PORTADA_CARDS[3]:
    c1, c2 = st.columns(2)
    with c1:
        mig = f["situacion_migratoria"].value_counts().reset_index()
        mig.columns = ["situacion_migratoria", "registros"]
        fig_mig = px.pie(mig, names="situacion_migratoria", values="registros", title="Situacion migratoria",
                          hole=0.4, color="situacion_migratoria", color_discrete_map=COLOR_SITUACION_MIGRATORIA)
        fig_mig.update_layout(height=380)
        _style_chart(fig_mig, legend=True)
        st.plotly_chart(fig_mig, use_container_width=True)
    with c2:
        mov = f["situacion_movilidad"].value_counts().head(10).reset_index()
        mov.columns = ["situacion_movilidad", "registros"]
        fig_mov = px.bar(mov.sort_values("registros"), x="registros", y="situacion_movilidad", orientation="h",
                         title="Situacion de movilidad", color="situacion_movilidad",
                         color_discrete_map=COLOR_SITUACION_MOVILIDAD, text="registros")
        fig_mov.update_traces(texttemplate="%{text:,}", textposition="outside", cliponaxis=False)
        fig_mov.update_layout(height=380)
        fig_mov.update_xaxes(range=[0, _headroom(mov["registros"].max())])
        _style_chart(fig_mov)
        st.plotly_chart(fig_mov, use_container_width=True)
    st.stop()

if seccion_actual == PORTADA_CARDS[4]:
    plot_si_bars(
        ["atencion_emergente", "kit_aseo", "kit_salud", "kit_escolar"],
        "Asistencia humanitaria entregada (% SI)",
        color=COLOR_SECCIONES[PORTADA_CARDS[4]],
    )
    st.stop()

if seccion_actual == PORTADA_CARDS[5]:
    plot_si_bars(
        [
            "part_talleres_capacitacion", "part_talleres_sensibilizacion",
            "part_encuentros_comunitarios", "part_talleres_nna",
            "part_redes_comunitarias",
        ],
        "Participacion e integracion comunitaria (% SI)",
        color=COLOR_SECCIONES[PORTADA_CARDS[5]],
    )
    st.stop()
