
import hashlib
import hmac
import secrets
import sqlite3
from datetime import date, datetime
from io import BytesIO
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ============================================================
# CONFIGURACIÓN
# ============================================================

st.set_page_config(
    page_title="Jet Airways Academy | Sistema de Ventas",
    page_icon="✈️",
    layout="wide",
)


# ============================================================
# IDENTIDAD VISUAL JET AIRWAYS ACADEMY
# ============================================================

def aplicar_estilo_corporativo():
    st.markdown(
        """
        <style>
        :root {
            --jet-navy: #061B42;
            --jet-navy-2: #0A2759;
            --jet-blue: #123D7A;
            --jet-red: #E30613;
            --jet-red-hover: #B9040E;
            --jet-white: #FFFFFF;
            --jet-soft: #DCE7F7;
        }

        .stApp {
            background:
                radial-gradient(circle at 85% 15%, rgba(25, 72, 139, 0.34), transparent 30%),
                linear-gradient(135deg, var(--jet-navy) 0%, #071D46 48%, #020D22 100%);
            color: var(--jet-white);
        }

        [data-testid="stHeader"] {
            background: rgba(6, 27, 66, 0.82);
        }

        [data-testid="stSidebar"] {
            background: linear-gradient(180deg, #041432 0%, #082653 100%);
            border-right: 1px solid rgba(255,255,255,0.12);
        }

        [data-testid="stSidebar"] * {
            color: var(--jet-white);
        }

        h1, h2, h3, h4, h5, h6, p, label,
        [data-testid="stCaptionContainer"] {
            color: var(--jet-white) !important;
        }

        div[data-testid="stForm"] {
            background: rgba(9, 36, 80, 0.90);
            border: 1px solid rgba(255, 255, 255, 0.22);
            border-radius: 18px;
            padding: 1.5rem 1.6rem 1.1rem 1.6rem;
            box-shadow: 0 20px 60px rgba(0, 0, 0, 0.28);
        }

        div[data-testid="stTextInput"] input,
        div[data-testid="stNumberInput"] input,
        div[data-baseweb="select"] > div,
        div[data-testid="stDateInput"] input {
            background: rgba(255, 255, 255, 0.96) !important;
            color: #071A3E !important;
            border-radius: 10px !important;
        }

        div[data-testid="stTextInput"] input::placeholder {
            color: #6B7280 !important;
        }

        .stButton > button,
        .stDownloadButton > button,
        div[data-testid="stFormSubmitButton"] button {
            background: linear-gradient(90deg, #E30613 0%, #F32632 100%) !important;
            color: white !important;
            border: none !important;
            border-radius: 10px !important;
            font-weight: 700 !important;
            min-height: 3rem;
            box-shadow: 0 8px 20px rgba(227, 6, 19, 0.25);
        }

        .stButton > button:hover,
        .stDownloadButton > button:hover,
        div[data-testid="stFormSubmitButton"] button:hover {
            background: var(--jet-red-hover) !important;
            transform: translateY(-1px);
        }

        div[data-testid="stMetric"] {
            background: rgba(255, 255, 255, 0.09);
            border: 1px solid rgba(255, 255, 255, 0.16);
            border-radius: 14px;
            padding: 1rem;
        }

        div[data-testid="stMetric"] label,
        div[data-testid="stMetricValue"] {
            color: white !important;
        }

        div[data-testid="stDataFrame"] {
            background: white;
            border-radius: 12px;
            overflow: hidden;
        }

        [data-testid="stExpander"] {
            background: rgba(9, 36, 80, 0.78);
            border: 1px solid rgba(255, 255, 255, 0.16);
            border-radius: 12px;
        }

        .jet-brand-title {
            text-align: center;
            color: white;
            font-size: 2rem;
            font-weight: 800;
            margin-top: 0.15rem;
            margin-bottom: 0.15rem;
            letter-spacing: 0.02em;
        }

        .jet-brand-subtitle {
            text-align: center;
            color: #DCE7F7;
            font-size: 1rem;
            margin-bottom: 1.4rem;
        }

        .jet-footer {
            text-align: center;
            color: rgba(255,255,255,0.58);
            font-size: 0.78rem;
            margin-top: 2rem;
        }

        .block-container {
            padding-top: 2.2rem;
            padding-bottom: 2rem;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


aplicar_estilo_corporativo()

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "ventas.db"
LOGO_PATH = BASE_DIR / "assets" / "logo_jet_airways.png"

TIPO_CAMBIO_USD_BS = 6.96

COMISIONES = {
    "PPL": {"monto": 100.0, "moneda": "USD"},
    "PPC": {"monto": 275.0, "moneda": "USD"},
    "IFR": {"monto": 275.0, "moneda": "USD"},
    "TCP": {"monto": 50.0, "moneda": "USD"},
    "EOU": {"monto": 30.0, "moneda": "USD"},
    "AGT": {"monto": 120.0, "moneda": "BOB"},
    "AGR": {"monto": 140.0, "moneda": "BOB"},
    "MTC": {"monto": 60.0, "moneda": "USD"},
}

VENDEDORES_PREDETERMINADOS = [
    "Ana",
    "Carlos",
    "Daniela",
    "José",
    "María",
]

ROLES = ["Administrador", "Supervisor", "Vendedor"]


# ============================================================
# BASE DE DATOS
# ============================================================

def conectar():
    return sqlite3.connect(DB_PATH, check_same_thread=False)


def crear_base_datos():
    """Crea o actualiza las tablas sin borrar las ventas existentes."""
    with conectar() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS ventas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fecha TEXT NOT NULL,
                curso TEXT NOT NULL,
                monto REAL NOT NULL,
                cliente TEXT NOT NULL,
                vendedor TEXT NOT NULL,
                comision_original REAL NOT NULL,
                moneda_comision TEXT NOT NULL,
                tipo_cambio REAL NOT NULL,
                comision_bs REAL NOT NULL,
                ingreso_neto_bs REAL NOT NULL,
                creado_en TEXT NOT NULL
            )
            """
        )

        columnas_ventas = {
            fila[1] for fila in conn.execute("PRAGMA table_info(ventas)").fetchall()
        }
        if "creado_por" not in columnas_ventas:
            conn.execute(
                "ALTER TABLE ventas ADD COLUMN creado_por TEXT NOT NULL DEFAULT ''"
            )

        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS usuarios (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT NOT NULL,
                usuario TEXT NOT NULL UNIQUE,
                password_hash TEXT NOT NULL,
                salt TEXT NOT NULL,
                rol TEXT NOT NULL,
                activo INTEGER NOT NULL DEFAULT 1,
                creado_en TEXT NOT NULL
            )
            """
        )
        conn.commit()


def crear_hash_password(password, salt_hex=None):
    if salt_hex is None:
        salt = secrets.token_bytes(32)
        salt_hex = salt.hex()
    else:
        salt = bytes.fromhex(salt_hex)

    resultado = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        salt,
        200_000,
    )
    return resultado.hex(), salt_hex


def verificar_password(password, hash_guardado, salt_hex):
    hash_ingresado, _ = crear_hash_password(password, salt_hex)
    return hmac.compare_digest(hash_ingresado, hash_guardado)


def crear_usuario(nombre, usuario, password, rol):
    nombre = nombre.strip()
    usuario = usuario.strip().lower()

    if not nombre or not usuario or not password:
        raise ValueError("Nombre, usuario y contraseña son obligatorios.")
    if rol not in ROLES:
        raise ValueError("El rol seleccionado no es válido.")
    if len(password) < 8:
        raise ValueError("La contraseña debe tener al menos 8 caracteres.")

    password_hash, salt = crear_hash_password(password)

    try:
        with conectar() as conn:
            conn.execute(
                """
                INSERT INTO usuarios (
                    nombre, usuario, password_hash, salt, rol, activo, creado_en
                )
                VALUES (?, ?, ?, ?, ?, 1, ?)
                """,
                (
                    nombre,
                    usuario,
                    password_hash,
                    salt,
                    rol,
                    datetime.now().isoformat(timespec="seconds"),
                ),
            )
            conn.commit()
    except sqlite3.IntegrityError as error:
        raise ValueError("Ese nombre de usuario ya existe.") from error


def asegurar_administrador_inicial():
    with conectar() as conn:
        cantidad = conn.execute("SELECT COUNT(*) FROM usuarios").fetchone()[0]

    if cantidad == 0:
        crear_usuario(
            "Administrador principal",
            "admin",
            "Admin123!",
            "Administrador",
        )


def autenticar_usuario(usuario, password):
    usuario = usuario.strip().lower()
    with conectar() as conn:
        fila = conn.execute(
            """
            SELECT id, nombre, usuario, password_hash, salt, rol, activo
            FROM usuarios
            WHERE usuario = ?
            """,
            (usuario,),
        ).fetchone()

    if fila is None:
        return None

    id_usuario, nombre, usuario_db, hash_guardado, salt, rol, activo = fila
    if not activo or not verificar_password(password, hash_guardado, salt):
        return None

    return {
        "id": id_usuario,
        "nombre": nombre,
        "usuario": usuario_db,
        "rol": rol,
    }


def cargar_usuarios():
    with conectar() as conn:
        return pd.read_sql_query(
            """
            SELECT id, nombre, usuario, rol, activo, creado_en
            FROM usuarios
            ORDER BY nombre
            """,
            conn,
        )


def cambiar_estado_usuario(id_usuario, activo):
    with conectar() as conn:
        conn.execute(
            "UPDATE usuarios SET activo = ? WHERE id = ?",
            (1 if activo else 0, int(id_usuario)),
        )
        conn.commit()


def cambiar_password_usuario(id_usuario, nueva_password):
    if len(nueva_password) < 8:
        raise ValueError("La contraseña debe tener al menos 8 caracteres.")

    password_hash, salt = crear_hash_password(nueva_password)
    with conectar() as conn:
        conn.execute(
            "UPDATE usuarios SET password_hash = ?, salt = ? WHERE id = ?",
            (password_hash, salt, int(id_usuario)),
        )
        conn.commit()


def guardar_venta(fecha, curso, monto, cliente, vendedor, tipo_cambio, creado_por):
    datos = COMISIONES[curso]
    comision_bs = (
        datos["monto"] * tipo_cambio
        if datos["moneda"] == "USD"
        else datos["monto"]
    )
    ingreso_neto = monto - comision_bs

    with conectar() as conn:
        conn.execute(
            """
            INSERT INTO ventas (
                fecha, curso, monto, cliente, vendedor,
                comision_original, moneda_comision, tipo_cambio,
                comision_bs, ingreso_neto_bs, creado_por, creado_en
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                fecha.isoformat(),
                curso,
                float(monto),
                cliente.strip(),
                vendedor.strip(),
                datos["monto"],
                datos["moneda"],
                float(tipo_cambio),
                float(comision_bs),
                float(ingreso_neto),
                creado_por,
                datetime.now().isoformat(timespec="seconds"),
            ),
        )
        conn.commit()

    return comision_bs, ingreso_neto


def cargar_ventas():
    with conectar() as conn:
        df = pd.read_sql_query(
            "SELECT * FROM ventas ORDER BY fecha DESC, id DESC",
            conn,
        )

    if not df.empty:
        df["fecha"] = pd.to_datetime(df["fecha"])
        df["creado_en"] = pd.to_datetime(df["creado_en"])
    return df


def actualizar_venta(id_venta, fecha, curso, monto, cliente, vendedor, tipo_cambio):
    datos = COMISIONES[curso]
    comision_bs = (
        datos["monto"] * tipo_cambio
        if datos["moneda"] == "USD"
        else datos["monto"]
    )
    ingreso_neto = monto - comision_bs

    with conectar() as conn:
        conn.execute(
            """
            UPDATE ventas
            SET fecha = ?, curso = ?, monto = ?, cliente = ?, vendedor = ?,
                comision_original = ?, moneda_comision = ?, tipo_cambio = ?,
                comision_bs = ?, ingreso_neto_bs = ?
            WHERE id = ?
            """,
            (
                fecha.isoformat(),
                curso,
                float(monto),
                cliente.strip(),
                vendedor.strip(),
                datos["monto"],
                datos["moneda"],
                float(tipo_cambio),
                float(comision_bs),
                float(ingreso_neto),
                int(id_venta),
            ),
        )
        conn.commit()


def eliminar_venta(id_venta):
    with conectar() as conn:
        conn.execute("DELETE FROM ventas WHERE id = ?", (int(id_venta),))
        conn.commit()


def importar_dataframe(df, tipo_cambio):
    df = df.copy()
    df.columns = df.columns.astype(str).str.strip()
    df = df.loc[:, ~df.columns.str.contains("^Unnamed", case=False, regex=True)]

    if "Tipo_ingreso" in df.columns and "Curso" not in df.columns:
        df = df.rename(columns={"Tipo_ingreso": "Curso"})

    obligatorias = ["Fecha", "Curso", "Monto", "Cliente", "Vendedor"]
    faltantes = [c for c in obligatorias if c not in df.columns]
    if faltantes:
        raise ValueError(f"Faltan columnas obligatorias: {faltantes}")

    df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce", dayfirst=True)
    df["Curso"] = df["Curso"].astype(str).str.strip().str.upper()
    df["Monto"] = (
        df["Monto"]
        .astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("Bs", "", regex=False)
        .str.replace("$", "", regex=False)
        .str.strip()
    )
    df["Monto"] = pd.to_numeric(df["Monto"], errors="coerce")
    df["Cliente"] = df["Cliente"].astype(str).str.strip()
    df["Vendedor"] = df["Vendedor"].astype(str).str.strip()

    errores = df[
        df["Fecha"].isna()
        | df["Monto"].isna()
        | (df["Monto"] <= 0)
        | ~df["Curso"].isin(COMISIONES.keys())
        | (df["Cliente"] == "")
        | (df["Vendedor"] == "")
    ].copy()

    validos = df.drop(index=errores.index).copy()

    registros = []
    for _, fila in validos.iterrows():
        curso = fila["Curso"]
        datos = COMISIONES[curso]
        comision_bs = (
            datos["monto"] * tipo_cambio
            if datos["moneda"] == "USD"
            else datos["monto"]
        )
        ingreso_neto = fila["Monto"] - comision_bs

        registros.append(
            (
                fila["Fecha"].date().isoformat(),
                curso,
                float(fila["Monto"]),
                fila["Cliente"],
                fila["Vendedor"],
                datos["monto"],
                datos["moneda"],
                float(tipo_cambio),
                float(comision_bs),
                float(ingreso_neto),
                datetime.now().isoformat(timespec="seconds"),
            )
        )

    if registros:
        with conectar() as conn:
            conn.executemany(
                """
                INSERT INTO ventas (
                    fecha, curso, monto, cliente, vendedor,
                    comision_original, moneda_comision, tipo_cambio,
                    comision_bs, ingreso_neto_bs, creado_en
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                registros,
            )
            conn.commit()

    return len(validos), errores


# ============================================================
# REPORTE EXCEL
# ============================================================

def generar_reporte_excel(df):
    if df.empty:
        raise ValueError("No existen ventas para exportar.")

    trabajo = df.copy()
    trabajo["Fecha"] = trabajo["fecha"].dt.date
    trabajo["Mes"] = trabajo["fecha"].dt.to_period("M").astype(str)

    ventas = trabajo[
        [
            "id",
            "Fecha",
            "curso",
            "monto",
            "cliente",
            "vendedor",
            "comision_original",
            "moneda_comision",
            "tipo_cambio",
            "comision_bs",
            "ingreso_neto_bs",
        ]
    ].rename(
        columns={
            "id": "ID",
            "curso": "Curso",
            "monto": "Ingreso_bruto_Bs",
            "cliente": "Cliente",
            "vendedor": "Vendedor",
            "comision_original": "Comision_original",
            "moneda_comision": "Moneda_comision",
            "tipo_cambio": "Tipo_cambio",
            "comision_bs": "Comision_Bs",
            "ingreso_neto_bs": "Ingreso_neto_Bs",
        }
    )

    resumen_curso = (
        trabajo.groupby("curso", as_index=False)
        .agg(
            Ingreso_bruto_Bs=("monto", "sum"),
            Comisiones_Bs=("comision_bs", "sum"),
            Ingreso_neto_Bs=("ingreso_neto_bs", "sum"),
            Cantidad_ventas=("id", "count"),
            Venta_promedio_Bs=("monto", "mean"),
        )
        .rename(columns={"curso": "Curso"})
        .sort_values("Ingreso_neto_Bs", ascending=False)
    )

    resumen_vendedor = (
        trabajo.groupby("vendedor", as_index=False)
        .agg(
            Ingreso_bruto_Bs=("monto", "sum"),
            Comisiones_Bs=("comision_bs", "sum"),
            Ingreso_neto_Bs=("ingreso_neto_bs", "sum"),
            Cantidad_ventas=("id", "count"),
        )
        .rename(columns={"vendedor": "Vendedor"})
        .sort_values("Ingreso_neto_Bs", ascending=False)
    )

    resumen_diario = (
        trabajo.groupby(["Fecha", "curso"], as_index=False)
        .agg(
            Ingreso_bruto_Bs=("monto", "sum"),
            Comisiones_Bs=("comision_bs", "sum"),
            Ingreso_neto_Bs=("ingreso_neto_bs", "sum"),
            Cantidad_ventas=("id", "count"),
        )
        .rename(columns={"curso": "Curso"})
        .sort_values(["Fecha", "Curso"])
    )

    resumen_mensual = (
        trabajo.groupby(["Mes", "curso"], as_index=False)
        .agg(
            Ingreso_bruto_Bs=("monto", "sum"),
            Comisiones_Bs=("comision_bs", "sum"),
            Ingreso_neto_Bs=("ingreso_neto_bs", "sum"),
            Cantidad_ventas=("id", "count"),
        )
        .rename(columns={"curso": "Curso"})
        .sort_values(["Mes", "Curso"])
    )

    curso_top = resumen_curso.iloc[0]["Curso"]
    vendedor_top = resumen_vendedor.iloc[0]["Vendedor"]

    dashboard = pd.DataFrame(
        {
            "Indicador": [
                "Ingreso bruto total",
                "Comisiones totales",
                "Ingreso neto total",
                "Cantidad de ventas",
                "Ticket promedio",
                "Curso con mayor ingreso neto",
                "Vendedor con mayor ingreso neto",
            ],
            "Valor": [
                trabajo["monto"].sum(),
                trabajo["comision_bs"].sum(),
                trabajo["ingreso_neto_bs"].sum(),
                len(trabajo),
                trabajo["monto"].mean(),
                curso_top,
                vendedor_top,
            ],
        }
    )

    salida = BytesIO()

    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        dashboard.to_excel(writer, sheet_name="Dashboard", index=False, startrow=1)
        ventas.to_excel(writer, sheet_name="Ventas", index=False)
        resumen_curso.to_excel(writer, sheet_name="Resumen por curso", index=False)
        resumen_vendedor.to_excel(writer, sheet_name="Resumen vendedores", index=False)
        resumen_diario.to_excel(writer, sheet_name="Resumen diario", index=False)
        resumen_mensual.to_excel(writer, sheet_name="Resumen mensual", index=False)

        for curso in COMISIONES:
            ventas[ventas["Curso"] == curso].to_excel(
                writer,
                sheet_name=curso,
                index=False,
            )

    salida.seek(0)
    wb = load_workbook(salida)

    azul = "1F4E78"
    azul_claro = "D9EAF7"
    blanco = "FFFFFF"
    borde = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    for ws in wb.worksheets:
        fila_encabezado = 2 if ws.title == "Dashboard" else 1
        ws.freeze_panes = f"A{fila_encabezado + 1}"
        ws.auto_filter.ref = ws.dimensions

        for celda in ws[fila_encabezado]:
            celda.fill = PatternFill("solid", fgColor=azul)
            celda.font = Font(color=blanco, bold=True)
            celda.alignment = Alignment(horizontal="center")
            celda.border = borde

        for fila in ws.iter_rows():
            for celda in fila:
                celda.border = borde

        for columna in ws.columns:
            letra = get_column_letter(columna[0].column)
            ancho = max(
                len(str(celda.value)) if celda.value is not None else 0
                for celda in columna
            )
            ws.column_dimensions[letra].width = min(ancho + 3, 34)

        for fila in ws.iter_rows(min_row=fila_encabezado + 1):
            for celda in fila:
                encabezado = ws.cell(row=fila_encabezado, column=celda.column).value
                if encabezado and any(
                    palabra in str(encabezado)
                    for palabra in ["Monto", "Ingreso", "Comision", "Valor", "Venta"]
                ):
                    if isinstance(celda.value, (int, float)):
                        celda.number_format = '#,##0.00'

    ws_dashboard = wb["Dashboard"]
    ws_dashboard["A1"] = "DASHBOARD EJECUTIVO DE VENTAS"
    ws_dashboard.merge_cells("A1:B1")
    ws_dashboard["A1"].font = Font(size=16, bold=True, color=blanco)
    ws_dashboard["A1"].fill = PatternFill("solid", fgColor=azul)
    ws_dashboard["A1"].alignment = Alignment(horizontal="center")

    for fila in range(3, ws_dashboard.max_row + 1):
        ws_dashboard[f"A{fila}"].fill = PatternFill("solid", fgColor=azul_claro)
        ws_dashboard[f"A{fila}"].font = Font(bold=True)

    ws_curso = wb["Resumen por curso"]
    grafico = BarChart()
    grafico.title = "Ingreso neto por curso"
    grafico.y_axis.title = "Bolivianos"
    grafico.x_axis.title = "Curso"
    datos = Reference(ws_curso, min_col=4, min_row=1, max_row=ws_curso.max_row)
    categorias = Reference(ws_curso, min_col=1, min_row=2, max_row=ws_curso.max_row)
    grafico.add_data(datos, titles_from_data=True)
    grafico.set_categories(categorias)
    grafico.height = 8
    grafico.width = 14
    ws_curso.add_chart(grafico, "G2")

    final = BytesIO()
    wb.save(final)
    final.seek(0)
    return final.getvalue()


# ============================================================
# INTERFAZ
# ============================================================

crear_base_datos()
asegurar_administrador_inicial()

if "usuario_actual" not in st.session_state:
    st.session_state.usuario_actual = None


def cerrar_sesion():
    st.session_state.usuario_actual = None
    st.rerun()


if st.session_state.usuario_actual is None:
    izquierda, centro, derecha = st.columns([1, 1.05, 1])

    with centro:
        if LOGO_PATH.exists():
            logo_col_1, logo_col_2, logo_col_3 = st.columns([0.45, 1, 0.45])
            with logo_col_2:
                st.image(str(LOGO_PATH), use_container_width=True)

        st.markdown(
            '<div class="jet-brand-title">Sistema de Gestión de Ventas</div>',
            unsafe_allow_html=True,
        )
        st.markdown(
            '<div class="jet-brand-subtitle">'
            'Jet Airways Academy · Acceso exclusivo para usuarios autorizados'
            '</div>',
            unsafe_allow_html=True,
        )

        with st.form("formulario_login"):
            usuario_login = st.text_input(
                "Usuario",
                placeholder="Ingresa tu nombre de usuario",
            )
            password_login = st.text_input(
                "Contraseña",
                type="password",
                placeholder="Ingresa tu contraseña",
            )
            entrar = st.form_submit_button(
                "Iniciar sesión",
                type="primary",
                use_container_width=True,
            )

        if entrar:
            datos_usuario = autenticar_usuario(usuario_login, password_login)
            if datos_usuario is None:
                st.error("Usuario o contraseña incorrectos.")
            else:
                st.session_state.usuario_actual = datos_usuario
                st.rerun()

        with st.expander("Acceso inicial"):
            st.write("Usuario: `admin`")
            st.write("Contraseña: `Admin123!`")
            st.warning("Cambia esta contraseña después del primer ingreso.")

        st.markdown(
            '<div class="jet-footer">'
            'Jet Airways Academy · Sistema interno de gestión © 2026'
            '</div>',
            unsafe_allow_html=True,
        )

    st.stop()

usuario_actual = st.session_state.usuario_actual
rol_actual = usuario_actual["rol"]

cabecera_1, cabecera_2 = st.columns([0.15, 0.85])
with cabecera_1:
    if LOGO_PATH.exists():
        st.image(str(LOGO_PATH), use_container_width=True)
with cabecera_2:
    st.title("Jet Airways Academy")
    st.subheader("Sistema Profesional de Gestión de Ventas")
st.caption(
    f"Sesión: **{usuario_actual['nombre']}** · "
    f"Usuario: **{usuario_actual['usuario']}** · "
    f"Rol: **{rol_actual}**"
)

with st.sidebar:
    st.header("Navegación")

    opciones = ["Registrar venta", "Dashboard", "Consultar ventas"]

    if rol_actual in ["Administrador", "Supervisor"]:
        opciones.append("Exportar y respaldo")

    if rol_actual == "Administrador":
        opciones.extend(
            [
                "Editar o eliminar",
                "Importar Excel",
                "Usuarios",
                "Configuración",
            ]
        )

    seccion = st.radio("Selecciona una sección", opciones)
    st.divider()
    st.write(f"**{usuario_actual['nombre']}**")
    st.caption(rol_actual)
    if st.button("Cerrar sesión", use_container_width=True):
        cerrar_sesion()

# ------------------------------------------------------------
# REGISTRAR
# ------------------------------------------------------------

if seccion == "Registrar venta":
    st.subheader("Registrar nueva venta")

    with st.form("registro_venta", clear_on_submit=True):
        c1, c2 = st.columns(2)

        with c1:
            fecha = st.date_input("Fecha", value=date.today())
            curso = st.selectbox("Curso", list(COMISIONES.keys()))
            monto = st.number_input(
                "Monto de la venta (Bs)",
                min_value=0.01,
                value=1000.00,
                step=10.00,
            )

        with c2:
            cliente = st.text_input("Cliente")
            vendedor = st.selectbox(
                "Vendedor",
                VENDEDORES_PREDETERMINADOS + ["Otro"],
            )
            vendedor_otro = ""
            if vendedor == "Otro":
                vendedor_otro = st.text_input("Nombre del vendedor")

            tipo_cambio = st.number_input(
                "Tipo de cambio USD/BOB",
                min_value=0.01,
                value=float(TIPO_CAMBIO_USD_BS),
                step=0.01,
            )

        datos_comision = COMISIONES[curso]
        comision_estimada = (
            datos_comision["monto"] * tipo_cambio
            if datos_comision["moneda"] == "USD"
            else datos_comision["monto"]
        )
        neto_estimado = monto - comision_estimada

        m1, m2, m3 = st.columns(3)
        m1.metric("Ingreso bruto", f"Bs {monto:,.2f}")
        m2.metric("Comisión", f"Bs {comision_estimada:,.2f}")
        m3.metric("Ingreso neto", f"Bs {neto_estimado:,.2f}")

        guardar = st.form_submit_button(
            "Guardar venta",
            type="primary",
            use_container_width=True,
        )

        if guardar:
            vendedor_final = vendedor_otro.strip() if vendedor == "Otro" else vendedor

            if not cliente.strip():
                st.error("Debes ingresar el nombre del cliente.")
            elif not vendedor_final:
                st.error("Debes ingresar el nombre del vendedor.")
            else:
                comision, neto = guardar_venta(
                    fecha,
                    curso,
                    monto,
                    cliente,
                    vendedor_final,
                    tipo_cambio,
                    usuario_actual["usuario"],
                )
                st.success(
                    f"Venta guardada. Comisión: Bs {comision:,.2f} | "
                    f"Ingreso neto: Bs {neto:,.2f}"
                )
                if neto < 0:
                    st.warning("La comisión supera el monto de la venta.")

# ------------------------------------------------------------
# DASHBOARD
# ------------------------------------------------------------

elif seccion == "Dashboard":
    df = cargar_ventas()

    if rol_actual == "Vendedor" and not df.empty:
        df = df[df["creado_por"] == usuario_actual["usuario"]].copy()

    if df.empty:
        st.info("Todavía no existen ventas registradas.")
        st.stop()

    st.subheader("Dashboard ejecutivo")

    minimo = df["fecha"].min().date()
    maximo = df["fecha"].max().date()

    c1, c2, c3 = st.columns(3)
    with c1:
        fechas = st.date_input(
            "Rango de fechas",
            value=(minimo, maximo),
            min_value=minimo,
            max_value=maximo,
        )
    with c2:
        cursos = st.multiselect(
            "Cursos",
            sorted(df["curso"].unique()),
            default=sorted(df["curso"].unique()),
        )
    with c3:
        vendedores = st.multiselect(
            "Vendedores",
            sorted(df["vendedor"].unique()),
            default=sorted(df["vendedor"].unique()),
        )

    if isinstance(fechas, tuple) and len(fechas) == 2:
        inicio, fin = fechas
    else:
        inicio = fin = fechas

    filtrado = df[
        (df["fecha"].dt.date >= inicio)
        & (df["fecha"].dt.date <= fin)
        & (df["curso"].isin(cursos))
        & (df["vendedor"].isin(vendedores))
    ].copy()

    if filtrado.empty:
        st.warning("No existen datos para los filtros seleccionados.")
        st.stop()

    total_bruto = filtrado["monto"].sum()
    comisiones = filtrado["comision_bs"].sum()
    neto = filtrado["ingreso_neto_bs"].sum()
    cantidad = len(filtrado)
    promedio = filtrado["monto"].mean()

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Ingreso bruto", f"Bs {total_bruto:,.2f}")
    k2.metric("Comisiones", f"Bs {comisiones:,.2f}")
    k3.metric("Ingreso neto", f"Bs {neto:,.2f}")
    k4.metric("Ventas", f"{cantidad:,}")
    k5.metric("Ticket promedio", f"Bs {promedio:,.2f}")

    por_curso = (
        filtrado.groupby("curso", as_index=False)
        .agg(
            Ingreso_bruto=("monto", "sum"),
            Ingreso_neto=("ingreso_neto_bs", "sum"),
            Ventas=("id", "count"),
        )
        .sort_values("Ingreso_neto", ascending=False)
    )

    por_vendedor = (
        filtrado.groupby("vendedor", as_index=False)
        .agg(
            Ingreso_neto=("ingreso_neto_bs", "sum"),
            Ventas=("id", "count"),
        )
        .sort_values("Ingreso_neto", ascending=False)
    )

    g1, g2 = st.columns(2)

    with g1:
        fig = px.bar(
            por_curso,
            x="curso",
            y="Ingreso_neto",
            title="Ingreso neto por curso",
            labels={"curso": "Curso", "Ingreso_neto": "Ingreso neto (Bs)"},
        )
        st.plotly_chart(fig, use_container_width=True)

    with g2:
        fig = px.pie(
            por_curso,
            names="curso",
            values="Ingreso_neto",
            title="Participación del ingreso neto",
            hole=0.4,
        )
        st.plotly_chart(fig, use_container_width=True)

    tendencia = (
        filtrado.assign(Fecha=filtrado["fecha"].dt.date)
        .groupby("Fecha", as_index=False)["ingreso_neto_bs"]
        .sum()
    )

    fig = px.line(
        tendencia,
        x="Fecha",
        y="ingreso_neto_bs",
        markers=True,
        title="Evolución diaria del ingreso neto",
        labels={"ingreso_neto_bs": "Ingreso neto (Bs)"},
    )
    st.plotly_chart(fig, use_container_width=True)

    g3, g4 = st.columns(2)
    with g3:
        st.write("#### Ranking de cursos")
        st.dataframe(por_curso, use_container_width=True, hide_index=True)
    with g4:
        st.write("#### Ranking de vendedores")
        st.dataframe(por_vendedor, use_container_width=True, hide_index=True)

# ------------------------------------------------------------
# CONSULTAR
# ------------------------------------------------------------

elif seccion == "Consultar ventas":
    df = cargar_ventas()

    if rol_actual == "Vendedor" and not df.empty:
        df = df[df["creado_por"] == usuario_actual["usuario"]].copy()
    st.subheader("Consultar y filtrar ventas")

    if df.empty:
        st.info("No existen ventas registradas.")
        st.stop()

    c1, c2, c3 = st.columns(3)
    with c1:
        curso_filtro = st.selectbox(
            "Curso",
            ["Todos"] + sorted(df["curso"].unique().tolist()),
        )
    with c2:
        vendedor_filtro = st.selectbox(
            "Vendedor",
            ["Todos"] + sorted(df["vendedor"].unique().tolist()),
        )
    with c3:
        cliente_busqueda = st.text_input("Buscar cliente")

    filtrado = df.copy()

    if curso_filtro != "Todos":
        filtrado = filtrado[filtrado["curso"] == curso_filtro]

    if vendedor_filtro != "Todos":
        filtrado = filtrado[filtrado["vendedor"] == vendedor_filtro]

    if cliente_busqueda.strip():
        filtrado = filtrado[
            filtrado["cliente"].str.contains(
                cliente_busqueda.strip(),
                case=False,
                na=False,
            )
        ]

    tabla = filtrado.rename(
        columns={
            "id": "ID",
            "fecha": "Fecha",
            "curso": "Curso",
            "monto": "Monto",
            "cliente": "Cliente",
            "vendedor": "Vendedor",
            "comision_bs": "Comisión Bs",
            "ingreso_neto_bs": "Ingreso neto Bs",
            "creado_por": "Registrado por",
        }
    )[
        [
            "ID",
            "Fecha",
            "Curso",
            "Monto",
            "Cliente",
            "Vendedor",
            "Comisión Bs",
            "Ingreso neto Bs",
            "Registrado por",
        ]
    ]

    st.dataframe(
        tabla,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Monto": st.column_config.NumberColumn(format="Bs %.2f"),
            "Comisión Bs": st.column_config.NumberColumn(format="Bs %.2f"),
            "Ingreso neto Bs": st.column_config.NumberColumn(format="Bs %.2f"),
        },
    )

# ------------------------------------------------------------
# EDITAR / ELIMINAR
# ------------------------------------------------------------

elif seccion == "Editar o eliminar":
    df = cargar_ventas()
    st.subheader("Editar o eliminar una venta")

    if df.empty:
        st.info("No existen ventas registradas.")
        st.stop()

    opciones = {
        f"ID {int(fila.id)} | {fila.fecha.date()} | {fila.curso} | "
        f"Bs {fila.monto:,.2f} | {fila.cliente}": int(fila.id)
        for fila in df.itertuples()
    }

    seleccion = st.selectbox("Selecciona una venta", list(opciones.keys()))
    id_venta = opciones[seleccion]
    registro = df[df["id"] == id_venta].iloc[0]

    with st.form("editar_venta"):
        c1, c2 = st.columns(2)

        with c1:
            fecha = st.date_input("Fecha", value=registro["fecha"].date())
            curso = st.selectbox(
                "Curso",
                list(COMISIONES.keys()),
                index=list(COMISIONES.keys()).index(registro["curso"]),
            )
            monto = st.number_input(
                "Monto",
                min_value=0.01,
                value=float(registro["monto"]),
                step=10.0,
            )

        with c2:
            cliente = st.text_input("Cliente", value=registro["cliente"])
            vendedor = st.text_input("Vendedor", value=registro["vendedor"])
            tipo_cambio = st.number_input(
                "Tipo de cambio",
                min_value=0.01,
                value=float(registro["tipo_cambio"]),
                step=0.01,
            )

        actualizar = st.form_submit_button(
            "Actualizar venta",
            type="primary",
            use_container_width=True,
        )

        if actualizar:
            if not cliente.strip() or not vendedor.strip():
                st.error("Cliente y vendedor son obligatorios.")
            else:
                actualizar_venta(
                    id_venta,
                    fecha,
                    curso,
                    monto,
                    cliente,
                    vendedor,
                    tipo_cambio,
                )
                st.success("Venta actualizada correctamente.")

    st.divider()
    confirmar = st.checkbox("Confirmo que deseo eliminar esta venta")

    if st.button(
        "Eliminar venta",
        disabled=not confirmar,
        use_container_width=True,
    ):
        eliminar_venta(id_venta)
        st.success("Venta eliminada correctamente.")

# ------------------------------------------------------------
# IMPORTAR
# ------------------------------------------------------------

elif seccion == "Importar Excel":
    st.subheader("Importar ventas desde Excel")
    st.write(
        "El archivo debe incluir: Fecha, Curso, Monto, Cliente y Vendedor. "
        "También se acepta una columna llamada Tipo_ingreso."
    )

    tipo_cambio = st.number_input(
        "Tipo de cambio que se aplicará a las ventas importadas",
        min_value=0.01,
        value=float(TIPO_CAMBIO_USD_BS),
        step=0.01,
    )

    archivo = st.file_uploader(
        "Selecciona el Excel",
        type=["xlsx", "xlsm"],
    )

    if archivo is not None:
        vista = pd.read_excel(archivo, sheet_name="Ventas")
        st.write("Vista previa")
        st.dataframe(vista.head(20), use_container_width=True)

        if st.button("Importar ventas", type="primary"):
            try:
                archivo.seek(0)
                datos = pd.read_excel(archivo, sheet_name="Ventas")
                cantidad, errores = importar_dataframe(datos, tipo_cambio)
                st.success(f"Se importaron {cantidad} ventas correctamente.")

                if not errores.empty:
                    st.warning(
                        f"{len(errores)} filas no se importaron por contener errores."
                    )
                    st.dataframe(errores, use_container_width=True)
            except Exception as error:
                st.error(f"No se pudo importar el archivo: {error}")

# ------------------------------------------------------------
# EXPORTAR / RESPALDO
# ------------------------------------------------------------

elif seccion == "Exportar y respaldo":
    st.subheader("Exportar reportes y crear respaldos")
    df = cargar_ventas()

    if df.empty:
        st.info("No existen ventas para exportar.")
        st.stop()

    reporte = generar_reporte_excel(df)

    st.download_button(
        "Descargar reporte profesional en Excel",
        data=reporte,
        file_name="reporte_profesional_ventas_cursos.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

    csv = df.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "Descargar respaldo CSV",
        data=csv,
        file_name="respaldo_ventas.csv",
        mime="text/csv",
        use_container_width=True,
    )

    with open(DB_PATH, "rb") as archivo_db:
        st.download_button(
            "Descargar copia completa de la base de datos",
            data=archivo_db.read(),
            file_name="ventas_respaldo.db",
            mime="application/octet-stream",
            use_container_width=True,
        )

# ------------------------------------------------------------
# CONFIGURACIÓN
# ------------------------------------------------------------

elif seccion == "Usuarios":
    st.subheader("Administración de usuarios")

    tab_crear, tab_estado, tab_password = st.tabs(
        ["Crear usuario", "Activar o bloquear", "Cambiar contraseña"]
    )

    with tab_crear:
        with st.form("crear_usuario_form"):
            nombre_nuevo = st.text_input("Nombre completo")
            usuario_nuevo = st.text_input("Nombre de usuario")
            password_nuevo = st.text_input("Contraseña", type="password")
            rol_nuevo = st.selectbox("Rol", ROLES)
            crear_nuevo = st.form_submit_button(
                "Crear usuario",
                type="primary",
                use_container_width=True,
            )

            if crear_nuevo:
                try:
                    crear_usuario(
                        nombre_nuevo,
                        usuario_nuevo,
                        password_nuevo,
                        rol_nuevo,
                    )
                    st.success("Usuario creado correctamente.")
                except ValueError as error:
                    st.error(str(error))

    usuarios_df = cargar_usuarios()
    mapa_usuarios = {
        f"{fila.nombre} ({fila.usuario}) · {fila.rol}": int(fila.id)
        for fila in usuarios_df.itertuples()
    }

    with tab_estado:
        st.dataframe(
            usuarios_df.rename(
                columns={
                    "id": "ID",
                    "nombre": "Nombre",
                    "usuario": "Usuario",
                    "rol": "Rol",
                    "activo": "Activo",
                    "creado_en": "Creado",
                }
            ),
            use_container_width=True,
            hide_index=True,
        )

        seleccion_estado = st.selectbox(
            "Selecciona un usuario",
            list(mapa_usuarios.keys()),
            key="seleccion_estado_usuario",
        )
        id_estado = mapa_usuarios[seleccion_estado]
        fila_estado = usuarios_df[usuarios_df["id"] == id_estado].iloc[0]
        estado_actual = "Activo" if int(fila_estado["activo"]) == 1 else "Bloqueado"
        nuevo_estado = st.selectbox(
            "Estado",
            ["Activo", "Bloqueado"],
            index=0 if estado_actual == "Activo" else 1,
        )

        if st.button("Guardar estado", use_container_width=True):
            if id_estado == usuario_actual["id"] and nuevo_estado == "Bloqueado":
                st.error("No puedes bloquear tu propia cuenta mientras la usas.")
            else:
                cambiar_estado_usuario(id_estado, nuevo_estado == "Activo")
                st.success("Estado actualizado correctamente.")

    with tab_password:
        seleccion_password = st.selectbox(
            "Selecciona un usuario",
            list(mapa_usuarios.keys()),
            key="seleccion_password_usuario",
        )
        id_password = mapa_usuarios[seleccion_password]
        nueva_password = st.text_input("Nueva contraseña", type="password")
        confirmar_password = st.text_input(
            "Confirmar nueva contraseña",
            type="password",
        )

        if st.button(
            "Cambiar contraseña",
            type="primary",
            use_container_width=True,
        ):
            if nueva_password != confirmar_password:
                st.error("Las contraseñas no coinciden.")
            else:
                try:
                    cambiar_password_usuario(id_password, nueva_password)
                    st.success("Contraseña actualizada correctamente.")
                except ValueError as error:
                    st.error(str(error))


elif seccion == "Configuración":
    st.subheader("Configuración actual")

    st.write(f"**Tipo de cambio predeterminado:** {TIPO_CAMBIO_USD_BS}")

    tabla_comisiones = pd.DataFrame(
        [
            {
                "Curso": curso,
                "Comisión": datos["monto"],
                "Moneda": datos["moneda"],
                "Comisión estimada en Bs": (
                    datos["monto"] * TIPO_CAMBIO_USD_BS
                    if datos["moneda"] == "USD"
                    else datos["monto"]
                ),
            }
            for curso, datos in COMISIONES.items()
        ]
    )

    st.dataframe(tabla_comisiones, use_container_width=True, hide_index=True)

    st.info(
        "Para cambiar permanentemente cursos, comisiones, vendedores o el tipo "
        "de cambio predeterminado, edita la sección CONFIGURACIÓN al inicio de app.py."
    )

    st.warning(
        "El acceso inicial es admin / Admin123!. Cambia esa contraseña desde Usuarios."
    )
