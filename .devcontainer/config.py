from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "ventas.db"
LOGO_PATH = BASE_DIR / "assets" / "logo_jet_airways.png"
LOGO_ICON_PATH = BASE_DIR / "assets" / "logo_isotipo.png"
BACKGROUND_PATH = BASE_DIR / "assets" / "fondo_avion.jpg"

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

VENDEDORES_PREDETERMINADOS = ["Ana", "Carlos", "Daniela", "José", "María"]
ROLES = ["Administrador", "Supervisor", "Vendedor"]
