from pathlib import Path


# ============================================================
# RAÍZ DEL PROYECTO
# ============================================================

PROJECT_ROOT = Path(__file__).resolve().parents[2]


# ============================================================
# CARPETAS
# ============================================================

ERP_DIR = PROJECT_ROOT / "01_ERP"
BACKUP_DIR = PROJECT_ROOT / "02_Backups"
DOCS_DIR = PROJECT_ROOT / "03_Documentacion"
SCRIPTS_DIR = PROJECT_ROOT / "04_Desarrollo" / "Scripts"
TESTS_DIR = PROJECT_ROOT / "04_Desarrollo" / "Pruebas"
EXPORT_DIR = PROJECT_ROOT / "05_Exportaciones"


# ============================================================
# ARCHIVO MAESTRO
# ============================================================

MASTER_FILE = (
    ERP_DIR /
    "CarnePremiumERP.xlsx"
)


# ============================================================
# VERSIONAMIENTO
# ============================================================

CURRENT_VERSION = "0.3.6"


# ============================================================
# CANTIDAD DE FILAS OPERATIVAS
# ============================================================

MAX_ROWS = 500


# ============================================================
# HOJAS DEL ERP
# ============================================================

SHEETS = [

    "00_INICIO",

    "01_PARAMETROS",

    "02_CATALOGOS",

    "03_CLIENTES",

    "04_PRODUCTOS",

    "05_VENTAS",

    "06_TICKET",

    "07_CAJA",

    "08_STOCK",

    "09_CTACTE",

    "10_DASHBOARD",

    "11_AUDITORIA",

    "12_VERSION",

    "13_LOG",

]