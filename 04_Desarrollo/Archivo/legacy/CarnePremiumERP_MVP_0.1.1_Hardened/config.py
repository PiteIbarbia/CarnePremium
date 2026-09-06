from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[2]
MASTER_FILE = PROJECT_ROOT / "01_ERP" / "CarnePremiumERP.xlsx"

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
