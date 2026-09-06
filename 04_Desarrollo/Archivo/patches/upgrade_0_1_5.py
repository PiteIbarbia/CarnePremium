from __future__ import annotations

from datetime import datetime
from pathlib import Path
import hashlib
import json
import shutil
import os
import tempfile

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo


# ============================================================
# CONFIGURACIÓN
# ============================================================

PROJECT_ROOT = (
    Path(__file__)
    .resolve()
    .parents[2]
)

MASTER_FILE = (
    PROJECT_ROOT
    / "01_ERP"
    / "CarnePremiumERP.xlsx"
)

BACKUP_DIR = (
    PROJECT_ROOT
    / "02_Backups"
)

STATE_FILE = (
    MASTER_FILE.parent
    / ".build_state.json"
)

VERSION = "0.1.5"

BUILD_NAME = (
    "MVP 0.1.5 - Ofertas automáticas en Ventas"
)

MAX_ROWS = 500


# ============================================================
# HASH
# ============================================================

def sha256_file(path: Path) -> str:

    sha = hashlib.sha256()

    with path.open(
        "rb"
    ) as file:

        for chunk in iter(
            lambda: file.read(
                1024 * 1024
            ),
            b""
        ):

            sha.update(
                chunk
            )

    return sha.hexdigest()


# ============================================================
# BACKUP
# ============================================================

def crear_backup():

    if not MASTER_FILE.exists():

        return None

    BACKUP_DIR.mkdir(
        parents=True,
        exist_ok=True
    )

    timestamp = datetime.now().strftime(
        "%Y%m%d_%H%M%S"
    )

    backup = (
        BACKUP_DIR
        / (
            "CarnePremiumERP_before_"
            f"v{VERSION}_"
            f"{timestamp}.xlsx"
        )
    )

    shutil.copy2(
        MASTER_FILE,
        backup
    )

    return backup


# ============================================================
# ENCABEZADOS
# ============================================================

def set_headers(
    ws,
    headers
):

    for column, header in enumerate(
        headers,
        start=1
    ):

        cell = ws.cell(
            1,
            column
        )

        cell.value = header

        cell.font = Font(
            bold=True,
            color="FFFFFF"
        )

        cell.fill = PatternFill(
            fill_type="solid",
            fgColor="7A3E13"
        )

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )


# ============================================================
# 23_OFERTAS
# ============================================================

def preparar_ofertas(
    workbook
):

    if (
        "23_OFERTAS"
        in workbook.sheetnames
    ):

        ws = workbook[
            "23_OFERTAS"
        ]

    else:

        ws = workbook.create_sheet(
            "23_OFERTAS"
        )

    headers = [

        "OfertaID",

        "ProductoID",

        "Producto",

        "NombreOferta",

        "PrecioOferta",

        "Descuento%",

        "FechaInicio",

        "FechaFin",

        "MinKg",

        "Estado",

        "Prioridad",

    ]

    set_headers(
        ws,
        headers
    )

    # No inventamos ofertas reales.
    # Solamente dejamos la estructura.

    for row in range(
        2,
        MAX_ROWS + 1
    ):

        ws.cell(
            row,
            5
        ).number_format = (
            '$ #,##0.00'
        )

        ws.cell(
            row,
            6
        ).number_format = (
            '0.0%'
        )

        ws.cell(
            row,
            7
        ).number_format = (
            'dd/mm/yyyy'
        )

        ws.cell(
            row,
            8
        ).number_format = (
            'dd/mm/yyyy'
        )

        ws.cell(
            row,
            9
        ).number_format = (
            '0.000'
        )

    widths = {

        "A": 16,
        "B": 16,
        "C": 30,
        "D": 30,
        "E": 18,
        "F": 14,
        "G": 14,
        "H": 14,
        "I": 12,
        "J": 14,
        "K": 12,
    }

    for column, width in widths.items():

        ws.column_dimensions[
            column
        ].width = width

    # Tabla
    existing = [
        table.name
        for table
        in ws.tables.values()
    ]

    if "OfertasTable" in existing:

        del ws.tables[
            "OfertasTable"
        ]

    table = Table(
        displayName="OfertasTable",
        ref="A1:K501"
    )

    table.tableStyleInfo = (
        TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False
        )
    )

    ws.add_table(
        table
    )

    ws.freeze_panes = "A2"


# ============================================================
# VENTAS
# ============================================================

def preparar_ventas(
    workbook
):

    ws = workbook[
        "05_VENTAS"
    ]

    headers = [

        "VentaID",
        "Fecha",
        "Hora",

        "Cliente",
        "TipoPrecio",
        "Producto",
        "Kg",

        "PrecioKg",
        "Importe",
        "Ganancia$",
        "Ganancia%",

        "MedioPago",
        "Vendedor",
        "Observaciones",

        "_ClienteID",
        "_ProductoID",

        "EstadoVenta",

        "TipoAplicado",
        "MotivoPrecio",
        "OfertaID",
        "OfertaNombre",

    ]

    set_headers(
        ws,
        headers
    )

    # --------------------------------------------------------
    # IMPORTANTE
    # --------------------------------------------------------
    #
    # Cliente y Producto:
    #   se escriben por nombre.
    #
    # O/P:
    #   son IDs técnicos.
    #
    # R/U:
    #   información comercial.
    # --------------------------------------------------------

    for row in range(
        2,
        MAX_ROWS + 1
    ):

        # VentaID
        ws.cell(
            row,
            1
        ).value = (
            f'=IF(OR(D{row}="",'
            f'F{row}=""),"",'
            f'"VEN"&TEXT('
            f'ROW()-1,'
            f'"000000"))'
        )

        # Fecha
        ws.cell(
            row,
            2
        ).value = (
            f'=IF(A{row}="","",TODAY())'
        )

        # Hora
        ws.cell(
            row,
            3
        ).value = (
            f'=IF(A{row}="","",NOW())'
        )

        # ----------------------------------------------------
        # ClienteID
        # ----------------------------------------------------

        ws.cell(
            row,
            15
        ).value = (

            f'=IF(D{row}="","",'
            f'IFERROR('
            f'INDEX('
            f'\'03_CLIENTES\'!'
            f'$A$2:$A$501,'
            f'MATCH('
            f'D{row},'
            f'\'03_CLIENTES\'!'
            f'$B$2:$B$501,'
            f'0'
            f')'
            f'),'
            f'""))'
        )

        # ----------------------------------------------------
        # ProductoID
        # ----------------------------------------------------

        ws.cell(
            row,
            16
        ).value = (

            f'=IF(F{row}="","",'
            f'IFERROR('
            f'INDEX('
            f'\'04_PRODUCTOS\'!'
            f'$A$2:$A$501,'
            f'MATCH('
            f'F{row},'
            f'\'04_PRODUCTOS\'!'
            f'$B$2:$B$501,'
            f'0'
            f')'
            f'),'
            f'""))'
        )

        # ----------------------------------------------------
        # Precio BASE
        # ----------------------------------------------------

        retail = (

            f'INDEX('
            f'\'04_PRODUCTOS\'!'
            f'$E$2:$E$501,'
            f'MATCH('
            f'F{row},'
            f'\'04_PRODUCTOS\'!'
            f'$B$2:$B$501,'
            f'0'
            f')'
            f')'
        )

        wholesale = (

            f'INDEX('
            f'\'04_PRODUCTOS\'!'
            f'$F$2:$F$501,'
            f'MATCH('
            f'F{row},'
            f'\'04_PRODUCTOS\'!'
            f'$B$2:$B$501,'
            f'0'
            f')'
            f')'
        )

        special = (

            f'INDEX('
            f'\'04_PRODUCTOS\'!'
            f'$G$2:$G$501,'
            f'MATCH('
            f'F{row},'
            f'\'04_PRODUCTOS\'!'
            f'$B$2:$B$501,'
            f'0'
            f')'
            f')'
        )

        ws.cell(
            row,
            8
        ).value = (

            f'=IFERROR('

            f'IF(E{row}="Mayorista",'
            f'IF('
            f'{wholesale}="",'
            f'{retail},'
            f'{wholesale}'
            f'),'

            f'IF(E{row}="Especial",'
            f'IF('
            f'{special}="",'
            f'{retail},'
            f'{special}'
            f'),'

            f'{retail}'
            f')'
            f'),'
            f'"")'
        )

        # Importe
        ws.cell(
            row,
            9
        ).value = (

            f'=IF('
            f'AND('
            f'G{row}<>"",'
            f'H{row}<>""'
            f'),'
            f'G{row}*H{row},'
            f'"")'
        )

        # Ganancia
        cost = (

            f'INDEX('
            f'\'04_PRODUCTOS\'!'
            f'$D$2:$D$501,'
            f'MATCH('
            f'F{row},'
            f'\'04_PRODUCTOS\'!'
            f'$B$2:$B$501,'
            f'0'
            f')'
            f')'
        )

        ws.cell(
            row,
            10
        ).value = (

            f'=IFERROR('
            f'IF('
            f'{cost}="",'
            f'"",'
            f'I{row}-'
            f'(G{row}*{cost})'
            f'),'
            f'"")'
        )

        # Ganancia %
        ws.cell(
            row,
            11
        ).value = (

            f'=IF('
            f'OR('
            f'I{row}="",'
            f'J{row}=""'
            f'),'
            f'"",'
            f'IFERROR('
            f'J{row}/I{row},'
            f'"")'
            f')'
        )

        # ----------------------------------------------------
        # Estado
        # ----------------------------------------------------

        ws.cell(
            row,
            17
        ).value = (

            f'=IF('
            f'AND('
            f'D{row}<>"",'
            f'F{row}<>"",'
            f'G{row}>0,'
            f'H{row}<>""'
            f'),'
            f'IF('
            f'O{row}="",'
            f'"CLIENTE NO ENCONTRADO",'
            f'IF('
            f'P{row}="",'
            f'"PRODUCTO NO ENCONTRADO",'
            f'"OK"'
            f')'
            f'),'
            f'""'
            f')'
        )

        # ----------------------------------------------------
        # Formatos
        # ----------------------------------------------------

        ws.cell(
            row,
            7
        ).number_format = (
            "0.000"
        )

        for column in range(
            8,
            11
        ):

            ws.cell(
                row,
                column
            ).number_format = (
                '$ #,##0.00'
            )

        ws.cell(
            row,
            11
        ).number_format = (
            "0.0%"
        )

        ws.cell(
            row,
            2
        ).number_format = (
            "dd/mm/yyyy"
        )

        ws.cell(
            row,
            3
        ).number_format = (
            "hh:mm:ss"
        )

    # --------------------------------------------------------
    # La columna Cliente y Producto
    # se mantiene libre para escritura por nombre.
    # --------------------------------------------------------

    # IDs técnicos ocultos.
    ws.column_dimensions[
        "O"
    ].hidden = True

    ws.column_dimensions[
        "P"
    ].hidden = True

    widths = {

        "A": 15,
        "B": 12,
        "C": 10,
        "D": 30,
        "E": 15,
        "F": 30,
        "G": 10,
        "H": 16,
        "I": 16,
        "J": 16,
        "K": 12,
        "L": 20,
        "M": 20,
        "N": 32,
        "O": 16,
        "P": 16,
        "Q": 24,
        "R": 18,
        "S": 32,
        "T": 16,
        "U": 30,
    }

    for column, width in widths.items():

        ws.column_dimensions[
            column
        ].width = width

    # Tabla
    existing = [
        table.name
        for table
        in ws.tables.values()
    ]

    if "VentasTable" in existing:

        del ws.tables[
            "VentasTable"
        ]

    table = Table(
        displayName="VentasTable",
        ref="A1:U501"
    )

    table.tableStyleInfo = (
        TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False
        )
    )

    ws.add_table(
        table
    )

    ws.freeze_panes = "A2"


# ============================================================
# DASHBOARD
# ============================================================

def preparar_dashboard(
    workbook
):

    if (
        "10_DASHBOARD"
        in workbook.sheetnames
    ):

        ws = workbook[
            "10_DASHBOARD"
        ]

    else:

        ws = workbook.create_sheet(
            "10_DASHBOARD"
        )

    try:

        ws.merge_cells(
            "A1:B1"
        )

    except ValueError:

        pass

    ws["A1"] = (
        "CARNE PREMIUM ERP - "
        "DASHBOARD OPERATIVO"
    )

    ws["A1"].font = Font(
        bold=True,
        size=16,
        color="FFFFFF"
    )

    ws["A1"].fill = PatternFill(
        fill_type="solid",
        fgColor="7A3E13"
    )

    ws["A1"].alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    metricas = [

        (
            "Ventas registradas",
            '=COUNTIF('
            '\'05_VENTAS\'!'
            '$A$2:$A$501,'
            '"<>"'
            ')'
        ),

        (
            "Facturación",
            '=SUM('
            '\'05_VENTAS\'!'
            '$I$2:$I$501'
            ')'
        ),

        (
            "Ganancia",
            '=SUM('
            '\'05_VENTAS\'!'
            '$J$2:$J$501'
            ')'
        ),

        (
            "Margen",
            '=IFERROR('
            'B6/B5,'
            '"")'
        ),

        (
            "Kg vendidos",
            '=SUM('
            '\'05_VENTAS\'!'
            '$G$2:$G$501'
            ')'
        ),

        (
            "Ticket promedio",
            '=IFERROR('
            'B5/B4,'
            '"")'
        ),

    ]

    # Las filas se escriben desde 3.
    for index, (
        label,
        formula
    ) in enumerate(
        metricas,
        start=3
    ):

        ws.cell(
            index,
            1
        ).value = label

        ws.cell(
            index,
            2
        ).value = formula

    ws["B5"].number_format = (
        '$ #,##0.00'
    )

    ws["B6"].number_format = (
        '$ #,##0.00'
    )

    ws["B7"].number_format = (
        '0.0%'
    )

    ws["B8"].number_format = (
        '0.000'
    )

    ws["B9"].number_format = (
        '$ #,##0.00'
    )

    ws.column_dimensions[
        "A"
    ].width = 30

    ws.column_dimensions[
        "B"
    ].width = 20


# ============================================================
# VALIDAR
# ============================================================

def validar(
    workbook
):

    requeridas = [

        "03_CLIENTES",
        "04_PRODUCTOS",
        "05_VENTAS",
        "23_OFERTAS",
        "10_DASHBOARD",
    ]

    for sheet in requeridas:

        if sheet not in (
            workbook.sheetnames
        ):

            raise RuntimeError(
                f"Falta hoja: {sheet}"
            )

    ventas = workbook[
        "05_VENTAS"
    ]

    headers = [
        cell.value
        for cell
        in ventas[1]
    ]

    requeridos_ventas = [

        "Cliente",
        "TipoPrecio",
        "Producto",
        "Kg",
        "PrecioKg",
        "Importe",
        "Ganancia$",
        "Ganancia%",
        "_ClienteID",
        "_ProductoID",
        "EstadoVenta",
        "TipoAplicado",
        "MotivoPrecio",
        "OfertaID",
        "OfertaNombre",
    ]

    for header in (
        requeridos_ventas
    ):

        if header not in headers:

            raise RuntimeError(
                "05_VENTAS: falta "
                f"{header}"
            )

    # Fórmulas existentes.
    for cell in [
        "A2",
        "B2",
        "C2",
        "H2",
        "I2",
        "J2",
        "K2",
        "O2",
        "P2",
        "Q2",
    ]:

        value = ventas[
            cell
        ].value

        if (
            not isinstance(
                value,
                str
            )
            or not value.startswith(
                "="
            )
        ):

            raise RuntimeError(
                f"05_VENTAS!{cell} "
                "no contiene fórmula."
            )


# ============================================================
# GUARDAR
# ============================================================

def guardar_seguro(
    workbook
):

    token = (
        f"{VERSION}-"
        f"{datetime.now().strftime('%Y%m%d%H%M%S')}"
    )

    temp_fd, temp_name = (
        tempfile.mkstemp(
            prefix="CarnePremiumERP_015_",
            suffix=".xlsx",
            dir=str(
                MASTER_FILE.parent
            )
        )
    )

    version_sheet = workbook["12_VERSION"]

    version_sheet["A1"] = "Campo"
    version_sheet["B1"] = "Valor"

    version_sheet["A2"] = "Versión"
    version_sheet["B2"] = VERSION

    version_sheet["A3"] = "Build"
    version_sheet["B3"] = BUILD_NAME

    version_sheet["A4"] = "Estado"
    version_sheet["B4"] = "Generado y validado"

    version_sheet["A5"] = "BuildToken"
    version_sheet["B5"] = token

    os.close(
        temp_fd
    )

    temp_path = Path(
        temp_name
    )

    try:

        workbook.save(
            temp_path
        )

        comprobacion = load_workbook(
            temp_path,
            data_only=False
        )

        try:

            validar(
                comprobacion
            )

        finally:

            comprobacion.close()

        backup = crear_backup()

        if backup:

            print(
                f"Backup creado: {backup}"
            )

        temp_path.replace(
            MASTER_FILE
        )

        STATE_FILE.write_text(
            json.dumps(
                {
                    "version":
                        VERSION,

                    "build":
                        BUILD_NAME,

                    "build_token":
                        token,

                    "sha256":
                        sha256_file(
                            MASTER_FILE
                        ),

                    "fecha":
                        datetime.now().isoformat(
                            timespec="seconds"
                        ),
                },

                indent=2,

                ensure_ascii=False
            ),

            encoding="utf-8"
        )

        return token

    except PermissionError as exc:

        raise PermissionError(
            "Windows no permite reemplazar "
            "el maestro. Cerrá Excel."
        ) from exc

    finally:

        if temp_path.exists():

            try:

                temp_path.unlink()

            except PermissionError:

                pass


# ============================================================
# MAIN
# ============================================================

def main():

    print()
    print("=" * 60)
    print(
        "CARNE PREMIUM ERP - UPGRADE 0.1.5"
    )
    print("=" * 60)
    print()

    if not MASTER_FILE.exists():

        raise SystemExit(
            f"No existe:\n{MASTER_FILE}"
        )

    print(
        "[1/4] Abriendo maestro..."
    )

    workbook = load_workbook(
        MASTER_FILE
    )

    try:

        print(
            "[2/4] Preparando 23_OFERTAS..."
        )

        preparar_ofertas(
            workbook
        )

        print(
            "[3/4] Integrando Ventas + Dashboard..."
        )

        preparar_ventas(
            workbook
        )

        preparar_dashboard(
            workbook
        )

        print(
            "[4/4] Validando y guardando..."
        )

        token = guardar_seguro(
            workbook
        )

    finally:

        workbook.close()

    print()
    print("=" * 60)
    print(
        "OK: MVP 0.1.5 INTEGRADO"
    )
    print("=" * 60)
    print()
    print(
        f"Versión: {VERSION}"
    )
    print(
        f"Token: {token}"
    )
    print(
        f"Maestro: {MASTER_FILE}"
    )
    print()


if __name__ == "__main__":
    main()