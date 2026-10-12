from __future__ import annotations

from datetime import datetime
from pathlib import Path
import hashlib
import json
import os
import shutil
import tempfile
import uuid

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

from config import MASTER_FILE, SHEETS

from dashboard import preparar_dashboard
from caja import sincronizar_ventas_en_caja
from cuenta_corriente import sincronizar_cuenta_corriente
from ticket import preparar_ticket
from ventas_excel import procesar_todas_las_ventas


# ============================================================
# VERSION
# ============================================================

BUILD_VERSION = "0.4.0"

BUILD_NAME = (
    "Build 0.4.0 - Ventas + Ticket automático + Cobros + Cuenta Corriente"
)

MAX_ROWS = 500

STATE_FILE = (
    MASTER_FILE.parent
    / ".build_state.json"
)

HEADER_FILL = "7A3E13"
HEADER_FONT = "FFFFFF"


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
# WORKBOOK
# ============================================================

def ensure_workbook(
    path: Path
):

    if path.exists():

        return load_workbook(
            path
        )

    workbook = Workbook()

    workbook.remove(
        workbook.active
    )

    return workbook


# ============================================================
# SHEETS
# ============================================================

def ensure_sheet(
    workbook,
    name: str
):

    if name not in (
        workbook.sheetnames
    ):

        workbook.create_sheet(
            name
        )

    return workbook[name]


# ============================================================
# HEADERS
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
            row=1,
            column=column
        )

        cell.value = header

        cell.font = Font(
            bold=True,
            color=HEADER_FONT
        )

        cell.fill = PatternFill(
            fill_type="solid",
            fgColor=HEADER_FILL
        )

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )


# ============================================================
# VALIDACIONES
# ============================================================

def clear_validations(
    ws
):

    ws.data_validations.dataValidation = []


def add_list_validation(
    ws,
    target_range: str,
    formula1: str
):

    validation = DataValidation(
        type="list",
        allow_blank=True
    )

    validation.formula1 = formula1

    validation.errorTitle = (
        "Valor no válido"
    )

    validation.error = (
        "Seleccione un valor permitido."
    )

    validation.promptTitle = (
        "Carne Premium ERP"
    )

    validation.prompt = (
        "Seleccione un valor."
    )

    ws.add_data_validation(
        validation
    )

    validation.add(
        target_range
    )

    return validation


# ============================================================
# TABLAS
# ============================================================

def ensure_table(
    ws,
    name: str,
    reference: str
):

    existing = [
        table.name
        for table
        in ws.tables.values()
    ]

    if name in existing:

        del ws.tables[name]

    table = Table(
        displayName=name,
        ref=reference
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


# ============================================================
# PARAMETROS
# ============================================================

def setup_parameters(
    workbook
):

    ws = ensure_sheet(
        workbook,
        "01_PARAMETROS"
    )

    set_headers(
        ws,
        [
            "Parametro",
            "Valor"
        ]
    )

    defaults = [

        (
            "Empresa",
            "Carne Premium"
        ),

        (
            "Moneda",
            "ARS"
        ),

        (
            "TipoPrecioDefault",
            "Minorista"
        ),

        (
            "MedioPagoDefault",
            "Efectivo"
        ),

        (
            "IVA",
            21
        ),
    ]

    for row, (
        key,
        value
    ) in enumerate(
        defaults,
        start=2
    ):

        ws.cell(
            row,
            1
        ).value = key

        if ws.cell(
            row,
            2
        ).value in (
            None,
            ""
        ):

            ws.cell(
                row,
                2
            ).value = value

    ws.column_dimensions[
        "A"
    ].width = 28

    ws.column_dimensions[
        "B"
    ].width = 24


# ============================================================
# CATALOGOS
# ============================================================

def setup_catalogs(
    workbook
):

    ws = ensure_sheet(
        workbook,
        "02_CATALOGOS"
    )

    set_headers(
        ws,
        [
            "TiposPrecio",
            "TiposCliente",
            "MediosPago",
            "Estados"
        ]
    )

    values = [

        [
            "Minorista",
            "Minorista",
            "Efectivo",
            "Activo"
        ],

        [
            "Mayorista",
            "Mayorista",
            "Transferencia",
            "Inactivo"
        ],

        [
            "Especial",
            "Especial",
            "Cuenta Corriente",
            None
        ],

        [
            None,
            None,
            "Tarjeta",
            None
        ],

        [
            None,
            None,
            "Otro",
            None
        ],
    ]

    for row, values_row in enumerate(
        values,
        start=2
    ):

        for column, value in enumerate(
            values_row,
            start=1
        ):

            if ws.cell(
                row,
                column
            ).value in (
                None,
                ""
            ):

                ws.cell(
                    row,
                    column
                ).value = value

    for column, width in {
        "A": 20,
        "B": 20,
        "C": 24,
        "D": 16
    }.items():

        ws.column_dimensions[
            column
        ].width = width


# ============================================================
# CLIENTES
# ============================================================

def setup_clients(
    workbook
):

    ws = ensure_sheet(
        workbook,
        "03_CLIENTES"
    )

    clear_validations(
        ws
    )

    set_headers(
        ws,
        [
            "ClienteID",
            "Nombre",
            "Telefono",
            "Direccion",
            "TipoCliente",
            "Estado",
            "Observaciones"
        ]
    )

    for row in range(
        2,
        MAX_ROWS + 1
    ):

        ws.cell(
            row,
            1
        ).value = (
            f'=IF(B{row}="","",'
            f'"CLI"&TEXT('
            f'ROW()-1,'
            f'"000000"))'
        )

        if ws.cell(
            row,
            6
        ).value in (
            None,
            ""
        ):

            ws.cell(
                row,
                6
            ).value = "Activo"

    add_list_validation(
        ws,
        "E2:E501",
        "'02_CATALOGOS'!$B$2:$B$4"
    )

    add_list_validation(
        ws,
        "F2:F501",
        "'02_CATALOGOS'!$D$2:$D$3"
    )

    ensure_table(
        ws,
        "ClientesTable",
        "A1:G501"
    )

    for column, width in {
        "A": 15,
        "B": 30,
        "C": 18,
        "D": 30,
        "E": 18,
        "F": 14,
        "G": 32
    }.items():

        ws.column_dimensions[
            column
        ].width = width

    ws.freeze_panes = "A2"


# ============================================================
# PRODUCTOS
# ============================================================

def setup_products(
    workbook
):

    ws = ensure_sheet(
        workbook,
        "04_PRODUCTOS"
    )

    clear_validations(
        ws
    )

    set_headers(
        ws,
        [
            "ProductoID",
            "Producto",
            "Categoria",
            "Costo",
            "PrecioMinorista",
            "PrecioMayorista",
            "PrecioEspecial",
            "Estado",
            "Observaciones"
        ]
    )

    for row in range(
        2,
        MAX_ROWS + 1
    ):

        ws.cell(
            row,
            1
        ).value = (
            f'=IF(B{row}="","",'
            f'"PRO"&TEXT('
            f'ROW()-1,'
            f'"000000"))'
        )

        if ws.cell(
            row,
            8
        ).value in (
            None,
            ""
        ):

            ws.cell(
                row,
                8
            ).value = "Activo"

        for column in range(
            4,
            8
        ):

            ws.cell(
                row,
                column
            ).number_format = (
                '$ #,##0.00'
            )

    add_list_validation(
        ws,
        "H2:H501",
        "'02_CATALOGOS'!$D$2:$D$3"
    )

    ensure_table(
        ws,
        "ProductosTable",
        "A1:I501"
    )

    for column, width in {
        "A": 15,
        "B": 30,
        "C": 20,
        "D": 16,
        "E": 18,
        "F": 18,
        "G": 18,
        "H": 14,
        "I": 32
    }.items():

        ws.column_dimensions[
            column
        ].width = width

    ws.freeze_panes = "A2"


# ============================================================
# VENTAS
# ============================================================

def setup_sales(
    workbook
):

    ws = ensure_sheet(
        workbook,
        "05_VENTAS"
    )

    clear_validations(
        ws
    )

    # Migración segura desde Build 0.3.6: W antes era TicketID manual.
    # Si Caja ya tenía un cobro automático asociado a la venta, conserva su
    # importe como PagoRecibido. No deduce el pago a partir de la venta bruta.
    legacy_ticket_column = ws.cell(1, 23).value == "TicketID"
    if legacy_ticket_column:
        prior_cash_by_sale = {}
        if "07_CAJA" in workbook.sheetnames:
            old_cash = workbook["07_CAJA"]
            old_cash_columns = {
                str(old_cash.cell(1, col).value).strip(): col
                for col in range(1, old_cash.max_column + 1)
                if old_cash.cell(1, col).value not in (None, "")
            }
            if {"ReferenciaVenta", "Ingreso", "Origen"}.issubset(old_cash_columns):
                for cash_row in range(2, old_cash.max_row + 1):
                    if str(old_cash.cell(cash_row, old_cash_columns["Origen"]).value or "").strip() != "VENTA_AUTO":
                        continue
                    ref = str(old_cash.cell(cash_row, old_cash_columns["ReferenciaVenta"]).value or "").strip()
                    income = old_cash.cell(cash_row, old_cash_columns["Ingreso"]).value
                    if ref and isinstance(income, (int, float)) and income > 0:
                        prior_cash_by_sale[ref] = float(income)

        for old_row in range(2, MAX_ROWS + 1):
            old_value = ws.cell(old_row, 23).value
            if old_value in (None, "") or isinstance(old_value, (int, float)):
                continue
            sale_ref = f"VEN{old_row - 1:06d}"
            migrated_receipt = prior_cash_by_sale.get(sale_ref)
            ws.cell(old_row, 23).value = migrated_receipt if migrated_receipt is not None else None

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
        "PrecioEspecialID",
        "PagoRecibido",
        "_TicketLineaKey",
        "EstadoCobro",
        "SaldoTicket"
    ]

    set_headers(
        ws,
        headers
    )

    for row in range(
        2,
        MAX_ROWS + 1
    ):

        # ----------------------------------------------------
        # VentaID
        # ----------------------------------------------------

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

        # ----------------------------------------------------
        # Fecha
        # Fecha/hora calculadas se conservan después de persistirse
        # por Python. Así no cambian en cada reconstrucción del libro.
        date_cell = ws.cell(row, 2)
        if (
            date_cell.value in (None, "")
            or (
                isinstance(date_cell.value, str)
                and date_cell.value.startswith("=")
            )
        ):
            date_cell.value = (
                f'=IF(A{row}="","",TODAY())'
            )

        # ----------------------------------------------------
        # Hora
        # ----------------------------------------------------
        time_cell = ws.cell(row, 3)
        if (
            time_cell.value in (None, "")
            or (
                isinstance(time_cell.value, str)
                and time_cell.value.startswith("=")
            )
        ):
            time_cell.value = (
                f'=IF(A{row}="","",NOW())'
            )

        # ----------------------------------------------------
        # ClienteID interno
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
        # ProductoID interno
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
        # Precio base para Excel
        #
        # El valor definitivo cuando hay oferta se confirmará
        # con el motor comercial al procesar la venta.
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

        # ----------------------------------------------------
        # Importe
        # ----------------------------------------------------

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

        # ----------------------------------------------------
        # Ganancia
        # ----------------------------------------------------

        ws.cell(
            row,
            10
        ).value = (

            f'=IFERROR('

            f'IF('
            f'{cost}="",'
            f'"",'
            f'I{row}-('
            f'G{row}*'
            f'{cost}'
            f')'
            f'),'

            f'"")'
        )

        # ----------------------------------------------------
        # Ganancia %
        # ----------------------------------------------------

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
        # Estado de venta
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

            f'IF('

            f'G{row}<=0,'
            f'"KG INVALIDO",'

            f'"OK"'

            f')'
            f')'
            f'),'

            f'""'

            f')'
        )

        # ----------------------------------------------------
        # Ticket automático por cliente y día. B puede conservar hora
        # en el valor datetime, por eso el agrupamiento compara el día completo.
        ticket_count = (
            f'COUNTIFS($O$2:O{row},O{row},'
            f'$B$2:B{row},">="&INT(B{row}),'
            f'$B$2:B{row},"<"&INT(B{row})+1,'
            f'$Q$2:Q{row},"OK")'
        )
        ws.cell(row, 24).value = (
            f'=IF(OR(A{row}="",B{row}="",O{row}=""),"",'
            f'IF(Q{row}<>"OK","",'
            f'"T-"&TEXT(B{row},"yyyymmdd")&"-"&O{row}&"|"&{ticket_count}))'
        )

        # ----------------------------------------------------
        # Estado y saldo del ticket: ventas menos pagos del mismo cliente/día.
        date_client_criteria = (
            f'$O$2:$O$501,O{row},'
            f'$B$2:$B$501,">="&INT(B{row}),'
            f'$B$2:$B$501,"<"&INT(B{row})+1,'
            f'$Q$2:$Q$501,"OK"'
        )
        ticket_total = f'SUMIFS($I$2:$I$501,{date_client_criteria})'
        received_total = f'SUMIFS($W$2:$W$501,{date_client_criteria})'
        rounded_total = f'ROUND({ticket_total},2)'
        rounded_received = f'ROUND({received_total},2)'

        ws.cell(row, 25).value = (
            f'=IF(A{row}="","",IF(Q{row}<>"OK","",'
            f'IF({rounded_total}<=0,"",'
            f'IF({rounded_received}<=0,"PENDIENTE",'
            f'IF({rounded_received}>{rounded_total},"A FAVOR",'
            f'IF({rounded_received}={rounded_total},"PAGADO / OK","SALDO PENDIENTE"))))))'
        )
        ws.cell(row, 26).value = (
            f'=IF(A{row}="","",IF(Q{row}<>"OK","",'
            f'ROUND({ticket_total}-{received_total},2)))'
        )

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
        ws.cell(row, 23).number_format = '$ #,##0.00'
        ws.cell(row, 25).alignment = Alignment(horizontal="center")
        ws.cell(row, 26).number_format = '$ #,##0.00'

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
    # Validaciones
    #
    # Cliente y Producto:
    # escritura por nombre.
    #
    # TipoPrecio + MedioPago:
    # listas controladas.
    # --------------------------------------------------------

    add_list_validation(
        ws,
        "E2:E501",
        "'02_CATALOGOS'!$A$2:$A$4"
    )

    add_list_validation(
        ws,
        "L2:L501",
        "'02_CATALOGOS'!$C$2:$C$6"
    )

    payment_amount_validation = DataValidation(
        type="decimal",
        operator="greaterThanOrEqual",
        formula1="0",
        allow_blank=True,
    )
    payment_amount_validation.errorTitle = "Importe de cobro inválido"
    payment_amount_validation.error = "Ingresá un importe cobrado igual o mayor que cero."
    payment_amount_validation.promptTitle = "Pago recibido"
    payment_amount_validation.prompt = "Registrá aquí el importe realmente cobrado. El estado y el saldo se calculan automáticamente."
    ws.add_data_validation(payment_amount_validation)
    payment_amount_validation.add("W2:W501")

    # --------------------------------------------------------
    # Tabla
    # --------------------------------------------------------

    ensure_table(
        ws,
        "VentasTable",
        "A1:Z501"
    )

    for column, width in {

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
        "U": 28,
        "V": 20,
        "W": 18,
        "X": 24,
        "Y": 22,
        "Z": 18,

    }.items():

        ws.column_dimensions[
            column
        ].width = width

    # IDs técnicos ocultos
    ws.column_dimensions[
        "O"
    ].hidden = True

    ws.column_dimensions[
        "P"
    ].hidden = True

    # V y X son columnas técnicas; W permite agrupar cortes por ticket.
    ws.column_dimensions["V"].hidden = True
    ws.column_dimensions["X"].hidden = True

    ws.freeze_panes = "A2"


# ============================================================
# INICIO
# ============================================================

def setup_home(
    workbook
):

    ws = ensure_sheet(
        workbook,
        "00_INICIO"
    )

    try:
        ws.merge_cells(
            "A1:B1"
        )
    except ValueError:
        pass

    ws["A1"] = (
        "CARNE PREMIUM ERP"
    )

    ws["A2"] = (
        "MVP 0.1.5"
    )

    rows = [

        (
            "Flujo",
            "Cliente → Producto → "
            "Kg → Motor Comercial → "
            "Precio → Venta"
        ),

        (
            "Paso 1",
            "Escribir nombre del cliente"
        ),

        (
            "Paso 2",
            "Escribir nombre del producto"
        ),

        (
            "Paso 3",
            "Cargar Kg"
        ),

        (
            "Paso 4",
            "El motor determina el precio"
        ),

        (
            "Paso 5",
            "Verificar precio aplicado "
            "y motivo"
        ),

        (
            "Estado",
            "MVP 0.1.5"
        ),

        (
            "Compatibilidad",
            "Excel 2019 + Google Sheets"
        ),
    ]

    for row, (
        key,
        value
    ) in enumerate(
        rows,
        start=4
    ):

        ws.cell(
            row,
            1
        ).value = key

        ws.cell(
            row,
            2
        ).value = value

    ws["A1"].font = Font(
        bold=True,
        color="FFFFFF",
        size=16
    )

    ws["A1"].fill = PatternFill(
        fill_type="solid",
        fgColor=HEADER_FILL
    )

    ws["A1"].alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    ws.column_dimensions[
        "A"
    ].width = 24

    ws.column_dimensions[
        "B"
    ].width = 60


# ============================================================
# VERSION
# ============================================================

def setup_version(
    workbook,
    token
):

    ws = ensure_sheet(
        workbook,
        "12_VERSION"
    )

    set_headers(
        ws,
        [
            "Campo",
            "Valor"
        ]
    )

    values = [

        (
            "Versión",
            BUILD_VERSION
        ),

        (
            "Build",
            BUILD_NAME
        ),

        (
            "Estado",
            "Generado y validado"
        ),

        (
            "BuildToken",
            token
        ),

        (
            "Generado",
            datetime.now().isoformat(
                timespec="seconds"
            )
        ),

    ]

    for row, (
        key,
        value
    ) in enumerate(
        values,
        start=2
    ):

        ws.cell(
            row,
            1
        ).value = key

        ws.cell(
            row,
            2
        ).value = value

    ws.column_dimensions[
        "A"
    ].width = 22

    ws.column_dimensions[
        "B"
    ].width = 65


# ============================================================
# LOG
# ============================================================

def setup_log(
    workbook,
    token
):

    ws = ensure_sheet(
        workbook,
        "13_LOG"
    )

    set_headers(
        ws,
        [
            "Fecha",
            "Versión",
            "Acción",
            "Detalle"
        ]
    )

    row = (
        ws.max_row
        + 1
    )

    ws.cell(
        row,
        1
    ).value = datetime.now()

    ws.cell(
        row,
        2
    ).value = BUILD_VERSION

    ws.cell(
        row,
        3
    ).value = "BUILD"

    ws.cell(
        row,
        4
    ).value = token


# ============================================================
# VALIDACION
# ============================================================

def validate_before_save(
    workbook,
    token
):

    # Hojas
    for sheet_name in SHEETS:

        if sheet_name not in (
            workbook.sheetnames
        ):

            raise RuntimeError(
                f"Falta hoja: "
                f"{sheet_name}"
            )

    # Encabezados críticos
    sales_headers = [
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
        "PrecioEspecialID",
        "PagoRecibido",
        "_TicketLineaKey",
        "EstadoCobro",
        "SaldoTicket"
    ]

    actual_headers = [
        cell.value
        for cell
        in workbook["05_VENTAS"][1]
    ]

    missing = [
        header
        for header
        in sales_headers
        if header not in actual_headers
    ]

    if missing:

        raise RuntimeError(
            f"05_VENTAS: faltan "
            f"{missing}"
        )

    # Fórmulas
    sales = workbook[
        "05_VENTAS"
    ]

    # A y X se mantienen como fórmulas. Las columnas H:Q
    # pueden ser valores persistidos al procesar ventas con Python.
    for address in ("A2", "X2", "Y2", "Z2"):
        value = sales[address].value
        if not isinstance(value, str) or not value.startswith("="):
            raise RuntimeError(f"Fórmula faltante: 05_VENTAS!{address}")

    row_has_data = any(
        sales.cell(2, column).value not in (None, "")
        for column in (4, 6, 7)
    )
    if row_has_data:
        if sales["Q2"].value in (None, ""):
            raise RuntimeError("05_VENTAS!Q2 no tiene estado de procesamiento.")
    else:
        for address in ("B2", "C2", "H2", "I2", "J2", "K2", "O2", "P2", "Q2"):
            value = sales[address].value
            if not isinstance(value, str) or not value.startswith("="):
                raise RuntimeError(f"Fórmula faltante: 05_VENTAS!{address}")

    # Validaciones
    if len(
        workbook[
            "03_CLIENTES"
        ].data_validations.dataValidation
    ) < 2:

        raise RuntimeError(
            "Validaciones de Clientes insuficientes."
        )

    if len(
        workbook[
            "04_PRODUCTOS"
        ].data_validations.dataValidation
    ) < 1:

        raise RuntimeError(
            "Validación de Productos faltante."
        )

    if len(
        sales.data_validations.dataValidation
    ) < 2:

        raise RuntimeError(
            "Validaciones de Ventas insuficientes."
        )

    # Dashboard
    dashboard = workbook[
        "10_DASHBOARD"
    ]

    for address in [
        "B3",
        "B4",
        "B5",
        "B6",
        "B7",
        "B8",
        "B9"
    ]:

        value = dashboard[
            address
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
                f"Dashboard!{address} "
                f"no tiene fórmula válida."
            )

    # Ticket: verificar encabezados de detalle y fórmulas.
    ticket = workbook["06_TICKET"]
    if [ticket.cell(8, column).value for column in range(1, 5)] != [
        "Corte", "Kg", "Precio por kg", "Importe"
    ]:
        raise RuntimeError("06_TICKET: encabezados de detalle incorrectos.")
    for address in ("A9", "B9", "C9", "D9", "D39"):
        value = ticket[address].value
        if not isinstance(value, str) or not value.startswith("="):
            raise RuntimeError(f"06_TICKET!{address}: fórmula faltante.")

    # Caja: validar estructura sin borrar movimientos existentes.
    caja = workbook["07_CAJA"]
    caja_headers = [
        "MovimientoID", "Fecha", "Tipo", "Concepto", "ReferenciaVenta",
        "MedioPago", "Ingreso", "Egreso", "Saldo", "Origen",
    ]
    if [caja.cell(1, column).value for column in range(1, 11)] != caja_headers:
        raise RuntimeError("07_CAJA: encabezados incorrectos.")

    # Version
    version = workbook[
        "12_VERSION"
    ]

    if version[
        "B2"
    ].value != BUILD_VERSION:

        raise RuntimeError(
            "Versión incorrecta."
        )

    if version[
        "B5"
    ].value != token:

        raise RuntimeError(
            "BuildToken incorrecto."
        )


# ============================================================
# BACKUP
# ============================================================

def create_backup():

    if not MASTER_FILE.exists():

        return None

    backup_dir = (
        MASTER_FILE
        .parent
        .parent
        / "02_Backups"
    )

    backup_dir.mkdir(
        parents=True,
        exist_ok=True
    )

    stamp = datetime.now().strftime(
        "%Y%m%d_%H%M%S"
    )

    backup_path = (
        backup_dir
        / (
            "CarnePremiumERP_before_"
            f"v{BUILD_VERSION}_"
            f"{stamp}.xlsx"
        )
    )

    shutil.copy2(
        MASTER_FILE,
        backup_path
    )

    return backup_path


# ============================================================
# GUARDADO ATOMICO
# ============================================================

def atomic_save(
    workbook
):

    MASTER_FILE.parent.mkdir(
        parents=True,
        exist_ok=True
    )

    # Forzar recálculo de fórmulas auxiliares al abrir en Excel.
    calculation = getattr(workbook, "calculation", None)
    if calculation is not None:
        calculation.fullCalcOnLoad = True
        calculation.forceFullCalc = True
        calculation.calcMode = "auto"

    token = (
        f"{BUILD_VERSION}-"
        f"{datetime.now().strftime('%Y%m%d-%H%M%S')}-"
        f"{uuid.uuid4().hex[:8]}"
    )

    setup_version(
        workbook,
        token
    )

    setup_log(
        workbook,
        token
    )

    validate_before_save(
        workbook,
        token
    )

    temp_fd, temp_name = (
        tempfile.mkstemp(
            prefix="CarnePremiumERP_",
            suffix=".xlsx",
            dir=str(
                MASTER_FILE.parent
            )
        )
    )

    os.close(
        temp_fd
    )

    temp_path = Path(
        temp_name
    )

    try:

        # Guardar
        workbook.save(
            temp_path
        )

        # Reabrir
        check = load_workbook(
            temp_path,
            data_only=False
        )

        try:

            validate_before_save(
                check,
                token
            )

        finally:

            check.close()

        # Backup
        backup = create_backup()

        if backup:

            print(
                f"Backup creado: "
                f"{backup}"
            )

        # Reemplazar maestro
        temp_path.replace(
            MASTER_FILE
        )

        # Estado
        STATE_FILE.write_text(

            json.dumps(

                {
                    "version":
                        BUILD_VERSION,

                    "build":
                        BUILD_NAME,

                    "build_token":
                        token,

                    "master_file":
                        str(
                            MASTER_FILE
                        ),

                    "sha256":
                        sha256_file(
                            MASTER_FILE
                        ),

                    "saved_at":
                        datetime.now().isoformat(
                            timespec="seconds"
                        )
                },

                indent=2,

                ensure_ascii=False
            ),

            encoding="utf-8"
        )

        return token

    except PermissionError as exc:

        raise PermissionError(

            "No se puede reemplazar "
            "CarnePremiumERP.xlsx. "
            "Cerrá Excel y cualquier "
            "programa que lo tenga abierto."
        ) from exc

    finally:

        if temp_path.exists():

            try:

                temp_path.unlink()

            except PermissionError:

                pass


# ============================================================
# BUILD
# ============================================================

def main():

    print()

    print(
        "=" * 60
    )

    print(
        "CARNE PREMIUM ERP"
    )

    print(
        "BUILD 0.3.6"
    )

    print(
        "=" * 60
    )

    print()

    workbook = ensure_workbook(
        MASTER_FILE
    )

    print(
        "[1/9] Parámetros..."
    )

    setup_parameters(
        workbook
    )

    print(
        "[2/9] Catálogos..."
    )

    setup_catalogs(
        workbook
    )

    print(
        "[3/9] Clientes..."
    )

    setup_clients(
        workbook
    )

    print(
        "[4/9] Productos..."
    )

    setup_products(
        workbook
    )

    print(
        "[5/9] Ventas..."
    )

    setup_sales(
        workbook
    )

    print(
        "[6/9] Dashboard..."
    )

    preparar_dashboard(
        workbook
    )

    setup_home(
        workbook
    )

    # ========================================================
    # Procesamiento de ventas existentes
    #
    # IMPORTANTE:
    # Si ya existen ventas reales, el motor puede recalcular
    # precio/motivo/oferta.
    # ========================================================

    print(
        "[7/9] Integrando motor comercial..."
    )

    resultados = (
        procesar_todas_las_ventas(
            workbook
        )
    )

    print(
        f"Ventas procesadas: "
        f"{len(resultados)}"
    )

    print("Sincronizando cobros reales con Caja...")
    cash_result = sincronizar_ventas_en_caja(workbook)
    print(f"Caja: {cash_result}")

    print("Sincronizando Cuenta Corriente...")
    account_result = sincronizar_cuenta_corriente(workbook)
    print(f"Cuenta Corriente: {account_result}")

    print("Preparando Ticket automático...")
    preparar_ticket(workbook)

    print()
    print(
        "Validando y guardando..."
    )

    token = atomic_save(
        workbook
    )

    workbook.close()

    print()
    print(
        "=" * 60
    )

    print(
        "OK: ERP GENERADO CORRECTAMENTE"
    )

    print(
        "=" * 60
    )

    print(
        f"Versión: {BUILD_VERSION}"
    )

    print(
        f"BuildToken: {token}"
    )

    print(
        f"Archivo: {MASTER_FILE}"
    )

    print()

    print(
        "BUILD 0.3.6 INTEGRADO"
    )

    print(
        "Ventas + Ticket + Caja + Ofertas + Dashboard: OK"
    )

    print()


if __name__ == "__main__":

    main()