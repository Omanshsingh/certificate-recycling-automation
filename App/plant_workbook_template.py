"""Create a clean plant register using the columns expected by the app."""
from __future__ import annotations

from io import BytesIO
import os
from pathlib import Path
import uuid

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

from certificate_core import read_batches


HEADERS = (
    "SN", "Inward Date", "Supplier Name", "Location", "Item",
    "Quantity in Pcs", "Per Pcs Weight in (KG)", "Weight (KG)",
    "Form 6 No", "Total Weight in KG", "Reference No.",
    "Comments", "Transporter/Vehicle Details", "Docket No./Vehicle No",
    "Weight Slip/Weight Scale", "COD No", "COD Date",
    "Certificate Issued",
)
WIDTHS = (8, 13, 22, 18, 19, 13, 15, 13, 14, 15, 16, 18, 26, 24, 24, 12, 13, 18)


def plant_workbook_bytes(plant: str) -> bytes:
    if plant not in {"plant1", "plant2"}:
        raise ValueError("Choose Plant 1 or Plant 2.")

    book = Workbook()
    sheet = book.active
    sheet.title = "Sheet1"
    sheet.sheet_view.showGridLines = True
    sheet.freeze_panes = "E2"
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    sheet.print_title_rows = "1:1"
    sheet.append(HEADERS)
    sheet.row_dimensions[1].height = 78

    # Match the real plant register's turquoise header and crisp black grid.
    header_fill = PatternFill("solid", fgColor="FF4DD0E1")
    black = "FF000000"
    header_edge = Side(style="thin", color=black)
    for index, width in enumerate(WIDTHS, 1):
        column = sheet.column_dimensions[sheet.cell(1, index).column_letter]
        column.width = width
        column.alignment = Alignment(wrap_text=True)
        cell = sheet.cell(1, index)
        cell.font = Font(name="Calibri", size=11, bold=True, color=black)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(left=header_edge, right=header_edge,
                             top=header_edge, bottom=header_edge)

    medium = Side(style="medium", color=black)
    thin = Side(style="thin", color=black)
    # One clean, fully bordered blank row is the entry point for direct Excel
    # edits. The app expands and groups it when a client has multiple items.
    for row in range(2, 3):
        for column in range(1, len(HEADERS) + 1):
            cell = sheet.cell(row, column)
            cell.font = Font(name="Calibri", size=11, color=black)
            cell.alignment = Alignment(horizontal="left", vertical="top",
                                      wrap_text=True)
            cell.border = Border(
                left=medium if column == 1 else thin,
                right=medium if column == len(HEADERS) else thin,
                top=medium,
                bottom=medium,
            )
            if column in {2, 17}:
                cell.number_format = "dd/mm/yy"
            if column in {6, 7, 8, 10}:
                cell.number_format = "0.00"

    status_validation = DataValidation(
        type="list", formula1='",Ready,Alert,Generated,Completed,N/A"',
        allow_blank=True,
    )
    status_validation.errorTitle = "Choose a listed status"
    status_validation.error = "Select a status from the list, or leave the cell blank."
    status_validation.showErrorMessage = True
    sheet.add_data_validation(status_validation)
    status_validation.add("R2:R10000")
    for value, fill, font in (
        ("READY", "FFB6D7A8", "FF006100"),
        ("ALERT", "FFC00000", "FFFFFFFF"),
        ("GENERATED", "FFFFE699", "FF7F6000"),
        ("GENERATED EXTERNALLY", "FFDDEBF7", "FF1F4E78"),
        ("COMPLETED", "FFFFE699", "FF7F6000"),
        ("N/A", "FFFFFFFF", "FF000000"),
    ):
        rule = FormulaRule(
            formula=[f'UPPER(TRIM($R2))="{value}"'],
            fill=PatternFill("solid", fgColor=fill),
            font=Font(color=font, bold=True),
        )
        sheet.conditional_formatting.add("R2:R10000", rule)

    output = BytesIO()
    book.save(output)
    book.close()
    data = output.getvalue()
    if read_batches(BytesIO(data)):
        raise ValueError("A new plant template must start with no client records.")
    return data


def create_plant_workbook(path: str | Path, plant: str) -> Path:
    destination = Path(path).expanduser().resolve()
    if destination.suffix.casefold() != ".xlsx":
        raise ValueError("Save the new plant register as an .xlsx file.")
    if not destination.parent.is_dir():
        raise ValueError("Choose an existing folder for the new workbook.")

    data = plant_workbook_bytes(plant)
    temporary = destination.with_name(f".{destination.stem}.{uuid.uuid4().hex}.tmp")
    try:
        with temporary.open("xb") as stream:
            stream.write(data)
            stream.flush()
            os.fsync(stream.fileno())
        # On Windows os.rename fails when destination already exists, so an
        # existing register can never be silently overwritten.
        os.rename(temporary, destination)
        read_batches(destination)
        return destination
    except FileExistsError as exc:
        raise ValueError("That file already exists. Choose a new name to protect the existing workbook.") from exc
    except Exception:
        destination.unlink(missing_ok=True)
        raise
    finally:
        temporary.unlink(missing_ok=True)
