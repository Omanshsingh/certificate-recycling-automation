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
    "Quantity in Pcs", "Weight (KG)", "Form 6 No", "Total Weight in KG",
    "Comments", "Transporter/Vehicle Details", "Docket No./Vehicle No",
    "Weight Slip/Weight Scale", "Reference No.", "COD No", "COD Date",
    "Certificate Issued",
)
WIDTHS = (12, 16, 30, 25, 28, 16, 16, 20, 20, 32, 30, 24, 24, 24, 14, 14, 20)


def plant_workbook_bytes(plant: str) -> bytes:
    if plant not in {"plant1", "plant2"}:
        raise ValueError("Choose Plant 1 or Plant 2.")

    book = Workbook()
    sheet = book.active
    sheet.title = "Sheet1"
    sheet.sheet_view.showGridLines = False
    sheet.freeze_panes = "E2"
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    sheet.print_title_rows = "1:1"
    sheet.append(HEADERS)
    sheet.row_dimensions[1].height = 42

    header_fill = PatternFill("solid", fgColor="087F5B")
    white = "FFFFFFFF"
    for index, width in enumerate(WIDTHS, 1):
        sheet.column_dimensions[sheet.cell(1, index).column_letter].width = width
        cell = sheet.cell(1, index)
        cell.font = Font(name="Calibri", size=11, bold=True, color=white)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=Side(style="medium", color="006B4D"))

    medium = Side(style="medium", color="7F9A92")
    thin = Side(style="thin", color="D6E0DD")
    material_columns = {5, 6, 7}
    for row in range(2, 5):
        sheet.row_dimensions[row].height = 24
        for column in range(1, len(HEADERS) + 1):
            cell = sheet.cell(row, column)
            cell.font = Font(name="Calibri", size=11, color="20312D")
            cell.alignment = Alignment(vertical="center", wrap_text=column in {3, 4, 5, 10, 11, 12, 13})
            cell.border = Border(
                left=medium if column == 1 else thin,
                right=medium if column == len(HEADERS) else thin,
                top=medium if row == 2 else thin,
                bottom=medium if row == 4 else thin,
            )
            if column in {2, 16}:
                cell.number_format = "dd.mm.yyyy"
            if column in {6, 7, 9}:
                cell.number_format = "0.00"
        if row == 2:
            sheet.cell(row, 3).font = Font(name="Calibri", size=11, color="20312D")

    # Leave a styled three-line client block as the app's new-client pattern.
    # The rows contain no sample data; the first real client starts at row 2.
    for column in range(1, len(HEADERS) + 1):
        if column not in material_columns:
            sheet.merge_cells(start_row=2, start_column=column,
                              end_row=4, end_column=column)
            left = medium if column == 1 else thin
            right = medium if column == len(HEADERS) else thin
            sheet.cell(2, column).border = Border(left=left, right=right, top=medium)
            sheet.cell(3, column).border = Border(left=left, right=right)
            sheet.cell(4, column).border = Border(left=left, right=right, bottom=medium)

    status_validation = DataValidation(
        type="list", formula1='"Ready,Alert,Generated,Generated Manual,Completed,N/A"',
        allow_blank=True,
    )
    status_validation.errorTitle = "Choose a listed status"
    status_validation.error = "Select a status from the list, or leave the cell blank."
    status_validation.showErrorMessage = True
    sheet.add_data_validation(status_validation)
    status_validation.add("Q2:Q10000")
    for value, fill, font in (
        ("READY", "FFB6D7A8", "FF006100"),
        ("ALERT", "FFFCE4D6", "FF7F6000"),
        ("GENERATED", "FFFFE699", "FF7F6000"),
        ("GENERATED MANUAL", "FFDDEBF7", "FF1F4E78"),
        ("COMPLETED", "FFFFE699", "FF7F6000"),
        ("N/A", "FFFFFFFF", "FF000000"),
    ):
        rule = FormulaRule(
            formula=[f'UPPER(TRIM($Q2))="{value}"'],
            fill=PatternFill("solid", fgColor=fill),
            font=Font(color=font, bold=True),
        )
        sheet.conditional_formatting.add("Q2:Q10000", rule)

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
