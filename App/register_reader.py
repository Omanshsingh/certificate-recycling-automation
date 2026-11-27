"""Reuse identical merge styles while reading an Excel register."""

from copy import copy

from openpyxl import __version__, load_workbook as _load_workbook
from openpyxl.cell.cell import Cell, MergedCell
from openpyxl.reader.excel import ExcelReader


class _RegisterReader(ExcelReader):
    def read_worksheets(self):
        # Each memo belongs to one workbook because style IDs are workbook
        # specific. Capture styles from openpyxl's own merge routine.
        merge_styles = {}
        create_sheet = self.wb.create_sheet
        sheets = []

        def create_sheet_with_merge_cache(*args, **kwargs):
            sheet = create_sheet(*args, **kwargs)
            clean_merge = sheet._clean_merge_range

            def clean_with_cached_styles(area):
                height = area.max_row - area.min_row + 1
                if (area.min_col != area.max_col or not 2 <= height <= 64
                        or type(area.start_cell) is not Cell):
                    clean_merge(area)
                    return
                # openpyxl has already merged the bottom/right borders into
                # the anchor. The remaining cell styles depend on that anchor
                # and the shape of the merge. The regular routine discards
                # any previous cells below the anchor before styling them.
                key = (height, tuple(area.start_cell._style or ()))
                styles = merge_styles.get(key)
                if styles is None:
                    clean_merge(area)
                    if len(merge_styles) < 512:
                        merge_styles[key] = tuple(
                            copy(sheet._cells[row, area.min_col]._style)
                            for row in range(area.min_row, area.max_row + 1)
                        )
                    return
                anchor_style, *placeholder_styles = styles
                area.start_cell._style = copy(anchor_style)
                column = area.min_col
                for row, style in enumerate(placeholder_styles, area.min_row + 1):
                    cell = MergedCell(sheet, row, column)
                    cell._style = copy(style)
                    sheet._cells[row, column] = cell

            sheet._clean_merge_range = clean_with_cached_styles
            sheets.append(sheet)
            return sheet

        self.wb.create_sheet = create_sheet_with_merge_cache
        try:
            super().read_worksheets()
        finally:
            del self.wb.create_sheet
            for sheet in sheets:
                del sheet._clean_merge_range


def load_workbook(filename, read_only=False, keep_vba=False, data_only=False,
                  keep_links=True, rich_text=False):
    if read_only or not __version__.startswith('3.1.'):
        return _load_workbook(filename, read_only=read_only, keep_vba=keep_vba,
                              data_only=data_only, keep_links=keep_links, rich_text=rich_text)
    reader = _RegisterReader(filename, read_only, keep_vba, data_only, keep_links, rich_text)
    try:
        reader.read()
    finally:
        reader.archive.close()
    return reader.wb
