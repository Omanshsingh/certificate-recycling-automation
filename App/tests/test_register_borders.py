"""Border repairs preserve the XLSX package, including cached formula values."""
from copy import copy
from io import BytesIO
import unittest
import xml.etree.ElementTree as ET
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.styles import Border, Font, PatternFill

from certificate_core import batch_content, read_batches
from plant_workbook_template import plant_workbook_bytes
from workbook_store import repair_register_borders, sort_register_by_inward_date


NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'


def package(data, changes):
    output = BytesIO()
    with ZipFile(BytesIO(data)) as source, ZipFile(output, 'w') as target:
        for info in source.infolist():
            target.writestr(info, changes.get(info.filename, source.read(info.filename)))
    return output.getvalue()


class RegisterBorderTests(unittest.TestCase):
    def fixture(self):
        book = load_workbook(BytesIO(plant_workbook_bytes('plant1')))
        sheet = book.active
        for col, value in enumerate(['1', '2026-10-01', 'Client', 'Pune', 'Laptop', 1, 2, 2,
                                     'F6-001', '=SUM(H2:H4)', 'MANUAL-REF'], 1):
            sheet.cell(2, col).value = value
        sheet['E4'] = 'Monitor'
        sheet['H4'] = 6
        for col in ('A', 'B', 'C', 'D', 'I', 'J', 'K', 'P', 'Q', 'R'):
            sheet.merge_cells(f'{col}2:{col}4')
        sheet['A3'].font = Font(name='Arial', italic=True, color='FF800080')
        sheet['A3'].fill = PatternFill('solid', fgColor='FFFFDDFF')
        # Mimic old files: no operational-column border or bottom-edge cell.
        for row in range(2, 5):
            for col in range(12, 16):
                sheet.cell(row, col).border = Border()
        output = BytesIO()
        book.save(output)
        book.close()
        data = output.getvalue()
        with ZipFile(BytesIO(data)) as archive:
            xml = archive.read('xl/worksheets/sheet1.xml').decode()
            styles = archive.read('xl/styles.xml').decode()
        xml = xml.replace('<f>SUM(H2:H4)</f><v></v>', '<f>SUM(H2:H4)</f><v>8</v>')
        # Excel emits these flags, but openpyxl's style serializer drops them.
        styles = styles.replace('pivotButton="0"', 'applyNumberFormat="1" applyFont="1" applyFill="0" pivotButton="0"')
        return package(data, {'xl/worksheets/sheet1.xml': xml.encode(), 'xl/styles.xml': styles.encode()})

    def test_every_physical_merged_edge_and_detail_column_is_bordered(self):
        data = repair_register_borders(self.fixture())
        with ZipFile(BytesIO(data)) as archive:
            sheet = ET.fromstring(archive.read('xl/worksheets/sheet1.xml'))
            styles = ET.fromstring(archive.read('xl/styles.xml'))
        cells = {c.attrib['r']: c for c in sheet.iter(NS + 'c')}
        xfs = list(styles.find(NS + 'cellXfs'))
        borders = list(styles.find(NS + 'borders'))

        def edge(coord, name):
            xf = xfs[int(cells[coord].attrib['s'])]
            return borders[int(xf.attrib['borderId'])].find(NS + name).attrib.get('style')

        for row in range(2, 5):
            for col in 'ABCDEFGHIJKLMNOPQR':
                self.assertEqual(edge(f'{col}{row}', 'left'), 'medium', f'{col}{row}')
                self.assertEqual(edge(f'{col}{row}', 'right'), 'medium', f'{col}{row}')
        for col in 'ABCDEFGHIJKLMNOPQR':
            self.assertEqual(edge(f'{col}2', 'top'), 'medium')
            self.assertEqual(edge(f'{col}4', 'bottom'), 'medium')
        self.assertIsNone(edge('A3', 'bottom'))
        self.assertEqual(edge('E3', 'bottom'), 'thin')
        for col in 'PQ':
            for row in (2, 3):
                self.assertEqual(edge(f'{col}{row}', 'bottom'), 'thin')
                self.assertEqual(edge(f'{col}{row+1}', 'top'), 'thin')
        merges = {area.attrib['ref'] for area in sheet.iter(NS + 'mergeCell')}
        self.assertNotIn('P2:P4', merges)
        self.assertNotIn('Q2:Q4', merges)
        self.assertIn('A2:A4', merges)

    def test_only_border_styles_change_and_formula_caches_survive(self):
        data = self.fixture()
        updated = repair_register_borders(data)
        with ZipFile(BytesIO(data)) as before, ZipFile(BytesIO(updated)) as after:
            for name in before.namelist():
                if name not in {'xl/styles.xml', 'xl/worksheets/sheet1.xml'}:
                    self.assertEqual(before.read(name), after.read(name), name)
            old_sheet = ET.fromstring(before.read('xl/worksheets/sheet1.xml'))
            new_sheet = ET.fromstring(after.read('xl/worksheets/sheet1.xml'))
            old_styles = ET.fromstring(before.read('xl/styles.xml'))
            new_styles = ET.fromstring(after.read('xl/styles.xml'))
            old_cells = {c.attrib['r']: c for c in old_sheet.iter(NS + 'c')}
            new_cells = {c.attrib['r']: c for c in new_sheet.iter(NS + 'c')}
            for coord, cell in old_cells.items():
                saved = new_cells[coord]
                old_xf = copy(old_styles.find(NS + 'cellXfs')[int(cell.attrib.get('s', 0))])
                new_xf = copy(new_styles.find(NS + 'cellXfs')[int(saved.attrib.get('s', 0))])
                old_xf.attrib.pop('borderId', None)
                new_xf.attrib.pop('borderId', None)
                self.assertEqual(ET.tostring(old_xf), ET.tostring(new_xf), coord)
            for root in (old_sheet, new_sheet):
                for row in root.find(NS + 'sheetData'):
                    for cell in list(row):
                        cell.attrib.pop('s', None)
                        if not len(cell) and set(cell.attrib) == {'r'}:
                            row.remove(cell)
            # Tracking columns intentionally become separate physical cells;
            # all other merges and worksheet payloads remain unchanged.
            merges = old_sheet.find(NS + 'mergeCells')
            for area in list(merges):
                if area.attrib['ref'] in {'P2:P4', 'Q2:Q4'}:
                    merges.remove(area)
            merges.attrib['count'] = str(len(merges))
            self.assertEqual(ET.tostring(old_sheet), ET.tostring(new_sheet))
        before = read_batches(BytesIO(data))
        after = read_batches(BytesIO(updated))
        self.assertEqual([batch_content(b) for b in before], [batch_content(b) for b in after])
        self.assertEqual(after[0]['total_weight'], 8)
        self.assertEqual(repair_register_borders(updated), updated)

    def test_missing_blank_rows_get_complete_merged_edges(self):
        data = self.fixture()
        with ZipFile(BytesIO(data)) as archive:
            root = ET.fromstring(archive.read('xl/worksheets/sheet1.xml'))
        rows = root.find(NS + 'sheetData')
        rows.remove(next(row for row in rows if row.attrib['r'] == '3'))
        data = package(data, {'xl/worksheets/sheet1.xml': ET.tostring(root)})
        updated = repair_register_borders(data)
        with ZipFile(BytesIO(updated)) as archive:
            root = ET.fromstring(archive.read('xl/worksheets/sheet1.xml'))
        self.assertEqual(len(next(row for row in root.find(NS + 'sheetData') if row.attrib['r'] == '3')), 18)
        self.assertEqual(repair_register_borders(updated), updated)

    def test_sorting_retains_the_repair(self):
        updated = sort_register_by_inward_date(self.fixture())
        self.assertEqual(repair_register_borders(updated), updated)


if __name__ == '__main__':
    unittest.main()
