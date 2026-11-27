"""Remove obsolete app-owned sheet locks without changing workbook content."""

from base64 import b64encode
from hashlib import sha512
from io import BytesIO
from zipfile import ZipFile

from register_reader import load_workbook

from certificate_core import select_register


# A recognizable protection record lets the app maintain its own sheet lock
# without treating a password chosen by the workbook owner as app-owned.
_SALT = b'EcorecoIssuedV1!'
_SPINS = 10_000


def _protection_hash():
    value = sha512(_SALT + 'Ecoreco issued inwards'.encode('utf-16le')).digest()
    for index in range(_SPINS):
        value = sha512(value + index.to_bytes(4, 'little')).digest()
    return b64encode(value).decode('ascii')


_SALT64 = b64encode(_SALT).decode('ascii')
_HASH = _protection_hash()
def _managed(sheet):
    protection = sheet.protection
    return (protection.algorithmName == 'SHA-512' and protection.saltValue == _SALT64
            and protection.spinCount == _SPINS and protection.hashValue == _HASH)


def externally_protected(sheet):
    return sheet.protection.sheet and not _managed(sheet)


def remove_issued_inward_protection(data):
    """Remove only the protection added by this app; leave workbook content intact."""
    from workbook_store import _element, _sheet_member

    book = load_workbook(BytesIO(data), data_only=False)
    try:
        sheet, _, _ = select_register(book)
        if not _managed(sheet):
            return data
        with ZipFile(BytesIO(data)) as source:
            member = _sheet_member(source, sheet.title)
            original = source.read(member).decode('utf-8')
            xml = _element(original, 'sheetProtection').sub('', original)
            xml = _element(xml, 'protectedRanges').sub('', xml)
            if xml == original:
                return data
            output = BytesIO()
            with ZipFile(output, 'w') as target:
                target.comment = source.comment
                for info in source.infolist():
                    content = xml.encode('utf-8') if info.filename == member else source.read(info.filename)
                    target.writestr(info, content)
            return output.getvalue()
    finally:
        book.close()


def protect_issued_inwards(data):
    """Backward-compatible name: app-issued workbook protection is retired."""
    return remove_issued_inward_protection(data)
