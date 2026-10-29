"""Files the live tests upload, built from the standard library so no fixture is checked in."""

from __future__ import annotations

import io
import os
import random
import struct
import zipfile
import zlib


def png(size: int = 24) -> bytes:
    """A small solid-colour PNG."""
    def chunk(kind: bytes, data: bytes) -> bytes:
        body = kind + data
        return struct.pack(">I", len(data)) + body + struct.pack(">I", zlib.crc32(body) & 0xFFFFFFFF)

    row = b"\x00" + bytes([0x2B, 0x6C, 0xB0]) * size
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", size, size, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(row * size)) + chunk(b"IEND", b""))


def random_bytes(n: int) -> bytes:
    return os.urandom(n)


_CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
       '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
       '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
       '<Default Extension="xml" ContentType="application/xml"/>{overrides}</Types>')
_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
         '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
         'Target="{target}"/></Relationships>')


def docx(text: str = "Live test document", *, padding_bytes: int = 0) -> bytes:
    """A minimal Word document. `padding_bytes` adds that much stored (uncompressed) paragraph text, so the
    file is about that big and still a valid document."""
    paragraphs = [f"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>"]
    if padding_bytes:
        rng = random.Random(7)
        letters = "abcdefghijklmnopqrstuvwxyz     "
        per = 400
        for _ in range(padding_bytes // (per + 40)):
            paragraphs.append("<w:p><w:r><w:t>" + "".join(rng.choice(letters) for _ in range(per)) + "</w:t></w:r></w:p>")
    document = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
                + "".join(paragraphs) + "</w:body></w:document>")
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_STORED if padding_bytes else zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", _CT.format(overrides=(
            '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.'
            'wordprocessingml.document.main+xml"/>')))
        z.writestr("_rels/.rels", _RELS.format(target="word/document.xml"))
        z.writestr("word/document.xml", document)
    return out.getvalue()


def xlsx(seed: str = "seed") -> bytes:
    """A workbook with one sheet, Sheet1, holding one text cell."""
    workbook = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>')
    workbook_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                     '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                     '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
                     'Target="worksheets/sheet1.xml"/></Relationships>')
    sheet = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'
             f'<row r="1"><c r="A1" t="inlineStr"><is><t>{seed}</t></is></c></row></sheetData></worksheet>')
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", _CT.format(overrides=(
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.'
            'spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.'
            'spreadsheetml.worksheet+xml"/>')))
        z.writestr("_rels/.rels", _RELS.format(target="xl/workbook.xml"))
        z.writestr("xl/workbook.xml", workbook)
        z.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        z.writestr("xl/worksheets/sheet1.xml", sheet)
    return out.getvalue()
