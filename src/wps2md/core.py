"""Core parser for legacy WPS (.wps) OLE2 Word-binary files.

Reads the OLE2 compound document, validates the FIB, walks PlcfBtePapx →
FKPs to recover paragraph style indices (istd), and exposes structured
output plus a Markdown renderer that respects Heading 1-9 styles.
"""
from __future__ import annotations

import re
import struct
from dataclasses import dataclass, field
from pathlib import Path
from typing import Union

import olefile  # type: ignore[import-untyped]

# ---------------------------------------------------------------------------
# FIB constants (MS-DOC §2.5.1)
# ---------------------------------------------------------------------------
_FIB_MAGIC_WORD97 = 0xA5EC
_FIB_MAGIC_WORD95 = 0xA5DC
_FIB_FLAGS_OFFSET = 0x0A
_FIB_ENCRYPTED_FLAG = 0x0100
_FIB_FCMIN = 0x18
_FIB_FCMAC = 0x1C
_FIB_CCP_TEXT = 0x4C
_FIB_CCP_FTN = 0x50
_FIB_CCP_HDD = 0x54
_FIB_CCP_ATN = 0x5C
_FIB_FC_PLCFBTEPAPX = 0x102
_FIB_LCB_PLCFBTEPAPX = 0x106
_MIN_DOC_SIZE = 0x200
_FKP_SIZE = 512

_U16 = struct.Struct("<H")
_U32 = struct.Struct("<I")

_TRANS = str.maketrans({
    "\x07": "\t", "\x0b": "\n", "\x0c": "\n\n", "\x0d": "\n",
    "\x13": None, "\x14": " ", "\x15": None,
    "\x01": None, "\x08": None, "\x19": None, "\x1e": None, "\x1f": None,
    "\xa0": " ", "\x00": None, "\x7f": None,
})


class WpsParseError(Exception):
    """Raised when the file cannot be parsed as a Word/WPS binary document."""


@dataclass
class Paragraph:
    """A single paragraph with its style and list/table flags.

    Attributes:
        istd: Word style index. Heading 1-9 → 1-9, Normal → 0.
        text: Cleaned paragraph text.
        in_table / is_row_end: From sprmPFInTable (0x2416) and sprmPTtp
            (0x2417). ``is_row_end`` marks the row-terminator paragraph.
        ilfo: 1-based list reference (sprmPIlfo). 0 → not a list item.
        ilvl: 0-based indent level (sprmPIlvl). Defaults to 0.
        list_ordered: True for ordered lists, False for bullets, None when
            the paragraph is not a list item.
    """

    istd: int
    text: str
    in_table: bool = False
    is_row_end: bool = False
    ilfo: int = 0
    ilvl: int = 0
    list_ordered: bool | None = None

    @property
    def heading_level(self) -> int:
        """Return 1-9 for Heading styles, else 0."""
        return self.istd if 1 <= self.istd <= 9 else 0


@dataclass
class WpsDocument:
    """Parsed WPS/.doc document."""

    main_text: str
    paragraphs: list[Paragraph] = field(default_factory=list)
    footnotes: str = ""
    headers_footers: str = ""
    annotations: str = ""
    encoding: str = "utf-16-le"
    num_pages: int | None = None


def _clean(text: str) -> str:
    if not text:
        return ""
    text = text.translate(_TRANS)
    text = re.sub(r"[\x00-\x08\x0e-\x1f\x7f]", "", text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


# Sprm operation codes used here (MS-DOC §2.6.4).
_SPRM_PF_IN_TABLE = 0x2416  # paragraph is inside a table cell
_SPRM_PT_TP = 0x2417        # paragraph terminates a table row
_SPRM_P_ILVL = 0x260A       # list indent level (0..8)
_SPRM_P_ILFO = 0x460B       # 1-based index into PlfLfo (0 = not a list item)

# Numbering Format Codes (MS-DOC §2.9.166). 23/255 = bullet (unordered);
# values 0-7, 22, 45-47 are common ordered formats. We treat 23 as the only
# strictly unordered case and everything else with a list reference as ordered.
_NFC_BULLET = 23
_NFC_NONE = 255


def _read_list_kinds(wd: bytes, table: bytes) -> dict[int, bool]:
    """Return {ilfo: is_ordered} mapping by walking PlcfLst and PlfLfo.

    The returned dict tells whether the list referenced by a 1-based ``ilfo``
    index uses an ordered numbering format (True) or a bullet (False).
    Unknown / unparseable lists are simply absent from the mapping.

    Layout (MS-DOC §2.4.1, §2.9.131, §2.9.150):
        PlcfLst: U16 cLst, then cLst LSTF (28 bytes each), then variable LVLs.
        PlfLfo:  U32 lfoMac, then lfoMac LFO (16 bytes each), then LFOData.
        LSTF.lsid is at offset 0 (I32); rgistdForLst follows; first LVL's
        LVLF.nfc is at offset 24 of each LVL after the LSTF array.
    Rather than recover full LVL offsets (which depend on grpprl sizes), we
    locate each LSTF's first LVLF.nfc by scanning the bytes that follow the
    LSTF array, which is good enough to classify ordered vs bullet for ilvl 0.
    """
    fib_base = 0x9A
    fc_lst = _U32.unpack_from(wd, fib_base + 47 * 8)[0]
    lcb_lst = _U32.unpack_from(wd, fib_base + 47 * 8 + 4)[0]
    fc_lfo = _U32.unpack_from(wd, fib_base + 49 * 8)[0]
    lcb_lfo = _U32.unpack_from(wd, fib_base + 49 * 8 + 4)[0]
    if lcb_lst < 2 or lcb_lfo < 4 or not table:
        return {}
    if fc_lst + lcb_lst > len(table) or fc_lfo + lcb_lfo > len(table):
        return {}

    # --- Parse PlcfLst: cLst + LSTF[cLst] + LVL blob ---
    lst_buf = table[fc_lst:fc_lst + lcb_lst]
    c_lst = _U16.unpack_from(lst_buf, 0)[0]
    lstf_size = 28
    if 2 + c_lst * lstf_size > len(lst_buf):
        return {}
    lsid_to_ordered: dict[int, bool] = {}
    lvl_blob_off = 2 + c_lst * lstf_size
    cursor = lvl_blob_off
    for i in range(c_lst):
        lsid = struct.unpack_from("<i", lst_buf, 2 + i * lstf_size)[0]
        # rgLVL count: 9 for multi-level lists, 1 for simple (bit at offset 26).
        flags = lst_buf[2 + i * lstf_size + 26] if 2 + i * lstf_size + 26 < len(lst_buf) else 0
        n_lvl = 1 if (flags & 0x10) else 9
        nfc_first: int | None = None
        for j in range(n_lvl):
            if cursor + 28 > len(lst_buf):
                break
            # LVLF (28 bytes): nfc at offset 24 (U8).
            nfc = lst_buf[cursor + 24]
            cb_grpprl_papx = lst_buf[cursor + 25]
            cb_grpprl_chpx = lst_buf[cursor + 26]
            # After LVLF: cbGrpprlPapx + cbGrpprlChpx + xst (variable).
            # xst: U16 cch + cch * U16 chars + U16 trailing reserved.
            lvl_data_off = cursor + 28 + cb_grpprl_papx + cb_grpprl_chpx
            if lvl_data_off + 2 > len(lst_buf):
                break
            cch = _U16.unpack_from(lst_buf, lvl_data_off)[0]
            lvl_end = lvl_data_off + 2 + cch * 2
            cursor = lvl_end
            if j == 0:
                nfc_first = nfc
        if nfc_first is not None and nfc_first != _NFC_NONE:
            lsid_to_ordered[lsid] = nfc_first != _NFC_BULLET

    # --- Parse PlfLfo: lfoMac + LFO[lfoMac] ---
    lfo_buf = table[fc_lfo:fc_lfo + lcb_lfo]
    lfo_mac = _U32.unpack_from(lfo_buf, 0)[0]
    lfo_size = 16
    if 4 + lfo_mac * lfo_size > len(lfo_buf):
        return {}
    ilfo_to_ordered: dict[int, bool] = {}
    for i in range(lfo_mac):
        lsid = struct.unpack_from("<i", lfo_buf, 4 + i * lfo_size)[0]
        if lsid in lsid_to_ordered:
            ilfo_to_ordered[i + 1] = lsid_to_ordered[lsid]  # ilfo is 1-based
    return ilfo_to_ordered


def _iter_sprms(grpprl: bytes):
    """Yield (opcode, value_bytes) for each sprm in a grpprl byte string.

    Operand size is encoded in bits 13-15 of the opcode (spra):
        0,1 → 1 byte; 2,4,5 → 2 bytes; 3 → 4 bytes; 7 → 3 bytes;
        6 → variable, length byte follows opcode (length includes itself).
    """
    j = 0
    n = len(grpprl)
    while j + 2 <= n:
        op = _U16.unpack_from(grpprl, j)[0]
        j += 2
        spra = (op >> 13) & 0x7
        if spra in (0, 1):
            oplen = 1
        elif spra in (2, 4, 5):
            oplen = 2
        elif spra == 3:
            oplen = 4
        elif spra == 7:
            oplen = 3
        elif spra == 6:
            if j >= n:
                break
            oplen = grpprl[j] + 1  # length byte itself counts
        else:
            break
        if j + oplen > n:
            break
        yield op, grpprl[j:j + oplen]
        j += oplen


def _read_paragraphs(
    wd: bytes, table: bytes, fc_min: int, fc_mac: int,
    ilfo_ordered: dict[int, bool] | None = None,
) -> list[Paragraph]:
    """Walk PlcfBtePapx → FKPs and return paragraphs in document order."""
    if len(wd) < _FIB_LCB_PLCFBTEPAPX + 4 or not table:
        return []
    fc_plcf = _U32.unpack_from(wd, _FIB_FC_PLCFBTEPAPX)[0]
    lcb_plcf = _U32.unpack_from(wd, _FIB_LCB_PLCFBTEPAPX)[0]
    if lcb_plcf < 12 or fc_plcf + lcb_plcf > len(table):
        return []

    plcf = table[fc_plcf:fc_plcf + lcb_plcf]
    n = (lcb_plcf - 4) // 8
    a_pn = struct.unpack_from(f"<{n}I", plcf, (n + 1) * 4)

    paragraphs: list[Paragraph] = []
    for pn in a_pn:
        fkp = wd[pn * _FKP_SIZE:(pn + 1) * _FKP_SIZE]
        if len(fkp) < _FKP_SIZE:
            continue
        cpara = fkp[_FKP_SIZE - 1]
        rgfc = struct.unpack_from(f"<{cpara + 1}I", fkp, 0)
        rgbx_off = (cpara + 1) * 4
        for i in range(cpara):
            fc_start, fc_end = rgfc[i], rgfc[i + 1]
            b_off = fkp[rgbx_off + i * 13]
            istd = 0
            in_table = False
            is_row_end = False
            ilfo = 0
            ilvl = 0
            if b_off != 0:
                papx_off = b_off * 2
                cb = fkp[papx_off]
                # PAPX layout (MS-DOC §2.9.32). When cb != 0, total length is
                # cb*2 bytes including istd+grpprl, grpprl starts at +3.
                # When cb == 0, the next byte cb' gives length cb'*2, grpprl
                # starts at +4 and istd is at +2.
                if cb != 0:
                    total = cb * 2
                    istd_pos = papx_off + 1
                    grp_start = papx_off + 3
                    grp_len = total - 3
                else:
                    cb2 = fkp[papx_off + 1] if papx_off + 1 < len(fkp) else 0
                    total = cb2 * 2
                    istd_pos = papx_off + 2
                    grp_start = papx_off + 4
                    grp_len = total - 4
                if istd_pos + 2 <= len(fkp):
                    istd = _U16.unpack_from(fkp, istd_pos)[0] & 0x0FFF
                if grp_len > 0 and grp_start + grp_len <= len(fkp):
                    for op, val in _iter_sprms(fkp[grp_start:grp_start + grp_len]):
                        if op == _SPRM_PF_IN_TABLE and val and val[0] != 0:
                            in_table = True
                        elif op == _SPRM_PT_TP and val and val[0] != 0:
                            is_row_end = True
                            in_table = True
                        elif op == _SPRM_P_ILFO and len(val) >= 2:
                            ilfo = _U16.unpack_from(val, 0)[0]
                        elif op == _SPRM_P_ILVL and val:
                            ilvl = val[0]
            s = max(fc_start, fc_min)
            e = min(fc_end, fc_mac)
            if e <= s:
                continue
            raw = wd[s:e].decode("utf-16-le", errors="replace")
            # For table cells the \x07 byte separates cells within a row; for
            # the row-end marker it's just the row terminator. Cleaning would
            # convert \x07 → tab, which we want for cell splitting, so let
            # render handle the raw text after minimal cleanup.
            text = _clean(raw)
            ordered: bool | None = None
            if ilfo and not in_table:
                if ilfo_ordered is not None and ilfo in ilfo_ordered:
                    ordered = ilfo_ordered[ilfo]
                else:
                    ordered = True  # have ilfo but no list-table info; assume ordered
            paragraphs.append(
                Paragraph(
                    istd=istd,
                    text=text,
                    in_table=in_table,
                    is_row_end=is_row_end,
                    ilfo=ilfo,
                    ilvl=ilvl,
                    list_ordered=ordered,
                )
            )
    return paragraphs


def parse(path: Union[str, Path]) -> WpsDocument:
    """Parse a .wps file and return a :class:`WpsDocument`.

    Only WPS Writer ``.wps`` files in OLE2 Word-binary form are supported.
    Raises :class:`WpsParseError` for non-.wps inputs, non-WPS binaries,
    encrypted files, or otherwise unreadable streams.
    """
    path = Path(path)
    if path.suffix.lower() != ".wps":
        raise WpsParseError(f"Only .wps files are supported (got {path.suffix!r})")
    ole = olefile.OleFileIO(str(path))
    try:
        if not ole.exists("WordDocument"):
            raise WpsParseError("No WordDocument stream — not a Word binary file")
        wd = ole.openstream("WordDocument").read()
        if len(wd) < _MIN_DOC_SIZE:
            raise WpsParseError("WordDocument stream too small")

        magic = _U16.unpack_from(wd, 0)[0]
        if magic not in (_FIB_MAGIC_WORD97, _FIB_MAGIC_WORD95):
            raise WpsParseError(f"Bad FIB magic: {hex(magic)}")

        flags = _U16.unpack_from(wd, _FIB_FLAGS_OFFSET)[0]
        if flags & _FIB_ENCRYPTED_FLAG:
            raise WpsParseError("File is encrypted / password-protected")

        # Word 97-2003 keeps the table stream in either 0Table or 1Table
        # depending on FIB flag fWhichTblStm; try 0Table first, then 1Table.
        if ole.exists("0Table"):
            table = ole.openstream("0Table").read()
        elif ole.exists("1Table"):
            table = ole.openstream("1Table").read()
        else:
            table = b""

        fc_min = _U32.unpack_from(wd, _FIB_FCMIN)[0]
        fc_mac = _U32.unpack_from(wd, _FIB_FCMAC)[0]
        ccp_text = _U32.unpack_from(wd, _FIB_CCP_TEXT)[0]
        ccp_ftn = _U32.unpack_from(wd, _FIB_CCP_FTN)[0]
        ccp_hdd = _U32.unpack_from(wd, _FIB_CCP_HDD)[0]
        ccp_atn = _U32.unpack_from(wd, _FIB_CCP_ATN)[0]

        text_bytes = fc_mac - fc_min
        total = ccp_text + ccp_ftn + ccp_hdd + ccp_atn
        if total > 0 and text_bytes == total * 2:
            mult, enc = 2, "utf-16-le"
        elif total > 0 and text_bytes == total:
            mult, enc = 1, "cp1252"
        else:
            mult, enc = 2, "utf-16-le"

        pos = fc_min
        main = wd[pos:pos + ccp_text * mult]; pos += ccp_text * mult
        ftn = wd[pos:pos + ccp_ftn * mult]; pos += ccp_ftn * mult
        hdd = wd[pos:pos + ccp_hdd * mult]; pos += ccp_hdd * mult
        atn = wd[pos:pos + ccp_atn * mult]

        ilfo_ordered = _read_list_kinds(wd, table) if mult == 2 else {}
        paragraphs = (
            _read_paragraphs(
                wd, table, fc_min, fc_min + ccp_text * mult, ilfo_ordered
            )
            if mult == 2
            else []
        )

        meta = ole.get_metadata()
        return WpsDocument(
            main_text=_clean(main.decode(enc, errors="replace")),
            paragraphs=paragraphs,
            footnotes=_clean(ftn.decode(enc, errors="replace")) if ftn else "",
            headers_footers=_clean(hdd.decode(enc, errors="replace")) if hdd else "",
            annotations=_clean(atn.decode(enc, errors="replace")) if atn else "",
            encoding=enc,
            num_pages=getattr(meta, "num_pages", None),
        )
    finally:
        ole.close()


def _escape_cell(text: str) -> str:
    """Escape characters that would break a Markdown table cell."""
    return text.replace("\\", "\\\\").replace("|", "\\|").replace("\n", " ").strip()


def _flush_table(rows: list[list[str]], out: list[str]) -> None:
    """Emit a Markdown pipe table for collected rows.

    The whole table is appended as a single block (lines joined by ``\n``)
    so that the outer ``\n\n`` join keeps blank lines only between blocks,
    not between rows of the same table.
    """
    if not rows:
        return
    width = max(len(r) for r in rows)
    norm = [r + [""] * (width - len(r)) for r in rows]
    header = norm[0]
    body = norm[1:] if len(norm) > 1 else []
    lines = [
        "| " + " | ".join(_escape_cell(c) for c in header) + " |",
        "|" + "|".join([" --- "] * width) + "|",
    ]
    for row in body:
        lines.append("| " + " | ".join(_escape_cell(c) for c in row) + " |")
    out.append("\n".join(lines))
    rows.clear()


def to_markdown(paragraphs: list[Paragraph]) -> str:
    """Render paragraphs to Markdown.

    Heading 1-9 (istd 1-9) become ``#``..``#########``. Paragraphs flagged
    as in-table are accumulated into a Markdown pipe table, with row
    boundaries from ``is_row_end``. All other paragraphs are emitted as
    plain text. Output ends with a single trailing newline.
    """
    out: list[str] = []
    table_rows: list[list[str]] = []
    current_row: list[str] = []
    # Counters keyed by (ilfo, ilvl) for ordered list numbering.
    list_counters: dict[tuple[int, int], int] = {}
    for p in paragraphs:
        if p.in_table:
            if p.is_row_end:
                # Row terminator paragraph; commit the row we've accumulated.
                if current_row:
                    table_rows.append(current_row)
                    current_row = []
            else:
                current_row.append(p.text)
            continue
        # Leaving a table region: flush whatever we collected.
        if current_row:
            table_rows.append(current_row)
            current_row = []
        if table_rows:
            _flush_table(table_rows, out)
        if not p.text:
            continue
        if p.heading_level:
            out.append(f"{'#' * p.heading_level} {p.text}")
            list_counters.clear()
        elif p.list_ordered is not None:
            indent = "  " * max(0, p.ilvl)
            key = (p.ilfo, p.ilvl)
            if p.list_ordered:
                list_counters[key] = list_counters.get(key, 0) + 1
                marker = f"{list_counters[key]}."
            else:
                marker = "-"
            out.append(f"{indent}{marker} {p.text}")
        else:
            out.append(p.text)
            list_counters.clear()
    # Flush any trailing table.
    if current_row:
        table_rows.append(current_row)
    if table_rows:
        _flush_table(table_rows, out)
    return "\n\n".join(out) + "\n" if out else ""
