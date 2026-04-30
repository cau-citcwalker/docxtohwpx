#!/usr/bin/env python3
"""
DOCX → HWPX converter (template-based).

Usage:
    python docx2hwpx.py input.docx output.hwpx

Approach:
  - The static skeleton of an HWPX (mimetype, version.xml, settings.xml,
    META-INF/*, Preview/*, content.hpf, header.xml, section0.xml) is shipped
    alongside this script as `skeleton.hwpx`. That skeleton is a minimal
    well-formed HWPX file that Hangul opens cleanly.
  - On conversion we keep the skeleton's static parts (and its preset
    charPr/paraPr/style/font tables) verbatim and *append* any new
    character/paragraph/font definitions our DOCX needs onto the existing
    tables, then rewrite section0.xml to reference them.
  - Binary assets (images) are added to BinData/ and registered in the OPF
    manifest (content.hpf).

Why this matters:
  HWPX validates strictly. A handcrafted file diverges from the OWPML schema
  in dozens of small ways (flag elements that should be absent vs. present,
  required attributes, namespace prefixes, etc.). Reusing a real skeleton
  lets us focus on emitting the *content* portion correctly.

Units:
  DOCX twip   = 1/1440 inch        →  HWPUNIT (1/7200 inch) = twip * 5
  DOCX EMU    = 1/914400 inch      →  HWPUNIT                = EMU // 127
  DOCX half-pt = 1/200 inch (size) →  HWPX size (1/100 pt)   = halfpt * 50

Supported:
  text, paragraphs (alignment, list/heading, indent), runs (bold/italic/
  underline/strike, size, color, font), tables (col + row span via vMerge,
  cell widths, nested), images (PNG/JPG/GIF/BMP), hyperlinks, footnotes,
  endnotes, accept-all-track-changes.

Not supported:
  Office Math equations (MathML/OMML to OWPML equation language is a
  separate transpiler), SmartArt, complex VML legacy drawings,
  multi-section headers/footers, comments.
"""
from __future__ import annotations

import argparse
import os
import posixpath
import re
import sys
import zipfile
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple, Union
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape as xml_escape

# ---------------------------------------------------------------------------
# DOCX namespaces
# ---------------------------------------------------------------------------
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"

REL_HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
REL_HEADER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
REL_FOOTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
REL_FOOTNOTES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
REL_ENDNOTES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"


def _w(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


def _r(tag: str) -> str:
    return f"{{{R_NS}}}{tag}"


def _local(tag: str) -> str:
    return tag.split("}", 1)[-1]


# ---------------------------------------------------------------------------
# Unit helpers
# ---------------------------------------------------------------------------
def twip_to_hwp(t: int) -> int:
    return int(t) * 5


def halfpt_to_hwp_size(hp: int) -> int:
    return int(hp) * 50


def emu_to_hwp(emu: int) -> int:
    return int(emu) // 127


# ---------------------------------------------------------------------------
# Intermediate representation
# ---------------------------------------------------------------------------
@dataclass(frozen=True)
class CharProps:
    bold: bool = False
    italic: bool = False
    underline: bool = False
    strike: bool = False
    size_halfpt: int = 20
    color: str = "#000000"
    font_name: str = "함초롬바탕"


@dataclass(frozen=True)
class ParaProps:
    align: str = "JUSTIFY"
    style_name: str = "Normal"
    list_level: Optional[int] = None
    list_type: Optional[str] = None


@dataclass
class Run:
    text: str
    char_props: CharProps = field(default_factory=CharProps)


@dataclass
class Hyperlink:
    url: str
    runs: List[Run] = field(default_factory=list)


@dataclass
class Image:
    bin_id: str
    width_hwp: int
    height_hwp: int
    char_props: CharProps = field(default_factory=CharProps)


@dataclass
class FootnoteRef:
    paragraphs: List["Paragraph"] = field(default_factory=list)
    char_props: CharProps = field(default_factory=CharProps)
    is_endnote: bool = False


@dataclass
class PageBreak:
    """An explicit page break in the middle of (or between) paragraphs.

    The writer treats this as a paragraph boundary: any inlines after a
    PageBreak are emitted in a new <hp:p pageBreak="1"> element.
    """
    pass


Inline = Union[Run, Hyperlink, Image, FootnoteRef, PageBreak]


@dataclass
class Paragraph:
    inlines: List[Inline] = field(default_factory=list)
    para_props: ParaProps = field(default_factory=ParaProps)


@dataclass
class TableCell:
    blocks: List[Union["Paragraph", "Table"]] = field(default_factory=list)
    grid_span: int = 1
    row_span: int = 1
    v_merge_continue: bool = False
    fill_color: Optional[str] = None  # "#RRGGBB" or None for no fill


@dataclass
class TableRow:
    cells: List[TableCell] = field(default_factory=list)
    height_hwp: int = 0


@dataclass
class Table:
    rows: List[TableRow] = field(default_factory=list)
    col_widths_hwp: List[int] = field(default_factory=list)


@dataclass
class HeaderFooter:
    paragraphs: List[Paragraph] = field(default_factory=list)
    apply_to: str = "BOTH"


@dataclass
class BinaryItem:
    bin_id: str
    href: str
    media_type: str
    fmt: str
    data: bytes


@dataclass
class Document:
    blocks: List[Union[Paragraph, Table]] = field(default_factory=list)
    binaries: List[BinaryItem] = field(default_factory=list)
    header: Optional[HeaderFooter] = None
    footer: Optional[HeaderFooter] = None
    page_width_hwp: int = 59528
    page_height_hwp: int = 84188
    margin_left_hwp: int = 8504
    margin_right_hwp: int = 8504
    margin_top_hwp: int = 5668
    margin_bottom_hwp: int = 4252
    margin_header_hwp: int = 4252
    margin_footer_hwp: int = 4252
    landscape: bool = False


# ---------------------------------------------------------------------------
# DOCX reader
# ---------------------------------------------------------------------------
_IMG_FMT_BY_EXT = {
    ".png": ("PNG", "image/png"),
    ".jpg": ("JPG", "image/jpeg"),
    ".jpeg": ("JPG", "image/jpeg"),
    ".gif": ("GIF", "image/gif"),
    ".bmp": ("BMP", "image/bmp"),
    ".wmf": ("WMF", "image/x-wmf"),
    ".emf": ("EMF", "image/x-emf"),
    ".tif": ("TIFF", "image/tiff"),
    ".tiff": ("TIFF", "image/tiff"),
}


class DocxReader:
    def __init__(self, path: str):
        self.path = path
        self.zip = zipfile.ZipFile(path, "r")
        self._part_rels: Dict[str, Dict[str, dict]] = {}
        self.styles = self._load_styles()
        self.numbering = self._load_numbering()
        self.doc_part = "word/document.xml"
        self._load_rels(self.doc_part)
        self._footnotes: Dict[str, List[Paragraph]] = {}
        self._endnotes: Dict[str, List[Paragraph]] = {}
        self._binaries: List[BinaryItem] = []
        self._bin_seen_targets: Dict[str, str] = {}

    def _rels_path_for(self, part_path: str) -> str:
        d, n = posixpath.split(part_path)
        return posixpath.join(d, "_rels", n + ".rels") if d else "_rels/" + n + ".rels"

    def _load_rels(self, part_path: str) -> Dict[str, dict]:
        if part_path in self._part_rels:
            return self._part_rels[part_path]
        rels_path = self._rels_path_for(part_path)
        rels: Dict[str, dict] = {}
        try:
            data = self.zip.read(rels_path)
        except KeyError:
            self._part_rels[part_path] = rels
            return rels
        try:
            root = ET.fromstring(data)
        except ET.ParseError:
            self._part_rels[part_path] = rels
            return rels
        for r in root.findall(f"{{{PKG_REL_NS}}}Relationship"):
            rid = r.get("Id")
            if rid:
                rels[rid] = {
                    "type": r.get("Type", ""),
                    "target": r.get("Target", ""),
                    "mode": r.get("TargetMode", "Internal"),
                }
        self._part_rels[part_path] = rels
        return rels

    def _resolve_target(self, part_path: str, target: str) -> str:
        if "://" in target or target.startswith("/"):
            return target
        d = posixpath.dirname(part_path)
        return posixpath.normpath(posixpath.join(d, target))

    def _load_styles(self) -> Dict[str, dict]:
        try:
            data = self.zip.read("word/styles.xml")
        except KeyError:
            return {}
        try:
            root = ET.fromstring(data)
        except ET.ParseError:
            return {}
        out: Dict[str, dict] = {}
        for st in root.findall(_w("style")):
            sid = st.get(_w("styleId"))
            if not sid:
                continue
            entry = {"id": sid}
            name_elt = st.find(_w("name"))
            entry["name"] = name_elt.get(_w("val"), sid) if name_elt is not None else sid
            entry["pPr"] = st.find(_w("pPr"))
            entry["rPr"] = st.find(_w("rPr"))
            out[sid] = entry
        return out

    def _load_numbering(self) -> dict:
        try:
            data = self.zip.read("word/numbering.xml")
        except KeyError:
            return {"abs": {}, "num": {}}
        try:
            root = ET.fromstring(data)
        except ET.ParseError:
            return {"abs": {}, "num": {}}
        abs_nums: Dict[str, Dict[str, str]] = {}
        for an in root.findall(_w("abstractNum")):
            aid = an.get(_w("abstractNumId"))
            if aid is None:
                continue
            levels: Dict[str, str] = {}
            for lvl in an.findall(_w("lvl")):
                ilvl = lvl.get(_w("ilvl"))
                fmt = lvl.find(_w("numFmt"))
                fmt_val = fmt.get(_w("val")) if fmt is not None else "decimal"
                if ilvl is not None and fmt_val is not None:
                    levels[ilvl] = fmt_val
            abs_nums[aid] = levels
        nums: Dict[str, str] = {}
        for n in root.findall(_w("num")):
            nid = n.get(_w("numId"))
            ai = n.find(_w("abstractNumId"))
            if nid is not None and ai is not None:
                v = ai.get(_w("val"))
                if v is not None:
                    nums[nid] = v
        return {"abs": abs_nums, "num": nums}

    def _ensure_notes_loaded(self):
        rels = self._part_rels[self.doc_part]
        for rid, info in rels.items():
            t = info["type"]
            if t == REL_FOOTNOTES:
                self._footnotes = self._parse_notes_part(
                    self._resolve_target(self.doc_part, info["target"]), "footnote"
                )
            elif t == REL_ENDNOTES:
                self._endnotes = self._parse_notes_part(
                    self._resolve_target(self.doc_part, info["target"]), "endnote"
                )

    def _parse_notes_part(self, part_path: str, kind: str) -> Dict[str, List[Paragraph]]:
        try:
            data = self.zip.read(part_path)
        except KeyError:
            return {}
        try:
            root = ET.fromstring(data)
        except ET.ParseError:
            return {}
        self._load_rels(part_path)
        out: Dict[str, List[Paragraph]] = {}
        elt_name = _w("footnote") if kind == "footnote" else _w("endnote")
        for note in root.findall(elt_name):
            nid = note.get(_w("id"))
            note_type = note.get(_w("type"))
            if note_type in ("separator", "continuationSeparator"):
                continue
            if nid is None:
                continue
            paras: List[Paragraph] = []
            for p in note.findall(_w("p")):
                paras.append(self._parse_paragraph(p, part_path))
            out[nid] = paras
        return out

    def parse(self) -> Document:
        data = self.zip.read(self.doc_part)
        root = ET.fromstring(data)
        body = root.find(_w("body"))
        doc = Document()
        if body is None:
            return doc
        self._ensure_notes_loaded()
        self._load_first_header_footer(doc)

        for child in body:
            tag = _local(child.tag)
            if tag == "p":
                doc.blocks.append(self._parse_paragraph(child, self.doc_part))
            elif tag == "tbl":
                doc.blocks.append(self._parse_table(child, self.doc_part))
            elif tag == "sectPr":
                self._apply_sect_props(child, doc)

        for p in body.findall(_w("p")):
            ppr = p.find(_w("pPr"))
            if ppr is not None:
                spr = ppr.find(_w("sectPr"))
                if spr is not None:
                    self._apply_sect_props(spr, doc)

        doc.binaries = self._binaries
        return doc

    def _apply_sect_props(self, spr: ET.Element, doc: Document):
        pgsz = spr.find(_w("pgSz"))
        if pgsz is not None:
            w = pgsz.get(_w("w"))
            h = pgsz.get(_w("h"))
            if w:
                doc.page_width_hwp = twip_to_hwp(int(w))
            if h:
                doc.page_height_hwp = twip_to_hwp(int(h))
            doc.landscape = (pgsz.get(_w("orient")) == "landscape")
        pgmar = spr.find(_w("pgMar"))
        if pgmar is not None:
            for src, attr in (
                ("left", "margin_left_hwp"),
                ("right", "margin_right_hwp"),
                ("top", "margin_top_hwp"),
                ("bottom", "margin_bottom_hwp"),
                ("header", "margin_header_hwp"),
                ("footer", "margin_footer_hwp"),
            ):
                v = pgmar.get(_w(src))
                if v:
                    try:
                        setattr(doc, attr, twip_to_hwp(int(v)))
                    except ValueError:
                        pass

    def _load_first_header_footer(self, doc: Document):
        try:
            data = self.zip.read(self.doc_part)
            root = ET.fromstring(data)
        except (KeyError, ET.ParseError):
            return
        sect_prs: List[ET.Element] = []
        for spr in root.iter(_w("sectPr")):
            sect_prs.append(spr)
        rels = self._part_rels[self.doc_part]
        chosen_header_rid: Optional[str] = None
        chosen_footer_rid: Optional[str] = None
        for spr in sect_prs:
            for hr in spr.findall(_w("headerReference")):
                rid = hr.get(_r("id"))
                t = hr.get(_w("type"), "default")
                if rid and chosen_header_rid is None and t in ("default", "first", "even"):
                    chosen_header_rid = rid
                    if t == "default":
                        break
            for fr in spr.findall(_w("footerReference")):
                rid = fr.get(_r("id"))
                t = fr.get(_w("type"), "default")
                if rid and chosen_footer_rid is None and t in ("default", "first", "even"):
                    chosen_footer_rid = rid
                    if t == "default":
                        break
        if chosen_header_rid and chosen_header_rid in rels:
            target = self._resolve_target(self.doc_part, rels[chosen_header_rid]["target"])
            doc.header = self._load_hf_part(target)
        if chosen_footer_rid and chosen_footer_rid in rels:
            target = self._resolve_target(self.doc_part, rels[chosen_footer_rid]["target"])
            doc.footer = self._load_hf_part(target)

    def _load_hf_part(self, part_path: str) -> Optional[HeaderFooter]:
        try:
            data = self.zip.read(part_path)
        except KeyError:
            return None
        try:
            root = ET.fromstring(data)
        except ET.ParseError:
            return None
        self._load_rels(part_path)
        hf = HeaderFooter()
        for child in root:
            tag = _local(child.tag)
            if tag == "p":
                hf.paragraphs.append(self._parse_paragraph(child, part_path))
            elif tag == "tbl":
                t = self._parse_table(child, part_path)
                for row in t.rows:
                    for cell in row.cells:
                        for cb in cell.blocks:
                            if isinstance(cb, Paragraph):
                                hf.paragraphs.append(cb)
        return hf

    def _parse_paragraph(self, p: ET.Element, part_path: str) -> Paragraph:
        ppr = p.find(_w("pPr"))
        para_props = self._read_para_props(ppr)
        style_rpr = None
        page_break_before = False
        if ppr is not None:
            pstyle = ppr.find(_w("pStyle"))
            if pstyle is not None:
                sid = pstyle.get(_w("val"))
                if sid in self.styles:
                    style_rpr = self.styles[sid].get("rPr")
            if ppr.find(_w("pageBreakBefore")) is not None:
                page_break_before = True
        para = Paragraph(para_props=para_props)
        if page_break_before:
            para.inlines.append(PageBreak())
        base_cp = self._read_char_props(style_rpr, CharProps())
        self._walk_inlines(p, para, base_cp, part_path)
        return para

    def _walk_inlines(self, parent: ET.Element, para: Paragraph,
                      base_cp: CharProps, part_path: str):
        for child in parent:
            tag = _local(child.tag)
            if tag == "pPr":
                continue
            elif tag == "r":
                self._parse_run(child, para, base_cp, part_path)
            elif tag == "hyperlink":
                self._parse_hyperlink(child, para, base_cp, part_path)
            elif tag == "ins":
                self._walk_inlines(child, para, base_cp, part_path)
            elif tag == "del":
                continue
            elif tag in ("smartTag", "fldSimple", "moveTo",
                         "customXmlInsRangeStart", "customXmlInsRangeEnd",
                         "customXmlDelRangeStart", "customXmlDelRangeEnd",
                         "bookmarkStart", "bookmarkEnd",
                         "proofErr", "permStart", "permEnd",
                         "moveFromRangeStart", "moveFromRangeEnd",
                         "moveToRangeStart", "moveToRangeEnd"):
                self._walk_inlines(child, para, base_cp, part_path)

    def _parse_hyperlink(self, h: ET.Element, para: Paragraph,
                         base_cp: CharProps, part_path: str):
        rid = h.get(_r("id"))
        anchor = h.get(_w("anchor"))
        url = ""
        if rid:
            rels = self._part_rels.get(part_path, {})
            target = rels.get(rid, {}).get("target", "")
            if target:
                url = target
        elif anchor:
            url = "#" + anchor
        tmp = Paragraph()
        self._walk_inlines(h, tmp, base_cp, part_path)
        flat_runs: List[Run] = []
        leftover_images: List[Image] = []
        for inl in tmp.inlines:
            if isinstance(inl, Run):
                flat_runs.append(inl)
            elif isinstance(inl, Hyperlink):
                flat_runs.extend(inl.runs)
            elif isinstance(inl, Image):
                leftover_images.append(inl)
        if url:
            para.inlines.append(Hyperlink(url=url, runs=flat_runs))
        else:
            for r in flat_runs:
                para.inlines.append(r)
        for img in leftover_images:
            para.inlines.append(img)

    def _read_para_props(self, ppr: Optional[ET.Element]) -> ParaProps:
        if ppr is None:
            return ParaProps()
        align = "JUSTIFY"
        style_name = "Normal"
        list_level: Optional[int] = None
        list_type: Optional[str] = None

        pstyle = ppr.find(_w("pStyle"))
        if pstyle is not None:
            sid = pstyle.get(_w("val"))
            if sid in self.styles:
                style_name = self.styles[sid].get("name", sid)
            elif sid:
                style_name = sid

        jc = ppr.find(_w("jc"))
        if jc is not None:
            align = {
                "left": "LEFT", "start": "LEFT",
                "right": "RIGHT", "end": "RIGHT",
                "center": "CENTER",
                "both": "JUSTIFY", "justify": "JUSTIFY",
                "distribute": "DISTRIBUTE",
            }.get(jc.get(_w("val")) or "", "JUSTIFY")

        numpr = ppr.find(_w("numPr"))
        if numpr is not None:
            ilvl = numpr.find(_w("ilvl"))
            num_id = numpr.find(_w("numId"))
            if ilvl is not None:
                try:
                    list_level = int(ilvl.get(_w("val"), "0"))
                except ValueError:
                    list_level = 0
            else:
                list_level = 0
            if num_id is not None:
                nid = num_id.get(_w("val"))
                if nid and nid != "0":
                    aid = self.numbering["num"].get(nid)
                    if aid is not None:
                        levels = self.numbering["abs"].get(aid, {})
                        fmt = levels.get(str(list_level or 0), "decimal")
                        list_type = "bullet" if fmt == "bullet" else "decimal"
                else:
                    list_level = None
        return ParaProps(
            align=align, style_name=style_name,
            list_level=list_level, list_type=list_type,
        )

    def _parse_run(self, r: ET.Element, para: Paragraph, base: CharProps,
                   part_path: str):
        rpr = r.find(_w("rPr"))
        cp = self._read_char_props(rpr, base)
        text_buf: List[str] = []

        def flush_text():
            t = "".join(text_buf)
            if t:
                para.inlines.append(Run(text=t, char_props=cp))
            text_buf.clear()

        for child in r:
            tag = _local(child.tag)
            if tag == "rPr":
                continue
            elif tag == "t":
                text_buf.append(child.text or "")
            elif tag == "tab":
                text_buf.append("\t")
            elif tag == "br":
                br_type = child.get(_w("type"))
                if br_type == "page":
                    flush_text()
                    para.inlines.append(PageBreak())
                else:
                    text_buf.append("\n")
            elif tag == "noBreakHyphen":
                text_buf.append("‑")
            elif tag == "softHyphen":
                text_buf.append("­")
            elif tag == "sym":
                ch = child.get(_w("char"))
                if ch:
                    try:
                        text_buf.append(chr(int(ch, 16)))
                    except ValueError:
                        pass
            elif tag == "delText":
                text_buf.append(child.text or "")
            elif tag == "drawing":
                flush_text()
                img = self._parse_drawing(child, part_path, cp)
                if img is not None:
                    para.inlines.append(img)
            elif tag in ("footnoteReference", "endnoteReference"):
                flush_text()
                nid = child.get(_w("id"))
                if nid is not None:
                    if tag == "footnoteReference":
                        notes = self._footnotes.get(nid, [])
                        para.inlines.append(FootnoteRef(paragraphs=notes,
                                                        char_props=cp,
                                                        is_endnote=False))
                    else:
                        notes = self._endnotes.get(nid, [])
                        para.inlines.append(FootnoteRef(paragraphs=notes,
                                                        char_props=cp,
                                                        is_endnote=True))
            elif tag == "object":
                for t_elt in child.iter(_w("t")):
                    text_buf.append(t_elt.text or "")
            elif _local(child.tag) in ("oMath", "oMathPara"):
                for t_elt in child.iter():
                    if _local(t_elt.tag) == "t":
                        text_buf.append(t_elt.text or "")
        flush_text()

    def _read_char_props(self, rpr: Optional[ET.Element], base: CharProps) -> CharProps:
        if rpr is None:
            return base
        bold = base.bold or (rpr.find(_w("b")) is not None and not _is_off(rpr.find(_w("b"))))
        italic = base.italic or (rpr.find(_w("i")) is not None and not _is_off(rpr.find(_w("i"))))
        u = rpr.find(_w("u"))
        underline = base.underline or (u is not None and u.get(_w("val")) not in (None, "none"))
        strike = base.strike or (rpr.find(_w("strike")) is not None and not _is_off(rpr.find(_w("strike"))))

        size_halfpt = base.size_halfpt
        sz = rpr.find(_w("sz"))
        if sz is not None:
            try:
                size_halfpt = int(sz.get(_w("val"), str(base.size_halfpt)))
            except ValueError:
                pass

        color = base.color
        c = rpr.find(_w("color"))
        if c is not None:
            v = c.get(_w("val"))
            if v and v.lower() != "auto" and re.match(r"^[0-9A-Fa-f]{6}$", v):
                color = "#" + v.upper()

        font_name = base.font_name
        rfonts = rpr.find(_w("rFonts"))
        if rfonts is not None:
            for attr in ("eastAsia", "ascii", "hAnsi", "cs"):
                v = rfonts.get(_w(attr))
                if v:
                    font_name = v
                    break

        return CharProps(
            bold=bool(bold), italic=bool(italic),
            underline=bool(underline), strike=bool(strike),
            size_halfpt=size_halfpt, color=color, font_name=font_name,
        )

    def _parse_drawing(self, drawing: ET.Element, part_path: str,
                       cp: CharProps) -> Optional[Image]:
        cx = cy = 0
        ext = drawing.find(f".//{{{WP_NS}}}extent")
        if ext is not None:
            try:
                cx = int(ext.get("cx", "0"))
                cy = int(ext.get("cy", "0"))
            except ValueError:
                pass
        blip = drawing.find(f".//{{{A_NS}}}blip")
        if blip is None:
            return None
        embed = blip.get(_r("embed"))
        if not embed:
            return None
        rels = self._part_rels.get(part_path, {})
        info = rels.get(embed)
        if not info or info.get("type") != REL_IMAGE:
            return None
        target = self._resolve_target(part_path, info["target"])
        bin_id = self._intern_binary(target)
        if bin_id is None:
            return None
        return Image(
            bin_id=bin_id,
            width_hwp=emu_to_hwp(cx) or 14000,
            height_hwp=emu_to_hwp(cy) or 14000,
            char_props=cp,
        )

    def _intern_binary(self, target_path: str) -> Optional[str]:
        if target_path in self._bin_seen_targets:
            return self._bin_seen_targets[target_path]
        try:
            data = self.zip.read(target_path)
        except KeyError:
            return None
        ext = posixpath.splitext(target_path)[1].lower()
        fmt, media_type = _IMG_FMT_BY_EXT.get(ext, ("PNG", "image/png"))
        bin_id = f"image{len(self._binaries) + 1}"
        href = f"BinData/{bin_id}{ext if ext else '.' + fmt.lower()}"
        self._binaries.append(BinaryItem(
            bin_id=bin_id, href=href, media_type=media_type, fmt=fmt, data=data,
        ))
        self._bin_seen_targets[target_path] = bin_id
        return bin_id

    def _parse_table(self, tbl: ET.Element, part_path: str) -> Table:
        table = Table()
        grid = tbl.find(_w("tblGrid"))
        if grid is not None:
            for gc in grid.findall(_w("gridCol")):
                w = gc.get(_w("w"))
                if w:
                    try:
                        table.col_widths_hwp.append(twip_to_hwp(int(w)))
                    except ValueError:
                        table.col_widths_hwp.append(0)

        raw_rows: List[List[Tuple[TableCell, Optional[str]]]] = []
        for tr in tbl.findall(_w("tr")):
            row = TableRow()
            trpr = tr.find(_w("trPr"))
            if trpr is not None:
                trh = trpr.find(_w("trHeight"))
                if trh is not None:
                    v = trh.get(_w("val"))
                    if v:
                        try:
                            row.height_hwp = twip_to_hwp(int(v))
                        except ValueError:
                            pass
            row_raw: List[Tuple[TableCell, Optional[str]]] = []
            for tc in tr.findall(_w("tc")):
                cell = TableCell()
                vmerge_marker: Optional[str] = None
                tcpr = tc.find(_w("tcPr"))
                if tcpr is not None:
                    gs = tcpr.find(_w("gridSpan"))
                    if gs is not None:
                        try:
                            cell.grid_span = int(gs.get(_w("val"), "1"))
                        except ValueError:
                            pass
                    vm = tcpr.find(_w("vMerge"))
                    if vm is not None:
                        v = vm.get(_w("val"), "continue")
                        vmerge_marker = "restart" if v == "restart" else "continue"
                    shd = tcpr.find(_w("shd"))
                    if shd is not None:
                        fill = shd.get(_w("fill"))
                        if fill and fill.lower() != "auto" and re.match(r"^[0-9A-Fa-f]{6}$", fill):
                            cell.fill_color = "#" + fill.upper()
                for child in tc:
                    tag = _local(child.tag)
                    if tag == "p":
                        cell.blocks.append(self._parse_paragraph(child, part_path))
                    elif tag == "tbl":
                        cell.blocks.append(self._parse_table(child, part_path))
                if not cell.blocks:
                    cell.blocks.append(Paragraph())
                row_raw.append((cell, vmerge_marker))
            row.cells = [c for c, _ in row_raw]
            table.rows.append(row)
            raw_rows.append(row_raw)

        active_starts: Dict[int, TableCell] = {}
        for row_raw in raw_rows:
            grid_col = 0
            for cell, vm in row_raw:
                if vm == "restart":
                    active_starts[grid_col] = cell
                elif vm == "continue":
                    origin = active_starts.get(grid_col)
                    if origin is not None:
                        origin.row_span += 1
                        cell.v_merge_continue = True
                else:
                    active_starts.pop(grid_col, None)
                grid_col += cell.grid_span
            seen_cols = set()
            grid_col = 0
            for cell, vm in row_raw:
                if vm in ("restart", "continue"):
                    seen_cols.add(grid_col)
                grid_col += cell.grid_span
            for col in list(active_starts.keys()):
                if col not in seen_cols:
                    active_starts.pop(col, None)

        return table

    def close(self):
        self.zip.close()


def _is_off(elt: Optional[ET.Element]) -> bool:
    if elt is None:
        return False
    v = elt.get(_w("val"))
    return v in ("0", "false")


# ---------------------------------------------------------------------------
# HWPX writer (template-based)
# ---------------------------------------------------------------------------
SKELETON_FILENAME = "skeleton.hwpx"

# All namespaces used in skeleton's section0.xml — copied so generated
# section XML has the same surrounding declarations.
_SECTION_NS_DECL = (
    'xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" '
    'xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" '
    'xmlns:hp10="http://www.hancom.co.kr/hwpml/2016/paragraph" '
    'xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" '
    'xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core" '
    'xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" '
    'xmlns:hhs="http://www.hancom.co.kr/hwpml/2011/history" '
    'xmlns:hm="http://www.hancom.co.kr/hwpml/2011/master-page" '
    'xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf" '
    'xmlns:dc="http://purl.org/dc/elements/1.1/" '
    'xmlns:opf="http://www.idpf.org/2007/opf/" '
    'xmlns:ooxmlchart="http://www.hancom.co.kr/hwpml/2016/ooxmlchart" '
    'xmlns:hwpunitchar="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar" '
    'xmlns:epub="http://www.idpf.org/2007/ops" '
    'xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"'
)


class HwpxWriter:
    """Generates HWPX by patching a known-good skeleton file.

    The skeleton already contains valid header.xml/charProperties etc. We
    *append* our custom CharProps/ParaProps/Fontfaces to the existing tables
    and use the new ids in our section0.xml.
    """

    def __init__(self, doc: Document, skeleton_path: str):
        self.doc = doc
        self.skeleton_path = skeleton_path
        # Read all skeleton entries into memory.
        with zipfile.ZipFile(skeleton_path, "r") as z:
            self.skeleton: Dict[str, bytes] = {n: z.read(n) for n in z.namelist()}

        # Parse skeleton header: find current itemCnts and inject points.
        self._header_str = self.skeleton["Contents/header.xml"].decode("utf-8")
        # Skeleton font ids that we know about: 0 = 함초롬돋움, 1 = 함초롬바탕
        self._skeleton_font_count = self._count_attr(
            r'<hh:fontface lang="HANGUL" fontCnt="(\d+)"', self._header_str
        )
        self._skeleton_charpr_count = self._count_attr(
            r'<hh:charProperties itemCnt="(\d+)"', self._header_str
        )
        self._skeleton_parapr_count = self._count_attr(
            r'<hh:paraProperties itemCnt="(\d+)"', self._header_str
        )
        self._skeleton_borderfill_count = self._count_attr(
            r'<hh:borderFills itemCnt="(\d+)"', self._header_str
        )
        # Cell borderFills: one per unique fill color seen across all tables.
        # The None key is the default "no fill" SOLID-bordered borderFill.
        # IDs are skeleton_count+1, +2, +3, ... in dict insertion order.
        self._cell_borderfills: Dict[Optional[str], int] = {
            None: self._skeleton_borderfill_count + 1
        }
        # The table itself (frame) uses the no-fill SOLID border.
        self._table_border_id = self._cell_borderfills[None]

        # Reverse lookups for known fonts.
        self._font_idx: Dict[str, int] = {
            "함초롬돋움": 0,
            "함초롬바탕": 1,
        }

        # Defaults: empty CharProps maps to skeleton charPr 0; empty ParaProps
        # to skeleton paraPr 0 (바탕글). These ids are baked into the skeleton.
        self._cp_idx: Dict[CharProps, int] = {CharProps(): 0}
        self._pp_idx: Dict[ParaProps, int] = {ParaProps(): 0}

        # New entries we'll append.
        self._new_fonts: List[str] = []
        self._new_cps: List[Tuple[int, CharProps, int]] = []  # (id, cp, font_id)
        self._new_pps: List[Tuple[int, ParaProps]] = []

        self._para_id_counter = 0
        self._ctrl_id_counter = 0
        self._footnote_counter = 0
        self._endnote_counter = 0

        self._collect_refs()

    @staticmethod
    def _count_attr(pattern: str, s: str) -> int:
        m = re.search(pattern, s)
        return int(m.group(1)) if m else 0

    # -- registration helpers ---------------------------------------------
    def _intern_font(self, name: str) -> int:
        if name in self._font_idx:
            return self._font_idx[name]
        new_id = self._skeleton_font_count + len(self._new_fonts)
        self._font_idx[name] = new_id
        self._new_fonts.append(name)
        return new_id

    def _intern_cp(self, cp: CharProps) -> int:
        if cp in self._cp_idx:
            return self._cp_idx[cp]
        font_id = self._intern_font(cp.font_name)
        new_id = self._skeleton_charpr_count + len(self._new_cps)
        self._cp_idx[cp] = new_id
        self._new_cps.append((new_id, cp, font_id))
        return new_id

    def _intern_pp(self, pp: ParaProps) -> int:
        if pp in self._pp_idx:
            return self._pp_idx[pp]
        new_id = self._skeleton_parapr_count + len(self._new_pps)
        self._pp_idx[pp] = new_id
        self._new_pps.append((new_id, pp))
        return new_id

    def _intern_cell_fill(self, color: Optional[str]) -> int:
        if color in self._cell_borderfills:
            return self._cell_borderfills[color]
        new_id = self._skeleton_borderfill_count + 1 + len(self._cell_borderfills)
        self._cell_borderfills[color] = new_id
        return new_id

    def _collect_refs(self):
        def visit_para(p: Paragraph):
            self._intern_pp(p.para_props)
            for inl in p.inlines:
                if isinstance(inl, Run):
                    self._intern_cp(inl.char_props)
                elif isinstance(inl, Hyperlink):
                    for r in inl.runs:
                        self._intern_cp(r.char_props)
                elif isinstance(inl, (Image, FootnoteRef)):
                    self._intern_cp(inl.char_props)
                    if isinstance(inl, FootnoteRef):
                        for fp in inl.paragraphs:
                            visit_para(fp)

        def visit_block(b):
            if isinstance(b, Paragraph):
                visit_para(b)
            elif isinstance(b, Table):
                for row in b.rows:
                    for cell in row.cells:
                        if cell.v_merge_continue:
                            continue
                        if cell.fill_color is not None:
                            self._intern_cell_fill(cell.fill_color)
                        for cb in cell.blocks:
                            visit_block(cb)

        for b in self.doc.blocks:
            visit_block(b)
        if self.doc.header:
            for p in self.doc.header.paragraphs:
                visit_para(p)
        if self.doc.footer:
            for p in self.doc.footer.paragraphs:
                visit_para(p)

    def _next_para_id(self) -> int:
        self._para_id_counter += 1
        return self._para_id_counter

    def _next_ctrl_id(self) -> int:
        self._ctrl_id_counter += 1
        return self._ctrl_id_counter

    # -- header.xml patching ----------------------------------------------
    def _patch_header(self) -> str:
        out = self._header_str
        out = self._patch_borderfills(out)
        if self._new_fonts:
            out = self._patch_fontfaces(out)
        if self._new_cps:
            out = self._patch_charpr(out)
        if self._new_pps:
            out = self._patch_parapr(out)
        return out

    def _patch_borderfills(self, header: str) -> str:
        """Append SOLID-bordered borderFills for tables/cells.

        We emit one borderFill per unique cell fill color (plus the default
        no-fill SOLID border at id N+1). Skeleton borderFills 1..N are all
        NONE-bordered, so any table cell using them would be invisible.
        """
        new_total = self._skeleton_borderfill_count + len(self._cell_borderfills)
        # Sort by id so output is stable.
        ordered = sorted(self._cell_borderfills.items(), key=lambda kv: kv[1])
        parts: List[str] = []
        for fill_color, fid in ordered:
            face = fill_color if fill_color else "none"
            parts.append(
                f'<hh:borderFill id="{fid}" threeD="0" shadow="0" '
                'centerLine="NONE" breakCellSeparateLine="0">'
                '<hh:slash type="NONE" Crooked="0" isCounter="0"/>'
                '<hh:backSlash type="NONE" Crooked="0" isCounter="0"/>'
                '<hh:leftBorder type="SOLID" width="0.12 mm" color="#000000"/>'
                '<hh:rightBorder type="SOLID" width="0.12 mm" color="#000000"/>'
                '<hh:topBorder type="SOLID" width="0.12 mm" color="#000000"/>'
                '<hh:bottomBorder type="SOLID" width="0.12 mm" color="#000000"/>'
                '<hh:diagonal type="SOLID" width="0.12 mm" color="#000000"/>'
                '<hc:fillBrush>'
                f'<hc:winBrush faceColor="{face}" hatchColor="#999999" alpha="0"/>'
                '</hc:fillBrush>'
                '</hh:borderFill>'
            )
        new_xml = "".join(parts)
        header = re.sub(
            r'<hh:borderFills itemCnt="\d+">',
            f'<hh:borderFills itemCnt="{new_total}">',
            header,
        )
        header = header.replace("</hh:borderFills>", new_xml + "</hh:borderFills>")
        return header

    def _patch_fontfaces(self, header: str) -> str:
        """Append our new fonts to every <hh:fontface lang="X" fontCnt="N">
        block, bumping fontCnt accordingly."""
        # Build font XML pieces (one per new font).
        new_font_xml: Dict[str, str] = {}
        for offset, name in enumerate(self._new_fonts):
            fid = self._skeleton_font_count + offset
            new_font_xml[name] = (
                f'<hh:font id="{fid}" face="{xml_escape(name)}" '
                f'type="TTF" isEmbedded="0"/>'
            )
        joined = "".join(new_font_xml.values())
        added = len(self._new_fonts)

        def repl(m: re.Match) -> str:
            lang = m.group(1)
            old_cnt = int(m.group(2))
            new_cnt = old_cnt + added
            return f'<hh:fontface lang="{lang}" fontCnt="{new_cnt}">'

        # Step 1: patch all opening fontface tags.
        header = re.sub(
            r'<hh:fontface lang="(\w+)" fontCnt="(\d+)">',
            repl, header,
        )
        # Step 2: insert new font entries before each closing </hh:fontface>.
        header = header.replace("</hh:fontface>", joined + "</hh:fontface>")
        return header

    def _patch_charpr(self, header: str) -> str:
        new_total = self._skeleton_charpr_count + len(self._new_cps)
        new_xml = "".join(self._charpr_xml(cid, cp, fid) for cid, cp, fid in self._new_cps)
        header = re.sub(
            r'<hh:charProperties itemCnt="\d+">',
            f'<hh:charProperties itemCnt="{new_total}">',
            header,
        )
        header = header.replace("</hh:charProperties>", new_xml + "</hh:charProperties>")
        return header

    def _patch_parapr(self, header: str) -> str:
        new_total = self._skeleton_parapr_count + len(self._new_pps)
        new_xml = "".join(self._parapr_xml(pid, pp) for pid, pp in self._new_pps)
        header = re.sub(
            r'<hh:paraProperties itemCnt="\d+">',
            f'<hh:paraProperties itemCnt="{new_total}">',
            header,
        )
        header = header.replace("</hh:paraProperties>", new_xml + "</hh:paraProperties>")
        return header

    def _charpr_xml(self, cp_id: int, cp: CharProps, font_id: int) -> str:
        """Generate one <hh:charPr> block.

        Important schema rules learned from the Skeleton:
        - <hh:bold/>, <hh:italic/> are flag elements — present only when on.
        - <hh:strikeout/> has shape/color attrs but no text content.
        - <hh:underline type="..."/> always present (NONE means no underline).
        - <hh:emboss/>, <hh:engrave/>, <hh:supscript/>, <hh:subscript/> are
          flag elements — only present when on.
        - Only one borderFill ref (use 2 from skeleton's standard borderFill
          for "with diagonal" — same id used for skeleton's own charPrs).
        """
        height = halfpt_to_hwp_size(cp.size_halfpt)
        ul_type = "BOTTOM" if cp.underline else "NONE"
        ul_color = cp.color if cp.underline else "#000000"
        flag_bold = "<hh:bold/>" if cp.bold else ""
        flag_italic = "<hh:italic/>" if cp.italic else ""
        strikeout = ('<hh:strikeout shape="CONTINUOUS" color="#000000"/>'
                     if cp.strike
                     else '<hh:strikeout shape="NONE" color="#000000"/>')
        return (
            f'<hh:charPr id="{cp_id}" height="{height}" textColor="{cp.color}" '
            'shadeColor="none" useFontSpace="0" useKerning="0" symMark="NONE" '
            'borderFillIDRef="2">'
            f'<hh:fontRef hangul="{font_id}" latin="{font_id}" hanja="{font_id}" '
            f'japanese="{font_id}" other="{font_id}" symbol="{font_id}" user="{font_id}"/>'
            '<hh:ratio hangul="100" latin="100" hanja="100" japanese="100" '
            'other="100" symbol="100" user="100"/>'
            '<hh:spacing hangul="0" latin="0" hanja="0" japanese="0" '
            'other="0" symbol="0" user="0"/>'
            '<hh:relSz hangul="100" latin="100" hanja="100" japanese="100" '
            'other="100" symbol="100" user="100"/>'
            '<hh:offset hangul="0" latin="0" hanja="0" japanese="0" '
            'other="0" symbol="0" user="0"/>'
            f'{flag_italic}{flag_bold}'
            f'<hh:underline type="{ul_type}" shape="SOLID" color="{ul_color}"/>'
            f'{strikeout}'
            '<hh:outline type="NONE"/>'
            '<hh:shadow type="NONE" color="#C0C0C0" offsetX="10" offsetY="10"/>'
            '</hh:charPr>'
        )

    def _parapr_xml(self, pp_id: int, pp: ParaProps) -> str:
        """Generate <hh:paraPr> matching the skeleton's structure."""
        align = pp.align if pp.align in ("LEFT", "RIGHT", "CENTER", "JUSTIFY", "DISTRIBUTE") else "JUSTIFY"
        # heading: only OUTLINE is straightforward; bullets show in skel as
        # NONE plus a leading character — keep heading=NONE for both for now.
        if pp.list_type == "decimal":
            heading_lvl = min(((pp.list_level or 0) + 1), 7)
            heading_xml = f'<hh:heading type="OUTLINE" idRef="0" level="{heading_lvl}"/>'
            indent = (heading_lvl - 1) * 1000
        elif pp.list_type == "bullet":
            heading_xml = '<hh:heading type="NONE" idRef="0" level="0"/>'
            indent = ((pp.list_level or 0) + 1) * 1000
        else:
            heading_xml = '<hh:heading type="NONE" idRef="0" level="0"/>'
            indent = 0
        return (
            f'<hh:paraPr id="{pp_id}" tabPrIDRef="0" condense="0" '
            'fontLineHeight="0" snapToGrid="1" suppressLineNumbers="0" '
            'checked="0" textDir="LTR">'
            f'<hh:align horizontal="{align}" vertical="BASELINE"/>'
            f'{heading_xml}'
            '<hh:breakSetting breakLatinWord="KEEP_WORD" '
            'breakNonLatinWord="KEEP_WORD" widowOrphan="0" keepWithNext="0" '
            'keepLines="0" pageBreakBefore="0" lineWrap="BREAK"/>'
            '<hh:autoSpacing eAsianEng="0" eAsianNum="0"/>'
            '<hh:margin>'
            f'<hc:intent value="0" unit="HWPUNIT"/>'
            f'<hc:left value="{indent}" unit="HWPUNIT"/>'
            '<hc:right value="0" unit="HWPUNIT"/>'
            '<hc:prev value="0" unit="HWPUNIT"/>'
            '<hc:next value="0" unit="HWPUNIT"/>'
            '</hh:margin>'
            '<hh:lineSpacing type="PERCENT" value="160" unit="HWPUNIT"/>'
            '<hh:border borderFillIDRef="2" offsetLeft="0" offsetRight="0" '
            'offsetTop="0" offsetBottom="0" connect="0" ignoreMargin="0"/>'
            '</hh:paraPr>'
        )

    # -- section0.xml generation -----------------------------------------
    def _build_section(self) -> str:
        out: List[str] = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>']
        out.append(f'<hs:sec {_SECTION_NS_DECL}>')
        first = True
        for b in self.doc.blocks:
            if isinstance(b, Paragraph):
                out.append(self._paragraph_xml(b, with_secpr=first))
                first = False
            elif isinstance(b, Table):
                out.append(self._table_paragraph_xml(b, with_secpr=first))
                first = False
        if first:
            out.append(self._paragraph_xml(Paragraph(), with_secpr=True))
        out.append('</hs:sec>')
        return "".join(out)

    def _paragraph_xml(self, p: Paragraph, with_secpr: bool) -> str:
        """Emit a paragraph, splitting on PageBreak markers.

        A PageBreak in the inlines closes the current <hp:p> and opens a new
        <hp:p pageBreak="1" ...> for the rest. This matches how DOCX
        page breaks (w:br type="page") most commonly appear: in the middle
        of an otherwise normal paragraph.
        """
        pp_id = self._intern_pp(p.para_props)
        style_id = 0  # default to skeleton's "바탕글" style.

        # Split inlines on PageBreak boundaries.
        chunks: List[List[Inline]] = [[]]
        for inl in p.inlines:
            if isinstance(inl, PageBreak):
                chunks.append([])
            else:
                chunks[-1].append(inl)

        out: List[str] = []
        for ci, chunk in enumerate(chunks):
            para_id = self._next_para_id()
            page_break = "1" if ci > 0 else "0"
            chunk_with_secpr = with_secpr and ci == 0
            head = (
                f'<hp:p id="{para_id}" paraPrIDRef="{pp_id}" '
                f'styleIDRef="{style_id}" pageBreak="{page_break}" '
                'columnBreak="0" merged="0">'
            )
            body: List[str] = []
            if chunk_with_secpr:
                body.append(
                    '<hp:run charPrIDRef="0">'
                    + self._secpr_xml()
                    + '<hp:ctrl><hp:colPr id="" type="NEWSPAPER" layout="LEFT" '
                    'colCount="1" sameSz="1" sameGap="0"/></hp:ctrl>'
                    + '</hp:run>'
                )
            if chunk:
                for inl in chunk:
                    body.append(self._inline_run_xml(inl))
            elif chunk_with_secpr:
                body.append('<hp:run charPrIDRef="0"><hp:t/></hp:run>')
            else:
                body.append('<hp:run charPrIDRef="0"><hp:t/></hp:run>')
            body.append(_lineseg_xml())
            out.append(head + "".join(body) + "</hp:p>")
        return "".join(out)

    def _secpr_xml(self) -> str:
        d = self.doc
        landscape = "WIDELY" if not d.landscape else "NARROWLY"
        return (
            '<hp:secPr id="" textDirection="HORIZONTAL" spaceColumns="1134" '
            'tabStop="8000" tabStopVal="4000" tabStopUnit="HWPUNIT" '
            'outlineShapeIDRef="1" memoShapeIDRef="0" '
            'textVerticalWidthHead="0" masterPageCnt="0">'
            '<hp:grid lineGrid="0" charGrid="0" wonggojiFormat="0"/>'
            '<hp:startNum pageStartsOn="BOTH" page="0" pic="0" tbl="0" equation="0"/>'
            '<hp:visibility hideFirstHeader="0" hideFirstFooter="0" '
            'hideFirstMasterPage="0" border="SHOW_ALL" fill="SHOW_ALL" '
            'hideFirstPageNum="0" hideFirstEmptyLine="0" showLineNumber="0"/>'
            '<hp:lineNumberShape restartType="0" countBy="0" distance="0" '
            'startNumber="0"/>'
            f'<hp:pagePr landscape="{landscape}" width="{d.page_width_hwp}" '
            f'height="{d.page_height_hwp}" gutterType="LEFT_ONLY">'
            f'<hp:margin header="{d.margin_header_hwp}" footer="{d.margin_footer_hwp}" '
            f'gutter="0" left="{d.margin_left_hwp}" right="{d.margin_right_hwp}" '
            f'top="{d.margin_top_hwp}" bottom="{d.margin_bottom_hwp}"/>'
            '</hp:pagePr>'
            '<hp:footNotePr>'
            '<hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" '
            'suffixChar=")" supscript="0"/>'
            '<hp:noteLine length="-1" type="SOLID" width="0.12 mm" color="#000000"/>'
            '<hp:noteSpacing betweenNotes="283" belowLine="567" aboveLine="850"/>'
            '<hp:numbering type="CONTINUOUS" newNum="1"/>'
            '<hp:placement place="EACH_COLUMN" beneathText="0"/>'
            '</hp:footNotePr>'
            '<hp:endNotePr>'
            '<hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" '
            'suffixChar=")" supscript="0"/>'
            '<hp:noteLine length="14692344" type="SOLID" width="0.12 mm" color="#000000"/>'
            '<hp:noteSpacing betweenNotes="0" belowLine="567" aboveLine="850"/>'
            '<hp:numbering type="CONTINUOUS" newNum="1"/>'
            '<hp:placement place="END_OF_DOCUMENT" beneathText="0"/>'
            '</hp:endNotePr>'
            '<hp:pageBorderFill type="BOTH" borderFillIDRef="1" '
            'textBorder="PAPER" headerInside="0" footerInside="0" '
            'fillArea="PAPER">'
            '<hp:offset left="1417" right="1417" top="1417" bottom="1417"/>'
            '</hp:pageBorderFill>'
            '<hp:pageBorderFill type="EVEN" borderFillIDRef="1" '
            'textBorder="PAPER" headerInside="0" footerInside="0" '
            'fillArea="PAPER">'
            '<hp:offset left="1417" right="1417" top="1417" bottom="1417"/>'
            '</hp:pageBorderFill>'
            '<hp:pageBorderFill type="ODD" borderFillIDRef="1" '
            'textBorder="PAPER" headerInside="0" footerInside="0" '
            'fillArea="PAPER">'
            '<hp:offset left="1417" right="1417" top="1417" bottom="1417"/>'
            '</hp:pageBorderFill>'
            '</hp:secPr>'
        )

    def _inline_run_xml(self, inl: Inline) -> str:
        if isinstance(inl, Run):
            return self._run_xml(inl)
        if isinstance(inl, Hyperlink):
            return self._hyperlink_xml(inl)
        if isinstance(inl, Image):
            return self._image_run_xml(inl)
        if isinstance(inl, FootnoteRef):
            return self._footnote_run_xml(inl)
        return ""

    def _run_xml(self, r: Run) -> str:
        cp_id = self._intern_cp(r.char_props)
        return f'<hp:run charPrIDRef="{cp_id}">{self._t_xml(r.text)}</hp:run>'

    def _t_xml(self, text: str) -> str:
        """Build the inner content of a single <hp:t>: text mixed with
        <hp:tab/> and <hp:lineBreak/> elements (HWPX requires those control
        elements to be inside <hp:t>, not as siblings of the run)."""
        parts: List[str] = ['<hp:t>']
        for piece in re.split(r"(\t|\n|\f)", text):
            if piece == "":
                continue
            if piece == "\t":
                parts.append('<hp:tab/>')
            elif piece in ("\n", "\f"):
                parts.append('<hp:lineBreak/>')
            else:
                parts.append(xml_escape(piece))
        parts.append('</hp:t>')
        return "".join(parts)

    # -- hyperlink, image, footnote ---------------------------------------
    def _hyperlink_xml(self, h: Hyperlink) -> str:
        fid = self._next_ctrl_id()
        cp_id = (self._intern_cp(h.runs[0].char_props)
                 if h.runs else 0)
        out: List[str] = []
        out.append(
            f'<hp:run charPrIDRef="{cp_id}">'
            f'<hp:ctrl><hp:fieldBegin id="{fid}" type="HYPERLINK" '
            f'editable="1" dirty="0" name="" '
            f'command="{xml_escape(h.url)}"/></hp:ctrl>'
            '</hp:run>'
        )
        for r in h.runs:
            out.append(self._run_xml(r))
        out.append(
            f'<hp:run charPrIDRef="{cp_id}">'
            f'<hp:ctrl><hp:fieldEnd id="{fid}" type="HYPERLINK" '
            f'editable="1" dirty="0"/></hp:ctrl>'
            '</hp:run>'
        )
        return "".join(out)

    def _image_run_xml(self, img: Image) -> str:
        cp_id = self._intern_cp(img.char_props)
        cid = self._next_ctrl_id()
        w = max(1, img.width_hwp)
        h = max(1, img.height_hwp)
        return (
            f'<hp:run charPrIDRef="{cp_id}">'
            f'<hp:pic id="{cid}" zOrder="0" numberingType="PICTURE" '
            'textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" '
            'dropcapstyle="None" reverse="0" instid="0">'
            f'<hp:sz width="{w}" widthRelTo="ABSOLUTE" '
            f'height="{h}" heightRelTo="ABSOLUTE" protect="0"/>'
            '<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" '
            'allowOverlap="0" holdAnchorAndSO="0" vertRelTo="LINE" '
            'horzRelTo="COLUMN" vertAlign="TOP" horzAlign="LEFT" '
            'vertOffset="0" horzOffset="0"/>'
            '<hp:outMargin left="0" right="0" top="0" bottom="0"/>'
            '<hp:lineShape color="#000000" width="0.1 mm" style="SOLID" '
            'alpha="0" startSz="NORMAL" endSz="NORMAL" '
            'outlineStyle="NORMAL" arrowStart="NORMAL" arrowEnd="NORMAL"/>'
            '<hp:imgRect>'
            '<hc:pt x="0" y="0"/>'
            f'<hc:pt x="{w}" y="0"/>'
            f'<hc:pt x="{w}" y="{h}"/>'
            f'<hc:pt x="0" y="{h}"/>'
            '</hp:imgRect>'
            '<hp:imgClip left="0" right="0" top="0" bottom="0"/>'
            '<hp:inMargin left="0" right="0" top="0" bottom="0"/>'
            f'<hp:img binaryItemIDRef="{img.bin_id}" bright="0" contrast="0" '
            'effect="REAL_PIC"/>'
            '<hp:effects/>'
            '</hp:pic>'
            '</hp:run>'
        )

    def _footnote_run_xml(self, fn: FootnoteRef) -> str:
        cp_id = self._intern_cp(fn.char_props)
        cid = self._next_ctrl_id()
        if fn.is_endnote:
            self._endnote_counter += 1
            num = self._endnote_counter
            tag = "hp:endNote"
        else:
            self._footnote_counter += 1
            num = self._footnote_counter
            tag = "hp:footNote"
        sub_paras: List[str] = []
        for p in fn.paragraphs:
            sub_paras.append(self._cell_paragraph_xml(p))
        if not sub_paras:
            sub_paras.append(self._cell_paragraph_xml(Paragraph()))
        usable = (
            self.doc.page_width_hwp
            - self.doc.margin_left_hwp
            - self.doc.margin_right_hwp
        )
        return (
            f'<hp:run charPrIDRef="{cp_id}">'
            '<hp:ctrl>'
            f'<{tag} id="{cid}" num="{num}">'
            '<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" '
            'vertAlign="TOP" linkListIDRef="0" linkListNextIDRef="0" '
            f'textWidth="{usable}" textHeight="0" hasTextRef="0" '
            'hasNumRef="0">'
            + "".join(sub_paras)
            + f'</hp:subList></{tag}>'
            '</hp:ctrl>'
            '</hp:run>'
        )

    # -- table -----------------------------------------------------------
    def _table_paragraph_xml(self, t: Table, with_secpr: bool) -> str:
        para_id = self._next_para_id()
        head = (
            f'<hp:p id="{para_id}" paraPrIDRef="0" styleIDRef="0" '
            'pageBreak="0" columnBreak="0" merged="0">'
        )
        body: List[str] = []
        if with_secpr:
            body.append(
                '<hp:run charPrIDRef="0">'
                + self._secpr_xml()
                + '<hp:ctrl><hp:colPr id="" type="NEWSPAPER" layout="LEFT" '
                'colCount="1" sameSz="1" sameGap="0"/></hp:ctrl>'
                + '</hp:run>'
            )
        body.append(f'<hp:run charPrIDRef="0">{self._table_ctl_xml(t)}</hp:run>')
        body.append(_lineseg_xml())
        return head + "".join(body) + "</hp:p>"

    def _table_ctl_xml(self, t: Table) -> str:
        if not t.rows:
            return ""
        col_count = max(len(t.col_widths_hwp), 1)
        for row in t.rows:
            total = sum(c.grid_span for c in row.cells)
            col_count = max(col_count, total)
        row_count = len(t.rows)

        col_widths = list(t.col_widths_hwp)
        if len(col_widths) < col_count:
            usable = (
                self.doc.page_width_hwp
                - self.doc.margin_left_hwp
                - self.doc.margin_right_hwp
            )
            default = max(2000, usable // max(col_count, 1))
            col_widths += [default] * (col_count - len(col_widths))

        row_heights = []
        for row in t.rows:
            row_heights.append(row.height_hwp if row.height_hwp > 0 else 1500)
        total_w = sum(col_widths)
        total_h = sum(row_heights)
        tbl_id = self._next_ctrl_id()

        rows_xml: List[str] = []
        for ri, row in enumerate(t.rows):
            cell_xml: List[str] = []
            ci_grid = 0
            for cell in row.cells:
                span = max(1, cell.grid_span)
                if ci_grid + span > col_count:
                    span = max(1, col_count - ci_grid)
                if cell.v_merge_continue:
                    ci_grid += span
                    continue
                cell_w = sum(col_widths[ci_grid:ci_grid + span])
                cell_h = sum(row_heights[ri:ri + cell.row_span]) or row_heights[ri]
                cell_paras: List[str] = []
                for cb in cell.blocks:
                    if isinstance(cb, Paragraph):
                        cell_paras.append(self._cell_paragraph_xml(cb))
                    elif isinstance(cb, Table):
                        cell_paras.append(self._table_paragraph_xml(cb, with_secpr=False))
                if not cell_paras:
                    cell_paras.append(self._cell_paragraph_xml(Paragraph()))
                cell_border_id = self._intern_cell_fill(cell.fill_color)
                cell_xml.append(
                    '<hp:tc name="" header="0" hasMargin="0" protect="0" '
                    f'editable="1" dirty="0" borderFillIDRef="{cell_border_id}">'
                    '<hp:subList id="" textDirection="HORIZONTAL" '
                    'lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" '
                    'linkListNextIDRef="0" '
                    f'textWidth="{max(0, cell_w - 1020)}" textHeight="0" '
                    'hasTextRef="0" hasNumRef="0">'
                    + "".join(cell_paras)
                    + '</hp:subList>'
                    f'<hp:cellAddr colAddr="{ci_grid}" rowAddr="{ri}"/>'
                    f'<hp:cellSpan colSpan="{span}" rowSpan="{cell.row_span}"/>'
                    f'<hp:cellSz width="{cell_w}" height="{cell_h}"/>'
                    '<hp:cellMargin left="510" right="510" top="141" bottom="141"/>'
                    '</hp:tc>'
                )
                ci_grid += span
            rows_xml.append(f'<hp:tr>{"".join(cell_xml)}</hp:tr>')

        return (
            f'<hp:tbl id="{tbl_id}" zOrder="0" numberingType="TABLE" '
            'textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" '
            'dropcapstyle="None" pageBreak="CELL" repeatHeader="1" '
            f'rowCnt="{row_count}" colCnt="{col_count}" cellSpacing="0" '
            f'borderFillIDRef="{self._table_border_id}" noAdjust="0">'
            f'<hp:sz width="{total_w}" widthRelTo="ABSOLUTE" '
            f'height="{total_h}" heightRelTo="ABSOLUTE" protect="0"/>'
            '<hp:pos treatAsChar="0" affectLSpacing="0" flowWithText="1" '
            'allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" '
            'horzRelTo="COLUMN" vertAlign="TOP" horzAlign="LEFT" '
            'vertOffset="0" horzOffset="0"/>'
            '<hp:outMargin left="283" right="283" top="283" bottom="283"/>'
            '<hp:inMargin left="510" right="510" top="141" bottom="141"/>'
            f"{''.join(rows_xml)}"
            '</hp:tbl>'
        )

    def _cell_paragraph_xml(self, p: Paragraph) -> str:
        pp_id = self._intern_pp(p.para_props)
        para_id = 0
        runs_xml = [self._inline_run_xml(inl) for inl in p.inlines]
        if not runs_xml:
            runs_xml.append('<hp:run charPrIDRef="0"><hp:t/></hp:run>')
        return (
            f'<hp:p id="{para_id}" paraPrIDRef="{pp_id}" styleIDRef="0" '
            'pageBreak="0" columnBreak="0" merged="0">'
            + "".join(runs_xml)
            + _lineseg_xml()
            + '</hp:p>'
        )

    # -- content.hpf patching for binaries -------------------------------
    def _patch_content_hpf(self) -> bytes:
        if not self.doc.binaries:
            return self.skeleton["Contents/content.hpf"]
        hpf = self.skeleton["Contents/content.hpf"].decode("utf-8")
        new_items = "".join(
            f'<opf:item id="{b.bin_id}" href="{b.href}" '
            f'media-type="{b.media_type}"/>'
            for b in self.doc.binaries
        )
        # Insert before </opf:manifest>
        hpf = hpf.replace("</opf:manifest>", new_items + "</opf:manifest>")
        return hpf.encode("utf-8")

    # -- write ------------------------------------------------------------
    def write(self, out_path: str) -> None:
        new_header = self._patch_header().encode("utf-8")
        new_section = self._build_section().encode("utf-8")
        new_hpf = self._patch_content_hpf()

        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
            # mimetype must be FIRST and STORED.
            zi = zipfile.ZipInfo("mimetype")
            zi.compress_type = zipfile.ZIP_STORED
            z.writestr(zi, self.skeleton["mimetype"])

            for name, data in self.skeleton.items():
                if name == "mimetype":
                    continue
                if name == "Contents/header.xml":
                    z.writestr(name, new_header)
                elif name == "Contents/section0.xml":
                    z.writestr(name, new_section)
                elif name == "Contents/content.hpf":
                    z.writestr(name, new_hpf)
                else:
                    z.writestr(name, data)
            for b in self.doc.binaries:
                z.writestr(b.href, b.data)


def _lineseg_xml() -> str:
    return (
        '<hp:linesegarray>'
        '<hp:lineseg textpos="0" vertpos="0" vertsize="1000" '
        'textheight="1000" baseline="850" spacing="600" horzpos="0" '
        'horzsize="40000" flags="393216"/>'
        '</hp:linesegarray>'
    )


# ---------------------------------------------------------------------------
# Top-level
# ---------------------------------------------------------------------------
def _find_skeleton_path() -> str:
    here = os.path.dirname(os.path.abspath(__file__))
    candidate = os.path.join(here, SKELETON_FILENAME)
    if not os.path.isfile(candidate):
        raise FileNotFoundError(
            f"Skeleton template not found: {candidate}. "
            "Place a known-good HWPX file alongside this script."
        )
    return candidate


def convert(input_path: str, output_path: str,
            skeleton_path: Optional[str] = None) -> None:
    skel = skeleton_path or _find_skeleton_path()
    reader = DocxReader(input_path)
    try:
        doc = reader.parse()
    finally:
        reader.close()
    HwpxWriter(doc, skel).write(output_path)


def main() -> int:
    parser = argparse.ArgumentParser(description="Convert DOCX to HWPX.")
    parser.add_argument("input", help="Path to .docx file")
    parser.add_argument("output", help="Path to output .hwpx file")
    parser.add_argument("--skeleton", default=None,
                        help="Optional override for the HWPX skeleton template")
    args = parser.parse_args()
    try:
        convert(args.input, args.output, args.skeleton)
    except (FileNotFoundError, zipfile.BadZipFile) as e:
        print(f"error: {e}", file=sys.stderr)
        return 2
    print(f"wrote {args.output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
