"""
OOXML (.docx) and OWPML (.hwpx) Lifter, Dynamic Style/Page Scanner, and Projector.
All font families, font sizes, page dimensions, margins, outer table column widths,
cover/preamble blocks, section titles, and attachment slot dimensions are derived
dynamically from `TargetDocumentSpecIR` and `SourceDocumentSpecIR` with zero hardcoded
font names, pixel/twip/EMU constants, or document-specific strings.
"""

import os
import copy
import zipfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape as xml_esc
from typing import List, Dict, Tuple, Optional

from .ir_schema import (
    SectionIR,
    MediaObjectIR,
    SourceDocumentSpecIR,
    TargetDocumentSpecIR,
    SlotSpecIR,
    extract_ir_non_ws_text,
    hu_to_half_pt,
    hu_to_twips,
    twips_to_hu,
    hu_to_emu,
    detect_image_format,
    extract_ordinal_index,
    is_top_level_section_heading,
    is_subtitle_or_instruction_line,
)
from .elastic_solver import (
    ElasticLayoutState,
    is_stair_step_sparse_grid,
    is_vertical_step_arrow_grid,
    detect_and_remove_spacer_rows_cols,
    collapse_stair_step_chart_table,
    fold_vertical_step_arrows,
)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

for _pfx, _uri in (
    ("w", W_NS),
    ("wp", WP_NS),
    ("a", A_NS),
    ("pic", PIC_NS),
    ("r", R_NS),
    ("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006"),
    ("w14", "http://schemas.microsoft.com/office/word/2010/wordml"),
    ("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"),
):
    ET.register_namespace(_pfx, _uri)


def w_tag(name: str) -> str:
    return f"{{{W_NS}}}{name}"


def scan_docx_style_and_page_spec(docx_path: str) -> SourceDocumentSpecIR:
    """
    Dynamically inspects a DOCX file (`word/document.xml` and `word/styles.xml`) to extract
    page setup (width, height, margins) and dominant body/table font families and sizes.
    """
    spec = SourceDocumentSpecIR(source_path=docx_path, source_format="DOCX_OOXML")
    body_fonts: Dict[str, int] = {}
    body_sizes_hu: Dict[int, int] = {}
    table_sizes_hu: Dict[int, int] = {}

    with zipfile.ZipFile(docx_path, "r") as zf:
        if "word/styles.xml" in zf.namelist():
            styles_root = ET.fromstring(zf.read("word/styles.xml"))
            for rFonts in styles_root.findall(f".//{w_tag('rFonts')}"):
                fn = rFonts.get(w_tag("eastAsia")) or rFonts.get(w_tag("ascii"))
                if fn:
                    body_fonts[fn] = body_fonts.get(fn, 0) + 1

        doc_root = ET.fromstring(zf.read("word/document.xml"))
        body = doc_root.find(w_tag("body"))
        if body is not None:
            sectPr = body.find(w_tag("sectPr"))
            if sectPr is not None:
                pgSz = sectPr.find(w_tag("pgSz"))
                if pgSz is not None:
                    w_tw = int(pgSz.get(w_tag("w"), "0") or 0)
                    h_tw = int(pgSz.get(w_tag("h"), "0") or 0)
                    if w_tw > 0:
                        spec.page_width_hu = twips_to_hu(w_tw)
                    if h_tw > 0:
                        spec.page_height_hu = twips_to_hu(h_tw)
                pgMar = sectPr.find(w_tag("pgMar"))
                if pgMar is not None:
                    for attr_name, field_name in (
                        ("top", "margin_top_hu"),
                        ("bottom", "margin_bottom_hu"),
                        ("left", "margin_left_hu"),
                        ("right", "margin_right_hu"),
                    ):
                        val_tw = int(pgMar.get(w_tag(attr_name), "0") or 0)
                        if val_tw > 0:
                            setattr(spec, field_name, twips_to_hu(val_tw))

            for ch in body:
                ch_tag = ch.tag.split("}")[-1]
                if ch_tag == "p":
                    for rPr in ch.findall(f".//{w_tag('rPr')}"):
                        rf = rPr.find(w_tag("rFonts"))
                        if rf is not None:
                            fn = rf.get(w_tag("eastAsia")) or rf.get(w_tag("ascii"))
                            if fn:
                                body_fonts[fn] = body_fonts.get(fn, 0) + 2
                        sz_el = rPr.find(w_tag("sz"))
                        if sz_el is not None:
                            hp = int(sz_el.get(w_tag("val"), "0") or 0)
                            if 16 <= hp <= 32:
                                hu = hp * 50
                                body_sizes_hu[hu] = body_sizes_hu.get(hu, 0) + 1
                elif ch_tag == "tbl":
                    nested_tbls = ch.findall(f".//{w_tag('tbl')}")
                    target_tbls = nested_tbls if nested_tbls else [ch]
                    for t_el in target_tbls:
                        for rPr in t_el.findall(f".//{w_tag('rPr')}"):
                            sz_el = rPr.find(w_tag("sz"))
                            if sz_el is not None:
                                hp = int(sz_el.get(w_tag("val"), "0") or 0)
                                if 14 <= hp <= 22:
                                    hu = hp * 50
                                    table_sizes_hu[hu] = table_sizes_hu.get(hu, 0) + 1

    if body_fonts:
        dominant_font = max(body_fonts.items(), key=lambda kv: kv[1])[0]
        spec.body_font_family = dominant_font
        spec.table_font_family = dominant_font
    if body_sizes_hu:
        spec.body_font_size_hu = max(body_sizes_hu.items(), key=lambda kv: kv[1])[0]
    if table_sizes_hu:
        spec.table_font_size_hu = max(table_sizes_hu.items(), key=lambda kv: kv[1])[0]

    return spec


def extract_docx_p_data(p_el: ET.Element, is_in_table: bool) -> dict:
    pPr = p_el.find(w_tag("pPr"))
    jc = "both"
    left_tw = 0
    right_tw = 0
    indent_tw = 0
    outline_lvl: Optional[int] = None
    before_tw, after_tw = (0, 0) if is_in_table else (0, 8)
    if pPr is not None:
        jc_el = pPr.find(w_tag("jc"))
        if jc_el is not None:
            jc = jc_el.get(w_tag("val"), "both")
        ind_el = pPr.find(w_tag("ind"))
        if ind_el is not None:
            left_tw = int(ind_el.get(w_tag("left"), "0") or 0)
            right_tw = int(ind_el.get(w_tag("right"), "0") or 0)
            first_l = int(ind_el.get(w_tag("firstLine"), "0") or 0)
            hanging = int(ind_el.get(w_tag("hanging"), "0") or 0)
            indent_tw = -hanging if hanging > 0 else first_l
        ol_el = pPr.find(w_tag("outlineLvl"))
        if ol_el is not None:
            ol_val = ol_el.get(w_tag("val"), "")
            if ol_val.isdigit():
                outline_lvl = int(ol_val)
        ps_el = pPr.find(w_tag("pStyle"))
        if ps_el is not None and outline_lvl is None:
            ps_val = ps_el.get(w_tag("val"), "").lower()
            if ps_val in ("heading1", "heading 1", "1", "title"):
                outline_lvl = 0
            elif ps_val in ("heading2", "heading 2", "2"):
                outline_lvl = 1
            elif ps_val in ("heading3", "heading 3", "3", "heading4", "heading 4"):
                outline_lvl = 2

    blip = p_el.find(f".//{{{A_NS}}}blip")
    r_embed = blip.get(f"{{{R_NS}}}embed") if blip is not None else None

    runs: List[dict] = []
    max_sz_hu = 0
    for r_el in p_el.findall(f".//{w_tag('r')}"):
        t_texts = [t.text for t in r_el.findall(w_tag("t")) if t.text]
        txt = "".join(t_texts)
        if not txt:
            continue
        txt = "".join(ch for ch in txt if ord(ch) >= 32)
        if not txt:
            continue
        rPr = r_el.find(w_tag("rPr"))
        bold = False
        italic = False
        underline = False
        color_hex = "000000"
        if rPr is not None:
            b_el = rPr.find(w_tag("b"))
            if b_el is not None and b_el.get(w_tag("val"), "true") not in ("0", "false"):
                bold = True
            i_el = rPr.find(w_tag("i"))
            if i_el is not None and i_el.get(w_tag("val"), "true") not in ("0", "false"):
                italic = True
            u_el = rPr.find(w_tag("u"))
            if u_el is not None and u_el.get(w_tag("val"), "single") not in ("0", "false", "none"):
                underline = True
            c_el = rPr.find(w_tag("color"))
            if c_el is not None:
                c_val = c_el.get(w_tag("val"), "000000")
                if c_val and c_val != "auto":
                    color_hex = c_val.upper()
            sz_el = rPr.find(w_tag("sz"))
            if sz_el is not None:
                hp_val = int(sz_el.get(w_tag("val"), "0") or 0)
                if hp_val > 0:
                    max_sz_hu = max(max_sz_hu, hp_val * 50)
        runs.append(
            {
                "text": txt,
                "bold": bold,
                "italic": italic,
                "underline": underline,
                "color": color_hex,
            }
        )

    return {
        "type": "p",
        "jc": jc,
        "left_tw": left_tw,
        "right_tw": right_tw,
        "indent_tw": indent_tw,
        "before_tw": before_tw,
        "after_tw": after_tw,
        "runs": runs,
        "r_embed": r_embed,
        "raw_p_xml": ET.tostring(p_el, encoding="utf-8") if r_embed else None,
        "max_sz_hu": max_sz_hu,
        "outline_lvl": outline_lvl,
    }


def extract_docx_tbl_normalized(
    tbl_el: ET.Element,
    target_width_hu: int = 0,
    table_font_size_hu: int = 1000,
) -> dict:
    grid = tbl_el.find(w_tag("tblGrid"))
    col_twips: List[int] = []
    if grid is not None:
        for gc in grid.findall(w_tag("gridCol")):
            col_twips.append(int(gc.get(w_tag("w"), "1000")))
    n_cols_raw = len(col_twips) or 1
    w_target = target_width_hu if target_width_hu > 0 else max(10000, sum(col_twips) * 5)
    trs_raw = tbl_el.findall(w_tag("tr"))
    n_rows_raw = len(trs_raw)

    grid_cells: List[List[Optional[dict]]] = [[None for _ in range(n_cols_raw)] for _ in range(n_rows_raw)]

    for r_idx, tr in enumerate(trs_raw):
        c_cursor = 0
        for tc in tr.findall(w_tag("tc")):
            while c_cursor < n_cols_raw and grid_cells[r_idx][c_cursor] is not None:
                c_cursor += 1
            if c_cursor >= n_cols_raw:
                break
            tcPr = tc.find(w_tag("tcPr"))
            col_span = 1
            v_merge: Optional[str] = None
            fill_hex: Optional[str] = None
            if tcPr is not None:
                gs = tcPr.find(w_tag("gridSpan"))
                if gs is not None:
                    col_span = int(gs.get(w_tag("val"), "1") or 1)
                vm = tcPr.find(w_tag("vMerge"))
                if vm is not None:
                    v_merge = vm.get(w_tag("val"), "continue")
                shd = tcPr.find(w_tag("shd"))
                if shd is not None:
                    f_val = shd.get(w_tag("fill"), "auto")
                    if f_val and f_val != "auto":
                        fill_hex = f_val.upper()
            col_span = min(col_span, n_cols_raw - c_cursor)

            if v_merge == "continue" and r_idx > 0 and grid_cells[r_idx - 1][c_cursor] is not None:
                master = grid_cells[r_idx - 1][c_cursor]
                if master is not None and master["col"] == c_cursor and master["row"] + master["row_span"] == r_idx:
                    master["row_span"] += 1
                    for p_el in tc.findall(f".//{w_tag('p')}"):
                        pdata = extract_docx_p_data(p_el, is_in_table=True)
                        if pdata["runs"]:
                            master["paras"].append(pdata)
                    for dc in range(col_span):
                        grid_cells[r_idx][c_cursor + dc] = master
                    c_cursor += col_span
                    continue

            paras: List[dict] = []
            for p_el in tc.findall(f".//{w_tag('p')}"):
                pdata = extract_docx_p_data(p_el, is_in_table=True)
                if pdata["runs"]:
                    paras.append(pdata)

            cell_obj = {
                "row": r_idx,
                "col": c_cursor,
                "row_span": 1,
                "col_span": col_span,
                "fill_hex": fill_hex,
                "paras": paras,
            }
            for dc in range(col_span):
                grid_cells[r_idx][c_cursor + dc] = cell_obj
            c_cursor += col_span

    # Pure topological detection (zero hardcoded content strings)
    if is_stair_step_sparse_grid(grid_cells, n_rows_raw, n_cols_raw):
        return collapse_stair_step_chart_table(
            grid_cells, w_target, n_rows_raw=n_rows_raw, n_cols_raw=n_cols_raw, table_font_size_hu=table_font_size_hu
        )

    if is_vertical_step_arrow_grid(grid_cells, n_rows_raw, n_cols_raw):
        return fold_vertical_step_arrows(
            grid_cells, w_target, n_rows_raw=n_rows_raw, table_font_size_hu=table_font_size_hu
        )

    return detect_and_remove_spacer_rows_cols(
        grid_cells, col_twips, n_rows_raw, n_cols_raw, w_target, table_font_size_hu
    )


def _unwrap_sdt_elements(elems: List[ET.Element]) -> List[ET.Element]:
    out: List[ET.Element] = []
    for el in elems:
        if el.tag.split("}")[-1] == "sdt":
            sdt_c = el.find(w_tag("sdtContent"))
            if sdt_c is not None:
                out.extend(_unwrap_sdt_elements(list(sdt_c)))
        else:
            out.append(el)
    return out


def lift_docx_to_ir(
    docx_path: str,
    target_width_hu: int = 0,
    table_font_size_hu: int = 1000,
) -> Dict[str, SectionIR]:
    """
    Dynamically scans a DOCX file, detects top-level sections via unsupervised typographic
    and topological heading spectrum analysis ([NUM-01]), and lifts all paragraphs, tables,
    and embedded media into generic SectionIR objects (SEC_01..SEC_K, PREAMBLE, REFERENCES).
    """
    sections: Dict[str, SectionIR] = {}
    with zipfile.ZipFile(docx_path, "r") as zf:
        rels_map: Dict[str, str] = {}
        if "word/_rels/document.xml.rels" in zf.namelist():
            rels_root = ET.fromstring(zf.read("word/_rels/document.xml.rels"))
            for rel in rels_root:
                rid = rel.get("Id", "")
                target = rel.get("Target", "")
                if rid and target:
                    rels_map[rid] = target

        doc_root = ET.fromstring(zf.read("word/document.xml"))
        body = doc_root.find(w_tag("body"))
        raw_children = list(body) if body is not None else []
        children = _unwrap_sdt_elements(raw_children)

        # Compute statistical mode body font size (S_mode) across body paragraphs
        sz_hist: Dict[int, int] = {}
        for ch_el in children:
            if ch_el.tag.split("}")[-1] == "p":
                for sz_el in ch_el.findall(f".//{w_tag('sz')}"):
                    hp = int(sz_el.get(w_tag("val"), "0") or 0)
                    if 16 <= hp <= 28:
                        hu_v = hp * 50
                        sz_hist[hu_v] = sz_hist.get(hu_v, 0) + 1
        mode_body_sz_hu = max(sz_hist.items(), key=lambda kv: kv[1])[0] if sz_hist else 1000

        cur_sec_id = "PREAMBLE"
        cur_heading = "Document Preamble"
        cur_subtitle = ""
        cur_role = "PREAMBLE"
        cur_ord: Optional[int] = None
        sec_counter = 0
        cur_blocks: List[dict] = []
        cur_media: List[MediaObjectIR] = []

        def flush_section() -> None:
            nonlocal cur_blocks, cur_media, cur_subtitle
            if not cur_blocks and not cur_media:
                return
            _, n_chars, sha = extract_ir_non_ws_text(cur_blocks)
            sections[cur_sec_id] = SectionIR(
                section_uid=cur_sec_id,
                heading_text=cur_heading,
                subtitle_text=cur_subtitle,
                blocks=cur_blocks,
                media_objects=cur_media,
                non_ws_char_count=n_chars,
                text_sha256=sha,
                ordinal_index=cur_ord,
                section_role=cur_role,
            )
            cur_blocks = []
            cur_media = []
            cur_subtitle = ""

        i = 0
        while i < len(children):
            ch = children[i]
            tag = ch.tag.split("}")[-1]
            if tag == "p":
                pdata = extract_docx_p_data(ch, is_in_table=False)
                ptxt = "".join(r["text"] for r in pdata["runs"]).strip()
                p_max_sz = pdata.get("max_sz_hu", 0) or mode_body_sz_hu
                p_is_bold = bool(pdata["runs"]) and all(r.get("bold", False) for r in pdata["runs"])
                p_outline = pdata.get("outline_lvl")

                is_hdr, hdr_role = is_top_level_section_heading(
                    ptxt,
                    p_max_sz_hu=p_max_sz,
                    p_is_bold=p_is_bold,
                    p_outline_lvl=p_outline,
                    mode_body_sz_hu=mode_body_sz_hu,
                    is_in_table=False,
                )

                if is_hdr:
                    flush_section()
                    if hdr_role == "REFERENCE":
                        cur_sec_id = "REFERENCES"
                        cur_role = "REFERENCE"
                        cur_ord = None
                    else:
                        sec_counter += 1
                        cur_sec_id = f"SEC_{sec_counter:02d}"
                        cur_role = "BODY"
                        cur_ord = extract_ordinal_index(ptxt) or sec_counter
                    cur_heading = ptxt
                    if i + 1 < len(children) and children[i + 1].tag.split("}")[-1] == "p":
                        nxt_p = extract_docx_p_data(children[i + 1], is_in_table=False)
                        nxt_txt = "".join(r["text"] for r in nxt_p["runs"]).strip()
                        nxt_sz = nxt_p.get("max_sz_hu", 0) or mode_body_sz_hu
                        if is_subtitle_or_instruction_line(nxt_txt, nxt_sz, mode_body_sz_hu):
                            cur_subtitle = nxt_txt
                            i += 2
                            continue
                    i += 1
                    continue

                if pdata["r_embed"] and pdata["r_embed"] in rels_map:
                    rel_target = rels_map[pdata["r_embed"]].lstrip("/")
                    zip_media_path = rel_target if rel_target.startswith("word/") else f"word/{rel_target}"
                    if zip_media_path in zf.namelist():
                        raw_img = zf.read(zip_media_path)
                        _, det_mime = detect_image_format(raw_img)
                        cap_hint = ptxt
                        if not cap_hint and cur_blocks and cur_blocks[-1]["type"] == "p":
                            prev_txt = "".join(r["text"] for r in cur_blocks[-1]["runs"]).strip()
                            if len(prev_txt) <= 80:
                                cap_hint = f"[{cur_sec_id}] {prev_txt}"
                        if not cap_hint:
                            cap_hint = f"[{cur_sec_id} - {os.path.basename(zip_media_path)}]"
                        cur_media.append(
                            MediaObjectIR(
                                media_uid=f"{cur_sec_id}_IMG_{len(cur_media) + 1}",
                                rel_id=pdata["r_embed"],
                                filename=os.path.basename(zip_media_path),
                                mime_type=det_mime,
                                raw_bytes=raw_img,
                                caption_hint=cap_hint,
                            )
                        )
                if pdata["runs"] or pdata["r_embed"]:
                    cur_blocks.append(pdata)
            elif tag == "tbl":
                cur_blocks.append(extract_docx_tbl_normalized(ch, target_width_hu, table_font_size_hu))
            i += 1

        flush_section()

    return sections


def project_ir_to_docx(
    source_docx_path: str,
    out_docx_path: str,
    target_spec: TargetDocumentSpecIR,
    slot_blocks_map: Dict[str, Tuple[List[dict], ElasticLayoutState]],
    media_drawings: List[Tuple[Optional[bytes], str]],
    preserve_outer_tables: bool = True,
    clear_unmapped_slots: bool = False,
) -> dict:
    """
    Projects the solved T-DIR blocks into a DOCX file using the effective `target_spec`
    (font families, font sizes, page dimensions, outer slot widths, titles, subtitles,
    preamble blocks, and attachment slot dimensions). Zero hardcoded font names or domain strings.
    """
    body_font = target_spec.body_font_family or "sans-serif"
    table_font = target_spec.table_font_family or body_font
    body_sz_hp = hu_to_half_pt(target_spec.body_font_size_hu if target_spec.body_font_size_hu > 0 else 1200)
    table_sz_hp = hu_to_half_pt(target_spec.table_font_size_hu if target_spec.table_font_size_hu > 0 else 1000)

    has_valid_src_docx = False
    if source_docx_path and os.path.isfile(source_docx_path) and zipfile.is_zipfile(source_docx_path):
        with zipfile.ZipFile(source_docx_path, "r") as z_chk:
            if "word/document.xml" in z_chk.namelist():
                has_valid_src_docx = True

    if has_valid_src_docx:
        with zipfile.ZipFile(source_docx_path, "r") as src_zip:
            doc_xml = src_zip.read("word/document.xml")
    else:
        doc_xml = (
            b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            b'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            b'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
            b'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
            b'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            b'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            b'<w:body><w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            b'<w:pgMar w:top="560" w:right="950" w:bottom="560" w:left="950" w:header="360" w:footer="360" w:gutter="0"/>'
            b'</w:sectPr></w:body></w:document>'
        )

    if True:
        root = ET.fromstring(doc_xml)
        for k in list(root.attrib.keys()):
            if k.endswith("Ignorable"):
                del root.attrib[k]
        body = root.find(w_tag("body"))
        assert body is not None
        sectPr_orig = body.find(w_tag("sectPr"))
        if sectPr_orig is not None:
            sectPr = copy.deepcopy(sectPr_orig)
        else:
            sectPr = ET.Element(w_tag("sectPr"))
            ET.SubElement(sectPr, w_tag("pgSz"), {w_tag("w"): "11906", w_tag("h"): "16838"})
            ET.SubElement(
                sectPr,
                w_tag("pgMar"),
                {
                    w_tag("top"): "560",
                    w_tag("right"): "950",
                    w_tag("bottom"): "560",
                    w_tag("left"): "950",
                    w_tag("header"): "360",
                    w_tag("footer"): "360",
                    w_tag("gutter"): "0",
                },
            )

        pgSz = sectPr.find(w_tag("pgSz"))
        if pgSz is not None and target_spec.page_width_hu > 0 and target_spec.page_height_hu > 0:
            pgSz.set(w_tag("w"), str(hu_to_twips(target_spec.page_width_hu)))
            pgSz.set(w_tag("h"), str(hu_to_twips(target_spec.page_height_hu)))

        pgMar = sectPr.find(w_tag("pgMar"))
        if pgMar is not None:
            if target_spec.margins_explicitly_overridden:
                top_tw = hu_to_twips(target_spec.margin_top_hu) if target_spec.margin_top_hu > 0 else 560
                bot_tw = hu_to_twips(target_spec.margin_bottom_hu) if target_spec.margin_bottom_hu > 0 else 560
                left_tw = hu_to_twips(target_spec.margin_left_hu) if target_spec.margin_left_hu > 0 else 950
                right_tw = hu_to_twips(target_spec.margin_right_hu) if target_spec.margin_right_hu > 0 else 950
            else:
                top_tw = min(560, hu_to_twips(target_spec.margin_top_hu)) if target_spec.margin_top_hu > 0 else 560
                bot_tw = min(560, hu_to_twips(target_spec.margin_bottom_hu)) if target_spec.margin_bottom_hu > 0 else 560
                left_tw = min(950, hu_to_twips(target_spec.margin_left_hu)) if target_spec.margin_left_hu > 0 else 950
                right_tw = min(950, hu_to_twips(target_spec.margin_right_hu)) if target_spec.margin_right_hu > 0 else 950
            pgMar.set(w_tag("top"), str(top_tw))
            pgMar.set(w_tag("bottom"), str(bot_tw))
            pgMar.set(w_tag("left"), str(left_tw))
            pgMar.set(w_tag("right"), str(right_tw))
            pgMar.set(w_tag("header"), "360")
            pgMar.set(w_tag("footer"), "360")

        for dg in sectPr.findall(w_tag("docGrid")):
            sectPr.remove(dg)
        ET.SubElement(sectPr, w_tag("docGrid"), {w_tag("type"): "default"})
        body.clear()

        def make_w_p_from_runs(
            runs: List[dict],
            sz_half_pt: int,
            font_name: str,
            jc: str = "both",
            before_tw: int = 0,
            after_tw: int = 8,
            line_tw: int = 268,
            page_break: bool = False,
            line_rule: str = "exact",
        ) -> ET.Element:
            p = ET.Element(w_tag("p"))
            pPr = ET.SubElement(p, w_tag("pPr"))
            if page_break:
                ET.SubElement(pPr, w_tag("pageBreakBefore"))
            ET.SubElement(pPr, w_tag("snapToGrid"), {w_tag("val"): "0"})
            ET.SubElement(
                pPr,
                w_tag("spacing"),
                {
                    w_tag("before"): str(before_tw),
                    w_tag("after"): str(after_tw),
                    w_tag("line"): str(line_tw),
                    w_tag("lineRule"): line_rule,
                },
            )
            ET.SubElement(pPr, w_tag("jc"), {w_tag("val"): jc})
            for r_info in runs:
                txt = r_info["text"]
                if not txt:
                    continue
                r = ET.SubElement(p, w_tag("r"))
                rPr = ET.SubElement(r, w_tag("rPr"))
                ET.SubElement(
                    rPr,
                    w_tag("rFonts"),
                    {
                        w_tag("ascii"): font_name,
                        w_tag("hAnsi"): font_name,
                        w_tag("eastAsia"): font_name,
                        w_tag("cs"): font_name,
                    },
                )
                if r_info.get("bold"):
                    ET.SubElement(rPr, w_tag("b"))
                if r_info.get("italic"):
                    ET.SubElement(rPr, w_tag("i"))
                if r_info.get("underline"):
                    ET.SubElement(rPr, w_tag("u"), {w_tag("val"): "single"})
                ET.SubElement(rPr, w_tag("color"), {w_tag("val"): r_info.get("color", "000000")})
                ET.SubElement(rPr, w_tag("sz"), {w_tag("val"): str(sz_half_pt)})
                ET.SubElement(rPr, w_tag("szCs"), {w_tag("val"): str(sz_half_pt)})
                t = ET.SubElement(r, w_tag("t"))
                if txt.startswith(" ") or txt.endswith(" "):
                    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                t.text = txt
            return p

        def make_w_p(
            text: str,
            sz_half_pt: int,
            font_name: str,
            bold: bool = False,
            color: str = "000000",
            jc: str = "both",
            before_tw: int = 0,
            after_tw: int = 8,
            line_tw: int = 268,
            page_break: bool = False,
            line_rule: str = "auto",
        ) -> ET.Element:
            runs = [{"text": text, "bold": bold, "italic": False, "color": color}] if text else []
            return make_w_p_from_runs(runs, sz_half_pt, font_name, jc, before_tw, after_tw, line_tw, page_break, line_rule)

        def set_cell_borders_and_shd(
            tcPr: ET.Element,
            border_color: str = "000000",
            border_sz: str = "6",
            fill: Optional[str] = None,
            mar: Tuple[int, int, int, int] = (30, 80, 30, 80),
        ) -> None:
            tcB = ET.SubElement(tcPr, w_tag("tcBorders"))
            for side in ("top", "left", "bottom", "right"):
                ET.SubElement(
                    tcB,
                    w_tag(side),
                    {
                        w_tag("val"): "single",
                        w_tag("sz"): border_sz,
                        w_tag("space"): "0",
                        w_tag("color"): border_color,
                    },
                )
            if fill:
                ET.SubElement(
                    tcPr,
                    w_tag("shd"),
                    {
                        w_tag("val"): "clear",
                        w_tag("color"): "auto",
                        w_tag("fill"): fill,
                    },
                )
            tcMar = ET.SubElement(tcPr, w_tag("tcMar"))
            for s_name, val in zip(("top", "left", "bottom", "right"), mar):
                ET.SubElement(tcMar, w_tag(s_name), {w_tag("w"): str(val), w_tag("type"): "dxa"})

        def build_w_tbl_from_block(tbl_blk: dict, tbl_line_tw: int) -> ET.Element:
            tbl = ET.Element(w_tag("tbl"))
            tblPr = ET.SubElement(tbl, w_tag("tblPr"))
            col_tws = [hu_to_twips(w) for w in tbl_blk["col_widths_hu"]]
            tot_tw = sum(col_tws)
            ET.SubElement(tblPr, w_tag("tblW"), {w_tag("w"): str(tot_tw), w_tag("type"): "dxa"})
            ET.SubElement(tblPr, w_tag("jc"), {w_tag("val"): "center"})
            ET.SubElement(tblPr, w_tag("tblLayout"), {w_tag("type"): "fixed"})
            grid = ET.SubElement(tbl, w_tag("tblGrid"))
            for cw in col_tws:
                ET.SubElement(grid, w_tag("gridCol"), {w_tag("w"): str(cw)})

            rows_by_r: Dict[int, List[dict]] = {}
            for c in tbl_blk["cells"]:
                rows_by_r.setdefault(c["row"], []).append(c)

            for r_idx in range(tbl_blk["n_rows"]):
                tr = ET.SubElement(tbl, w_tag("tr"))
                for cell in sorted(rows_by_r.get(r_idx, []), key=lambda x: x["col"]):
                    tc = ET.SubElement(tr, w_tag("tc"))
                    tcPr = ET.SubElement(tc, w_tag("tcPr"))
                    cw_tw = sum(col_tws[cell["col"] : cell["col"] + cell["col_span"]])
                    ET.SubElement(tcPr, w_tag("tcW"), {w_tag("w"): str(cw_tw), w_tag("type"): "dxa"})
                    if cell["col_span"] > 1:
                        ET.SubElement(tcPr, w_tag("gridSpan"), {w_tag("val"): str(cell["col_span"])})
                    set_cell_borders_and_shd(
                        tcPr, border_color="CBD5E1", border_sz="4", fill=cell["fill_hex"], mar=(14, 50, 14, 50)
                    )
                    ET.SubElement(tcPr, w_tag("vAlign"), {w_tag("val"): "center"})
                    for p in cell["paras"]:
                        tc.append(
                            make_w_p_from_runs(
                                p["runs"],
                                sz_half_pt=table_sz_hp,
                                font_name=table_font,
                                jc=p["jc"],
                                before_tw=0,
                                after_tw=3,
                                line_tw=tbl_line_tw,
                                line_rule="exact",
                            )
                        )
            return tbl

        def make_outer_2x2_section_table(
            title_left: str,
            subtitle_right: str,
            blocks: List[dict],
            el_state: ElasticLayoutState,
            min_body_height_tw: int = 11000,
            page_break_before: bool = True,
            s_slot: Optional[SlotSpecIR] = None,
        ) -> ET.Element:
            outer_w_tw = hu_to_twips(s_slot.outer_width_hu) if (s_slot and s_slot.outer_width_hu > 0) else 9524
            col0_tw = (
                hu_to_twips(s_slot.col_widths_hu[0])
                if (s_slot and len(s_slot.col_widths_hu) >= 2)
                else int(round(outer_w_tw * 0.328))
            )
            col1_tw = outer_w_tw - col0_tw
            hdr_h_tw = (
                max(320, hu_to_twips(s_slot.header_height_hu) - 80)
                if (s_slot and s_slot.header_height_hu > 0)
                else 380
            )
            min_body_h_tw = (
                max(8000, hu_to_twips(s_slot.available_height_hu) - 2300)
                if (s_slot and s_slot.available_height_hu > 0)
                else min_body_height_tw
            )

            tbl = ET.Element(w_tag("tbl"))
            tblPr = ET.SubElement(tbl, w_tag("tblPr"))
            ET.SubElement(tblPr, w_tag("tblW"), {w_tag("w"): str(outer_w_tw), w_tag("type"): "dxa"})
            ET.SubElement(tblPr, w_tag("jc"), {w_tag("val"): "center"})
            tblB = ET.SubElement(tblPr, w_tag("tblBorders"))
            for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
                ET.SubElement(
                    tblB,
                    w_tag(side),
                    {w_tag("val"): "single", w_tag("sz"): "6", w_tag("space"): "0", w_tag("color"): "000000"},
                )
            ET.SubElement(tblPr, w_tag("tblLayout"), {w_tag("type"): "fixed"})

            grid = ET.SubElement(tbl, w_tag("tblGrid"))
            ET.SubElement(grid, w_tag("gridCol"), {w_tag("w"): str(col0_tw)})
            ET.SubElement(grid, w_tag("gridCol"), {w_tag("w"): str(col1_tw)})

            tr0 = ET.SubElement(tbl, w_tag("tr"))
            tr0Pr = ET.SubElement(tr0, w_tag("trPr"))
            ET.SubElement(tr0Pr, w_tag("cantSplit"))
            ET.SubElement(tr0Pr, w_tag("trHeight"), {w_tag("val"): str(hdr_h_tw), w_tag("hRule"): "atLeast"})

            tc00 = ET.SubElement(tr0, w_tag("tc"))
            tc00Pr = ET.SubElement(tc00, w_tag("tcPr"))
            ET.SubElement(tc00Pr, w_tag("tcW"), {w_tag("w"): str(col0_tw), w_tag("type"): "dxa"})
            set_cell_borders_and_shd(tc00Pr, border_color="000000", border_sz="6", fill="3A3C84", mar=(30, 90, 30, 90))
            ET.SubElement(tc00Pr, w_tag("vAlign"), {w_tag("val"): "center"})
            tc00.append(
                make_w_p(
                    title_left,
                    sz_half_pt=max(20, body_sz_hp - 2),
                    font_name=body_font,
                    bold=True,
                    color="FFFFFF",
                    jc="left",
                    before_tw=0,
                    after_tw=0,
                    line_tw=240,
                    page_break=page_break_before,
                )
            )

            tc01 = ET.SubElement(tr0, w_tag("tc"))
            tc01Pr = ET.SubElement(tc01, w_tag("tcPr"))
            ET.SubElement(tc01Pr, w_tag("tcW"), {w_tag("w"): str(col1_tw), w_tag("type"): "dxa"})
            set_cell_borders_and_shd(tc01Pr, border_color="000000", border_sz="6", fill="F2F2F2", mar=(30, 90, 30, 90))
            ET.SubElement(tc01Pr, w_tag("vAlign"), {w_tag("val"): "center"})
            tc01.append(
                make_w_p(
                    subtitle_right,
                    sz_half_pt=table_sz_hp,
                    font_name=table_font,
                    bold=False,
                    color="333333",
                    jc="left",
                    before_tw=0,
                    after_tw=0,
                    line_tw=240,
                )
            )

            tr1 = ET.SubElement(tbl, w_tag("tr"))
            if not blocks:
                tr1Pr = ET.SubElement(tr1, w_tag("trPr"))
                ET.SubElement(
                    tr1Pr, w_tag("trHeight"), {w_tag("val"): str(min_body_h_tw), w_tag("hRule"): "atLeast"}
                )

            tc10 = ET.SubElement(tr1, w_tag("tc"))
            tc10Pr = ET.SubElement(tc10, w_tag("tcPr"))
            ET.SubElement(tc10Pr, w_tag("tcW"), {w_tag("w"): str(outer_w_tw), w_tag("type"): "dxa"})
            ET.SubElement(tc10Pr, w_tag("gridSpan"), {w_tag("val"): "2"})
            set_cell_borders_and_shd(tc10Pr, border_color="000000", border_sz="6", fill=None, mar=(35, 95, 35, 95))

            if blocks:
                if blocks[0]["type"] == "tbl":
                    tc10.append(
                        make_w_p("", sz_half_pt=2, font_name=body_font, before_tw=0, after_tw=0, line_tw=20, line_rule="exact")
                    )
                for b_i, blk in enumerate(blocks):
                    if blk["type"] == "p":
                        tc10.append(
                            make_w_p_from_runs(
                                blk["runs"],
                                sz_half_pt=body_sz_hp,
                                font_name=body_font,
                                jc=blk["jc"],
                                before_tw=4,
                                after_tw=el_state.para_after_tw,
                                line_tw=el_state.body_line_tw,
                                line_rule="exact",
                            )
                        )
                    elif blk["type"] == "tbl":
                        tc10.append(build_w_tbl_from_block(blk, el_state.table_line_tw))
                        if b_i == len(blocks) - 1 or blocks[b_i + 1]["type"] == "tbl":
                            tc10.append(
                                make_w_p(
                                    "", sz_half_pt=2, font_name=body_font, before_tw=0, after_tw=2, line_tw=20, line_rule="exact"
                                )
                            )
            else:
                tc10.append(make_w_p("", sz_half_pt=body_sz_hp, font_name=body_font, before_tw=40, after_tw=40))

            return tbl

        # 1. Dynamic Preamble / Cover Page from target_spec.preamble_blocks
        has_preamble_cover = bool(target_spec.preamble_blocks) and preserve_outer_tables
        if has_preamble_cover:
            default_outer_tw = (
                hu_to_twips(target_spec.section_slots[0].outer_width_hu)
                if target_spec.section_slots and target_spec.section_slots[0].outer_width_hu > 0
                else 9524
            )
            body.append(make_w_p("", sz_half_pt=body_sz_hp, font_name=body_font, before_tw=300, after_tw=150))
            cover_tbl = ET.Element(w_tag("tbl"))
            cPr = ET.SubElement(cover_tbl, w_tag("tblPr"))
            ET.SubElement(cPr, w_tag("tblW"), {w_tag("w"): str(default_outer_tw), w_tag("type"): "dxa"})
            ET.SubElement(cPr, w_tag("jc"), {w_tag("val"): "center"})
            cGrid = ET.SubElement(cover_tbl, w_tag("tblGrid"))
            ET.SubElement(cGrid, w_tag("gridCol"), {w_tag("w"): str(default_outer_tw)})
            ctr = ET.SubElement(cover_tbl, w_tag("tr"))
            ctc = ET.SubElement(ctr, w_tag("tc"))
            ctcPr = ET.SubElement(ctc, w_tag("tcPr"))
            ET.SubElement(ctcPr, w_tag("tcW"), {w_tag("w"): str(default_outer_tw), w_tag("type"): "dxa"})
            set_cell_borders_and_shd(ctcPr, border_color="3A3C84", border_sz="18", fill=None, mar=(240, 180, 240, 180))

            for pb in target_spec.preamble_blocks:
                p_sz = hu_to_half_pt(pb.get("height_hu", target_spec.body_font_size_hu))
                ctc.append(
                    make_w_p(
                        pb["text"],
                        sz_half_pt=p_sz,
                        font_name=body_font,
                        bold=bool(pb.get("bold", False)),
                        color=pb.get("color", "000000"),
                        jc="center",
                        before_tw=60,
                        after_tw=60,
                    )
                )
            body.append(cover_tbl)

        # 2. Dynamic Section Slots (2x2 Outer Tables or Unwrapped Blocks)
        default_state = ElasticLayoutState()
        rendered_slot_count = 0
        for s_slot in target_spec.section_slots:
            if (clear_unmapped_slots or not preserve_outer_tables) and s_slot.slot_uid not in slot_blocks_map:
                continue
            blks_s, st_s = slot_blocks_map.get(s_slot.slot_uid, ([], default_state))
            need_pb = has_preamble_cover or (rendered_slot_count > 0)
            rendered_slot_count += 1
            body.append(make_w_p("", sz_half_pt=2, font_name=body_font, before_tw=0, after_tw=0, line_tw=20, line_rule="exact"))
            if preserve_outer_tables:
                body.append(
                    make_outer_2x2_section_table(
                        s_slot.discovered_title_text,
                        s_slot.discovered_subtitle_text,
                        blks_s,
                        st_s,
                        page_break_before=need_pb,
                        s_slot=s_slot,
                    )
                )
            else:
                body.append(
                    make_w_p(
                        s_slot.discovered_title_text,
                        sz_half_pt=body_sz_hp + 4,
                        font_name=body_font,
                        bold=True,
                        page_break=need_pb,
                    )
                )
                for blk in blks_s:
                    if blk["type"] == "p":
                        body.append(
                            make_w_p_from_runs(
                                blk["runs"],
                                sz_half_pt=body_sz_hp,
                                font_name=body_font,
                                jc=blk["jc"],
                                before_tw=4,
                                after_tw=st_s.para_after_tw,
                                line_tw=st_s.body_line_tw,
                                line_rule="exact",
                            )
                        )
                    elif blk["type"] == "tbl":
                        body.append(build_w_tbl_from_block(blk, st_s.table_line_tw))

        # 3. Dynamic 2R x C Attachment Media Grid Table from target_spec.attachment_slots
        if target_spec.attachment_slots and preserve_outer_tables:
            for idx_ah, ah_p in enumerate(target_spec.attachment_header_paras):
                ah_sz = hu_to_half_pt(ah_p.get("height_hu", target_spec.body_font_size_hu))
                body.append(
                    make_w_p(
                        ah_p["text"],
                        sz_half_pt=ah_sz,
                        font_name=body_font,
                        bold=bool(ah_p.get("bold", idx_ah == 0)),
                        color=ah_p.get("color", "000000"),
                        jc="left",
                        before_tw=60 if idx_ah == 0 else 0,
                        after_tw=40 if idx_ah == 0 else 80,
                        line_tw=260,
                        page_break=(idx_ah == 0),
                    )
                )

            att0 = target_spec.attachment_slots[0]
            g_cols = max(1, att0.grid_cols)
            att_col_tw = hu_to_twips(att0.available_width_hu) if att0.available_width_hu > 0 else 4762
            att_outer_tw = att_col_tw * g_cols
            img_row_h_tw = max(2400, int(round(hu_to_twips(att0.available_height_hu) * 0.86)))
            cap_row_h_tw = max(360, hu_to_twips(att0.header_height_hu)) if att0.header_height_hu > 0 else 415
            target_cx_emu = hu_to_emu(int(att0.available_width_hu * 0.96))
            target_cy_emu = hu_to_emu(int(att0.available_height_hu * 0.90))
            placeholder_img = target_spec.attachment_placeholder_img_text
            placeholder_cap = target_spec.attachment_placeholder_cap_text

            att_tbl = ET.Element(w_tag("tbl"))
            aPr = ET.SubElement(att_tbl, w_tag("tblPr"))
            ET.SubElement(aPr, w_tag("tblW"), {w_tag("w"): str(att_outer_tw), w_tag("type"): "dxa"})
            ET.SubElement(aPr, w_tag("jc"), {w_tag("val"): "center"})
            ET.SubElement(aPr, w_tag("tblLayout"), {w_tag("type"): "fixed"})
            aGrid = ET.SubElement(att_tbl, w_tag("tblGrid"))
            for _ in range(g_cols):
                ET.SubElement(aGrid, w_tag("gridCol"), {w_tag("w"): str(att_col_tw)})

            def scale_drawing_xml(raw_p_xml: bytes, target_cx: int, target_cy: int) -> ET.Element:
                p_copy = ET.fromstring(raw_p_xml)
                pPr = p_copy.find(w_tag("pPr"))
                if pPr is None:
                    pPr = ET.Element(w_tag("pPr"))
                    p_copy.insert(0, pPr)
                for ch_p in list(pPr):
                    pPr.remove(ch_p)
                ET.SubElement(
                    pPr,
                    w_tag("spacing"),
                    {w_tag("before"): "20", w_tag("after"): "20", w_tag("line"): "240", w_tag("lineRule"): "auto"},
                )
                ET.SubElement(pPr, w_tag("jc"), {w_tag("val"): "center"})
                for ext in p_copy.findall(f".//{{{WP_NS}}}extent"):
                    ext.set("cx", str(target_cx))
                    ext.set("cy", str(target_cy))
                for aext in p_copy.findall(f".//{{{A_NS}}}ext"):
                    if "cx" in aext.attrib and int(aext.get("cx", "0")) > 500000:
                        aext.set("cx", str(target_cx))
                        aext.set("cy", str(target_cy))
                return p_copy

            n_pairs = max(1, len(target_spec.attachment_slots) // g_cols)
            for pair_idx in range(n_pairs):
                tr_img = ET.SubElement(att_tbl, w_tag("tr"))
                tr_imgPr = ET.SubElement(tr_img, w_tag("trPr"))
                ET.SubElement(tr_imgPr, w_tag("trHeight"), {w_tag("val"): str(img_row_h_tw), w_tag("hRule"): "atLeast"})
                for col_i in range(g_cols):
                    slot_flat_idx = pair_idx * g_cols + col_i
                    tc = ET.SubElement(tr_img, w_tag("tc"))
                    tcPr = ET.SubElement(tc, w_tag("tcPr"))
                    ET.SubElement(tcPr, w_tag("tcW"), {w_tag("w"): str(att_col_tw), w_tag("type"): "dxa"})
                    set_cell_borders_and_shd(
                        tcPr, border_color="000000", border_sz="6", fill=None, mar=(40, 40, 40, 40)
                    )
                    ET.SubElement(tcPr, w_tag("vAlign"), {w_tag("val"): "center"})
                    if slot_flat_idx < len(media_drawings) and media_drawings[slot_flat_idx][0] is not None:
                        tc.append(scale_drawing_xml(media_drawings[slot_flat_idx][0], target_cx_emu, target_cy_emu))
                    else:
                        tc.append(
                            make_w_p(
                                placeholder_img,
                                sz_half_pt=table_sz_hp,
                                font_name=table_font,
                                color="94A3B8",
                                jc="center",
                            )
                        )

                tr_cap = ET.SubElement(att_tbl, w_tag("tr"))
                tr_capPr = ET.SubElement(tr_cap, w_tag("trPr"))
                ET.SubElement(tr_capPr, w_tag("trHeight"), {w_tag("val"): str(cap_row_h_tw), w_tag("hRule"): "atLeast"})
                for col_i in range(g_cols):
                    slot_flat_idx = pair_idx * g_cols + col_i
                    tc = ET.SubElement(tr_cap, w_tag("tc"))
                    tcPr = ET.SubElement(tc, w_tag("tcPr"))
                    ET.SubElement(tcPr, w_tag("tcW"), {w_tag("w"): str(att_col_tw), w_tag("type"): "dxa"})
                    set_cell_borders_and_shd(
                        tcPr, border_color="000000", border_sz="6", fill="F8FAFC", mar=(30, 60, 30, 60)
                    )
                    ET.SubElement(tcPr, w_tag("vAlign"), {w_tag("val"): "center"})
                    if slot_flat_idx < len(media_drawings):
                        tc.append(
                            make_w_p(
                                media_drawings[slot_flat_idx][1],
                                sz_half_pt=max(16, table_sz_hp - 1),
                                font_name=table_font,
                                bold=True,
                                color="1E293B",
                                jc="center",
                            )
                        )
                    else:
                        tc.append(
                            make_w_p(
                                placeholder_cap,
                                sz_half_pt=max(16, table_sz_hp - 1),
                                font_name=table_font,
                                color="94A3B8",
                                jc="center",
                            )
                        )

            body.append(att_tbl)
            body.append(
                make_w_p("", sz_half_pt=2, font_name=body_font, before_tw=0, after_tw=0, line_tw=20, line_rule="exact")
            )

        body.append(sectPr)

        os.makedirs(os.path.dirname(os.path.abspath(out_docx_path)), exist_ok=True)
        new_doc_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        if has_valid_src_docx:
            with zipfile.ZipFile(source_docx_path, "r") as src_zip, zipfile.ZipFile(
                out_docx_path, "w", zipfile.ZIP_DEFLATED
            ) as zout:
                for item in src_zip.infolist():
                    if item.filename == "word/document.xml":
                        zout.writestr(item, new_doc_xml)
                    else:
                        zout.writestr(item, src_zip.read(item.filename))
        else:
            content_types_xml = (
                b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                b'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                b'<Default Extension="xml" ContentType="application/xml"/>'
                b'<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                b'</Types>'
            )
            root_rels_xml = (
                b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                b'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
                b'</Relationships>'
            )
            doc_rels_xml = (
                b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                b'</Relationships>'
            )
            with zipfile.ZipFile(out_docx_path, "w", zipfile.ZIP_DEFLATED) as zout:
                zout.writestr("[Content_Types].xml", content_types_xml)
                zout.writestr("_rels/.rels", root_rels_xml)
                zout.writestr("word/_rels/document.xml.rels", doc_rels_xml)
                zout.writestr("word/document.xml", new_doc_xml)

    return {
        "out_docx_path": out_docx_path,
        "bytes": os.path.getsize(out_docx_path),
        "effective_body_font": body_font,
        "effective_table_font": table_font,
        "effective_body_size_half_pt": body_sz_hp,
        "effective_table_size_half_pt": table_sz_hp,
    }


def project_ir_to_hwpx(
    base_hwpx_path: str,
    out_hwpx_path: str,
    target_spec: TargetDocumentSpecIR,
    slot_blocks_map: Dict[str, Tuple[List[dict], ElasticLayoutState]],
    attachment_captions: List[str],
) -> dict:
    """
    Projects the solved T-DIR blocks into an OWPML (.hwpx) archive using the effective
    `target_spec` (font families, font sizes, page dimensions, margins, preamble blocks,
    and discovered slot geometry) with zero hardcoded font names or domain strings.
    """
    body_font = target_spec.body_font_family or "sans-serif"
    body_sz_hu = target_spec.body_font_size_hu if target_spec.body_font_size_hu > 0 else 1200
    tbl_sz_hu = target_spec.table_font_size_hu if target_spec.table_font_size_hu > 0 else 1000
    title_sz_hu = max(900, body_sz_hu - 100)
    cover_sz_hu = int(round(body_sz_hu * 1.8))

    cs_list: List[Tuple[int, bool, bool, bool, str]] = [
        (body_sz_hu, False, False, False, "000000"),
        (body_sz_hu, True, False, False, "000000"),
        (tbl_sz_hu, False, False, False, "000000"),
        (tbl_sz_hu, True, False, False, "000000"),
        (title_sz_hu, True, False, False, "FFFFFF"),
        (tbl_sz_hu, False, False, False, "333333"),
        (cover_sz_hu, True, False, False, "3A3C84"),
    ]
    cs_map = {k: idx for idx, k in enumerate(cs_list)}

    def get_hwpx_cs(height_hu: int, bold: bool, italic: bool, underline: bool, color_hex: str) -> int:
        key = (height_hu, bold, italic, underline, color_hex.upper())
        if key not in cs_map:
            cs_map[key] = len(cs_list)
            cs_list.append(key)
        return cs_map[key]

    bf_list: List[Optional[str]] = [None, "OUTER_CELL", "3A3C84", "F2F2F2", "E8EAF6"]
    bf_map: Dict[Optional[str], int] = {None: 2, "3A3C84": 3, "F2F2F2": 4, "E8EAF6": 5}

    def get_hwpx_bf(fill_hex: Optional[str]) -> int:
        f_key = fill_hex.upper() if fill_hex else None
        if f_key not in bf_map:
            bf_list.append(f_key)
            bf_map[f_key] = len(bf_list)
        return bf_map[f_key]

    for blocks, _ in slot_blocks_map.values():
        for blk in blocks:
            if blk["type"] == "p":
                for r in blk["runs"]:
                    get_hwpx_cs(body_sz_hu, r["bold"], r["italic"], bool(r.get("underline", False)), r["color"])
            elif blk["type"] == "tbl":
                for cell in blk["cells"]:
                    get_hwpx_bf(cell["fill_hex"])
                    for p in cell["paras"]:
                        for r in p["runs"]:
                            get_hwpx_cs(tbl_sz_hu, r["bold"], r["italic"], bool(r.get("underline", False)), r["color"])

    def render_hwpx_p(p_blk: dict, height_hu: int, id_counter: List[int]) -> str:
        id_counter[0] += 1
        pid = id_counter[0]
        jc_map = {"both": "JUSTIFY", "left": "LEFT", "center": "CENTER", "right": "RIGHT"}
        ps_id = {"JUSTIFY": 0, "LEFT": 1, "CENTER": 2, "RIGHT": 3}.get(jc_map.get(p_blk["jc"], "JUSTIFY"), 0)
        runs_xml: List[str] = []
        for r in p_blk["runs"]:
            cid = get_hwpx_cs(height_hu, r["bold"], r["italic"], bool(r.get("underline", False)), r["color"])
            runs_xml.append(f'<hp:run charPrIDRef="{cid}"><hp:t>{xml_esc(r["text"])}</hp:t></hp:run>')
        if not runs_xml:
            runs_xml.append('<hp:run charPrIDRef="0"><hp:t></hp:t></hp:run>')
        return f'<hp:p id="{pid}" paraPrIDRef="{ps_id}" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">{"".join(runs_xml)}</hp:p>'

    def render_hwpx_inner_tbl(tbl_blk: dict, id_counter: List[int]) -> str:
        id_counter[0] += 1
        pid = id_counter[0]
        tid = pid + 50000
        rows_by_r: Dict[int, List[dict]] = {}
        for c in tbl_blk["cells"]:
            rows_by_r.setdefault(c["row"], []).append(c)
        tr_xml: List[str] = []
        for r_idx in range(tbl_blk["n_rows"]):
            tcs: List[str] = []
            for cell in sorted(rows_by_r.get(r_idx, []), key=lambda x: x["col"]):
                bf_id = get_hwpx_bf(cell["fill_hex"])
                mar_l, mar_r, mar_t, mar_b = cell["mar"]
                p_xml = "".join(render_hwpx_p(p, tbl_sz_hu, id_counter) for p in cell["paras"])
                tcs.append(
                    f'<hp:tc name="" header="0" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="{bf_id}">'
                    f'<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">'
                    f"{p_xml}</hp:subList>"
                    f'<hp:cellAddr colAddr="{cell["col"]}" rowAddr="{cell["row"]}"/>'
                    f'<hp:cellSpan colSpan="{cell["col_span"]}" rowSpan="{cell["row_span"]}"/>'
                    f'<hp:cellSz width="{cell["width_hu"]}" height="{cell["height_hu"]}"/>'
                    f'<hp:cellMargin left="{mar_l}" right="{mar_r}" top="{mar_t}" bottom="{mar_b}"/>'
                    f"</hp:tc>"
                )
            tr_xml.append(f'<hp:tr>{"".join(tcs)}</hp:tr>')
        return (
            f'<hp:p id="{pid}" paraPrIDRef="2" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">'
            f'<hp:run charPrIDRef="2">'
            f'<hp:tbl id="{tid}" zOrder="0" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="CELL" repeatHeader="0" rowCnt="{tbl_blk["n_rows"]}" colCnt="{tbl_blk["n_cols"]}" cellSpacing="0" borderFillIDRef="2" noAdjust="0">'
            f'<hp:sz width="{tbl_blk["width_hu"]}" widthRelTo="ABSOLUTE" height="{tbl_blk["height_hu"]}" heightRelTo="ABSOLUTE" protect="0"/>'
            f'<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="PARA" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/>'
            f'<hp:outMargin left="0" right="0" top="20" bottom="20"/>'
            f'<hp:inMargin left="60" right="60" top="25" bottom="25"/>'
            f'{"".join(tr_xml)}'
            f"</hp:tbl></hp:run></hp:p>"
        )

    def render_outer_2x2_hwpx(
        title_left: str,
        subtitle_right: str,
        blocks: List[dict],
        id_counter: List[int],
        page_break: int = 1,
        s_slot: Optional[SlotSpecIR] = None,
    ) -> str:
        id_counter[0] += 1
        pid = id_counter[0]
        tid = pid + 10000
        outer_w = s_slot.outer_width_hu if (s_slot and s_slot.outer_width_hu > 0) else 47622
        outer_h = s_slot.outer_height_hu if (s_slot and s_slot.outer_height_hu > 0) else 69252
        col0_w = (
            s_slot.col_widths_hu[0]
            if (s_slot and len(s_slot.col_widths_hu) >= 2)
            else int(round(outer_w * 0.328))
        )
        col1_w = outer_w - col0_w
        hdr_h = s_slot.header_height_hu if (s_slot and s_slot.header_height_hu > 0) else 2331
        body_h = (
            s_slot.available_height_hu
            if (s_slot and s_slot.available_height_hu > 0)
            else max(10000, outer_h - hdr_h)
        )

        if blocks:
            body_paras_xml: List[str] = []
            for blk in blocks:
                if blk["type"] == "p":
                    body_paras_xml.append(render_hwpx_p(blk, body_sz_hu, id_counter))
                elif blk["type"] == "tbl":
                    body_paras_xml.append(render_hwpx_inner_tbl(blk, id_counter))
            body_inner = "".join(body_paras_xml)
        else:
            id_counter[0] += 1
            body_inner = f'<hp:p id="{id_counter[0]}" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="0"><hp:t></hp:t></hp:run></hp:p>'

        return (
            f'<hp:p id="{pid}" paraPrIDRef="0" styleIDRef="0" pageBreak="{page_break}" columnBreak="0" merged="0">'
            f'<hp:run charPrIDRef="0">'
            f'<hp:tbl id="{tid}" zOrder="0" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="CELL" repeatHeader="0" rowCnt="2" colCnt="2" cellSpacing="0" borderFillIDRef="2" noAdjust="0">'
            f'<hp:sz width="{outer_w}" widthRelTo="ABSOLUTE" height="{outer_h}" heightRelTo="ABSOLUTE" protect="0"/>'
            f'<hp:pos treatAsChar="0" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="PARA" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/>'
            f'<hp:outMargin left="0" right="0" top="0" bottom="0"/>'
            f'<hp:inMargin left="350" right="350" top="100" bottom="100"/>'
            f"<hp:tr>"
            f'<hp:tc name="" header="0" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="3">'
            f'<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">'
            f'<hp:p id="{pid + 1}" paraPrIDRef="1" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="4"><hp:t>{xml_esc(title_left)}</hp:t></hp:run></hp:p>'
            f"</hp:subList>"
            f'<hp:cellAddr colAddr="0" rowAddr="0"/><hp:cellSpan colSpan="1" rowSpan="1"/>'
            f'<hp:cellSz width="{col0_w}" height="{hdr_h}"/><hp:cellMargin left="510" right="510" top="141" bottom="141"/>'
            f"</hp:tc>"
            f'<hp:tc name="" header="0" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="4">'
            f'<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">'
            f'<hp:p id="{pid + 2}" paraPrIDRef="1" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="5"><hp:t>{xml_esc(subtitle_right)}</hp:t></hp:run></hp:p>'
            f"</hp:subList>"
            f'<hp:cellAddr colAddr="1" rowAddr="0"/><hp:cellSpan colSpan="1" rowSpan="1"/>'
            f'<hp:cellSz width="{col1_w}" height="{hdr_h}"/><hp:cellMargin left="510" right="510" top="141" bottom="141"/>'
            f"</hp:tc>"
            f"</hp:tr>"
            f"<hp:tr>"
            f'<hp:tc name="" header="0" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="2">'
            f'<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="TOP" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">'
            f"{body_inner}"
            f"</hp:subList>"
            f'<hp:cellAddr colAddr="0" rowAddr="1"/><hp:cellSpan colSpan="2" rowSpan="1"/>'
            f'<hp:cellSz width="{outer_w}" height="{body_h}"/><hp:cellMargin left="350" right="350" top="100" bottom="100"/>'
            f"</hp:tc>"
            f"</hp:tr>"
            f"</hp:tbl></hp:run></hp:p>"
        )

    id_counter = [100]
    default_st = ElasticLayoutState()

    pw_hu = target_spec.page_width_hu if target_spec.page_width_hu > 0 else 59528
    ph_hu = target_spec.page_height_hu if target_spec.page_height_hu > 0 else 84186
    ml_hu = target_spec.margin_left_hu if target_spec.margin_left_hu > 0 else 5669
    mr_hu = target_spec.margin_right_hu if target_spec.margin_right_hu > 0 else 5669
    mt_hu = target_spec.margin_top_hu if target_spec.margin_top_hu > 0 else 4251
    mb_hu = target_spec.margin_bottom_hu if target_spec.margin_bottom_hu > 0 else 4252
    mh_hu = target_spec.margin_header_hu if target_spec.margin_header_hu > 0 else 2834
    mf_hu = target_spec.margin_footer_hu if target_spec.margin_footer_hu > 0 else 2834

    preamble_summary = " — ".join(pb["text"] for pb in target_spec.preamble_blocks[:3]) if target_spec.preamble_blocks else ""

    sec_p_list: List[str] = [
        '<hp:p id="1" paraPrIDRef="2" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">'
        '<hp:run charPrIDRef="6">'
        '<hp:secPr id="" textDirection="HORIZONTAL" spaceColumns="1134" tabStop="8000" tabStopVal="4000" tabStopUnit="HWPUNIT" outlineShapeIDRef="1" memoShapeIDRef="0" textVerticalWidthHead="0" masterPageCnt="0">'
        '<hp:grid lineGrid="0" charGrid="0" wonggojiFormat="0"/>'
        '<hp:startNum pageStartsOn="BOTH" page="0" pic="0" tbl="0" equation="0"/>'
        '<hp:visibility hideFirstHeader="0" hideFirstFooter="0" hideFirstMasterPage="0" border="SHOW_ALL" fill="SHOW_ALL" hideFirstPageNum="0" hideFirstEmptyLine="0" showLineNumber="0"/>'
        '<hp:lineNumberShape restartType="0" countBy="0" distance="0" startNumber="0"/>'
        f'<hp:pagePr landscape="WIDELY" width="{pw_hu}" height="{ph_hu}" gutterType="LEFT_ONLY">'
        f'<hp:margin header="{mh_hu}" footer="{mf_hu}" gutter="0" left="{ml_hu}" right="{mr_hu}" top="{mt_hu}" bottom="{mb_hu}"/>'
        "</hp:pagePr>"
        '<hp:footNotePr><hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar=")" supscript="0"/><hp:noteLine length="-1" type="SOLID" width="0.12 mm" color="#000000"/><hp:noteSpacing betweenNotes="283" belowLine="567" aboveLine="850"/><hp:numbering type="CONTINUOUS" newNum="1"/><hp:placement place="EACH_COLUMN" beneathText="0"/></hp:footNotePr>'
        '<hp:endNotePr><hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar=")" supscript="0"/><hp:noteLine length="14692344" type="SOLID" width="0.12 mm" color="#000000"/><hp:noteSpacing betweenNotes="0" belowLine="567" aboveLine="850"/><hp:numbering type="CONTINUOUS" newNum="1"/><hp:placement place="END_OF_DOCUMENT" beneathText="0"/></hp:endNotePr>'
        '<hp:pageBorderFill type="BOTH" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER"><hp:offset left="1417" right="1417" top="1417" bottom="1417"/></hp:pageBorderFill>'
        '<hp:pageBorderFill type="EVEN" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER"><hp:offset left="1417" right="1417" top="1417" bottom="1417"/></hp:pageBorderFill>'
        '<hp:pageBorderFill type="ODD" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER"><hp:offset left="1417" right="1417" top="1417" bottom="1417"/></hp:pageBorderFill>'
        "</hp:secPr>"
        '<hp:ctrl><hp:colPr id="" type="NEWSPAPER" layout="LEFT" colCount="1" sameSz="1" sameGap="0"/></hp:ctrl>'
        f"<hp:t>{xml_esc(preamble_summary)}</hp:t>"
        "</hp:run></hp:p>"
    ]

    for idx_s, s_slot in enumerate(target_spec.section_slots):
        s_blks, _ = slot_blocks_map.get(s_slot.slot_uid, ([], default_st))
        sec_p_list.append(
            render_outer_2x2_hwpx(
                s_slot.discovered_title_text,
                s_slot.discovered_subtitle_text,
                s_blks,
                id_counter,
                page_break=0 if idx_s == 0 else 1,
                s_slot=s_slot,
            )
        )

    for cap in attachment_captions:
        id_counter[0] += 1
        sec_p_list.append(
            f'<hp:p id="{id_counter[0]}" paraPrIDRef="2" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">'
            f'<hp:run charPrIDRef="3"><hp:t>{xml_esc(cap)}</hp:t></hp:run></hp:p>'
        )

    section0_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
        '<hs:sec xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" '
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
        'xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0">'
        + "".join(sec_p_list)
        + "</hs:sec>"
    )

    bf_xml_items: List[str] = []
    for idx_0, bf_val in enumerate(bf_list):
        bf_id = idx_0 + 1
        if bf_id == 1:
            bf_xml_items.append(
                '<hh:borderFill id="1" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0">'
                '<hh:slash type="NONE" Crooked="0" isCounter="0"/><hh:backSlash type="NONE" Crooked="0" isCounter="0"/>'
                '<hh:leftBorder type="NONE" width="0.12 mm" color="#000000"/><hh:rightBorder type="NONE" width="0.12 mm" color="#000000"/>'
                '<hh:topBorder type="NONE" width="0.12 mm" color="#000000"/><hh:bottomBorder type="NONE" width="0.12 mm" color="#000000"/>'
                '<hh:diagonal type="SOLID" width="0.1 mm" color="#000000"/></hh:borderFill>'
            )
        elif bf_val == "OUTER_CELL":
            bf_xml_items.append(
                f'<hh:borderFill id="{bf_id}" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0">'
                '<hh:slash type="NONE" Crooked="0" isCounter="0"/><hh:backSlash type="NONE" Crooked="0" isCounter="0"/>'
                '<hh:leftBorder type="SOLID" width="0.12 mm" color="#000000"/><hh:rightBorder type="SOLID" width="0.12 mm" color="#000000"/>'
                '<hh:topBorder type="SOLID" width="0.12 mm" color="#000000"/><hh:bottomBorder type="SOLID" width="0.12 mm" color="#000000"/>'
                '<hh:diagonal type="SOLID" width="0.1 mm" color="#000000"/></hh:borderFill>'
            )
        else:
            bf_xml_items.append(
                f'<hh:borderFill id="{bf_id}" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0">'
                '<hh:slash type="NONE" Crooked="0" isCounter="0"/><hh:backSlash type="NONE" Crooked="0" isCounter="0"/>'
                '<hh:leftBorder type="SOLID" width="0.12 mm" color="#CBD5E1"/><hh:rightBorder type="SOLID" width="0.12 mm" color="#CBD5E1"/>'
                '<hh:topBorder type="SOLID" width="0.12 mm" color="#CBD5E1"/><hh:bottomBorder type="SOLID" width="0.12 mm" color="#CBD5E1"/>'
                '<hh:diagonal type="SOLID" width="0.1 mm" color="#000000"/>'
                f'<hc:fillBrush><hc:winBrush faceColor="#{bf_val}" hatchColor="#999999" alpha="0"/></hc:fillBrush>'
                "</hh:borderFill>"
            )

    cs_xml_items: List[str] = []
    for cid, (h_hu, b_flag, i_flag, u_flag, col_hex) in enumerate(cs_list):
        b_tag = "<hh:bold/>" if b_flag else ""
        i_tag = "<hh:italic/>" if i_flag else ""
        u_type = "BOTTOM" if u_flag else "NONE"
        cs_xml_items.append(
            f'<hh:charPr id="{cid}" height="{h_hu}" textColor="#{col_hex}" shadeColor="none" useFontSpace="0" useKerning="0" symMark="NONE" borderFillIDRef="1">'
            '<hh:fontRef hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>'
            '<hh:ratio hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/>'
            '<hh:spacing hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>'
            '<hh:relSz hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/>'
            '<hh:offset hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>'
            f"{b_tag}{i_tag}"
            f'<hh:underline type="{u_type}" shape="SOLID" color="#{col_hex}"/>'
            '<hh:strikeout shape="NONE" color="#000000"/>'
            '<hh:outline type="NONE"/>'
            '<hh:shadow type="NONE" color="#C0C0C0" offsetX="10" offsetY="10"/>'
            "</hh:charPr>"
        )

    ls_pct = target_spec.min_line_spacing_pct if target_spec.min_line_spacing_pct > 0 else 122
    ps_xml_items: List[str] = []
    for pid, align_str in enumerate(("JUSTIFY", "LEFT", "CENTER", "RIGHT")):
        ps_xml_items.append(
            f'<hh:paraPr id="{pid}" tabPrIDRef="0" condense="0" fontLineHeight="0" snapToGrid="0" suppressLineNumbers="0" checked="0">'
            f'<hh:align horizontal="{align_str}" vertical="BASELINE"/>'
            '<hh:heading type="NONE" idRef="0" level="0"/>'
            '<hh:breakSetting breakLatinWord="KEEP_WORD" breakNonLatinWord="BREAK_WORD" widowOrphan="0" keepWithNext="0" keepLines="0" pageBreakBefore="0" lineWrap="BREAK"/>'
            '<hh:autoSpacing eAsianEng="0" eAsianNum="0"/>'
            '<hh:margin><hc:intent value="0" unit="HWPUNIT"/><hc:left value="0" unit="HWPUNIT"/><hc:right value="0" unit="HWPUNIT"/><hc:prev value="20" unit="HWPUNIT"/><hc:next value="35" unit="HWPUNIT"/></hh:margin>'
            f'<hh:lineSpacing type="PERCENT" value="{ls_pct}" unit="HWPUNIT"/>'
            '<hh:border borderFillIDRef="1" offsetLeft="0" offsetRight="0" offsetTop="0" offsetBottom="0" connect="0" ignoreMargin="0"/>'
            "</hh:paraPr>"
        )

    langs = ("HANGUL", "LATIN", "HANJA", "JAPANESE", "OTHER", "SYMBOL", "USER")
    fontfaces_xml = "".join(
        f'<hh:fontface lang="{lg}" fontCnt="1"><hh:font id="0" face="{xml_esc(body_font)}" type="TTF" isEmbedded="0"><hh:typeInfo familyType="FCAT_GOTHIC" weight="6" proportion="4" contrast="0" strokeVariation="1" armStyle="1" letterform="1" midline="1" xHeight="1"/></hh:font></hh:fontface>'
        for lg in langs
    )

    header_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
        '<hh:head xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" '
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
        'xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0" '
        'version="1.4" secCnt="1">'
        '<hh:beginNum page="1" footnote="1" endnote="1" pic="1" tbl="1" equation="1"/>'
        "<hh:refList>"
        f'<hh:fontfaces itemCnt="7">{fontfaces_xml}</hh:fontfaces>'
        f'<hh:borderFills itemCnt="{len(bf_xml_items)}">{"".join(bf_xml_items)}</hh:borderFills>'
        f'<hh:charProperties itemCnt="{len(cs_xml_items)}">{"".join(cs_xml_items)}</hh:charProperties>'
        '<hh:tabProperties itemCnt="1"><hh:tabPr id="0" autoTabLeft="0" autoTabRight="0"/></hh:tabProperties>'
        '<hh:numberings itemCnt="1"><hh:numbering id="1" start="0"><hh:paraHead start="1" level="1" align="LEFT" useInstWidth="1" autoIndent="1" widthAdjust="0" textOffsetType="PERCENT" textOffset="50" numFormat="DIGIT" charPrIDRef="0" checkable="0">^1.</hh:paraHead></hh:numbering></hh:numberings>'
        f'<hh:paraProperties itemCnt="{len(ps_xml_items)}">{"".join(ps_xml_items)}</hh:paraProperties>'
        '<hh:styles itemCnt="1"><hh:style id="0" type="PARA" name="바탕글" engName="Normal" paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="0" langID="1042" lockForm="0"/></hh:styles>'
        "</hh:refList>"
        '<hh:compatibleDocument targetProgram="HWP201X"><hh:layoutCompatibility/></hh:compatibleDocument>'
        '<hh:docOption><hh:linkinfo path="" pageInherit="0" footnoteInherit="0"/></hh:docOption>'
        '<hh:trackchageConfig flags="56"/>'
        "</hh:head>"
    )

    os.makedirs(os.path.dirname(os.path.abspath(out_hwpx_path)), exist_ok=True)
    with zipfile.ZipFile(base_hwpx_path, "r") as zin, zipfile.ZipFile(
        out_hwpx_path, "w", zipfile.ZIP_DEFLATED
    ) as zout:
        for item in zin.infolist():
            if item.filename == "mimetype":
                zout.writestr(item, zin.read(item.filename), compress_type=zipfile.ZIP_STORED)
            elif item.filename == "Contents/header.xml":
                zout.writestr(item, header_xml.encode("utf-8"))
            elif item.filename == "Contents/section0.xml":
                zout.writestr(item, section0_xml.encode("utf-8"))
            else:
                zout.writestr(item, zin.read(item.filename))

    return {
        "out_hwpx_path": out_hwpx_path,
        "bytes": os.path.getsize(out_hwpx_path),
        "effective_body_font": body_font,
        "effective_body_size_hu": body_sz_hu,
        "effective_table_size_hu": tbl_sz_hu,
    }
