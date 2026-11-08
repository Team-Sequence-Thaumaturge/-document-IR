"""
OOXML (.docx) and OWPML (.hwpx) Lifter, Dynamic Section/Slot Scanner, and Projector.
Supports both template-based transplantation and standalone format conversion.
"""

import os
import copy
import zipfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape as xml_esc
from typing import List, Dict, Tuple, Optional

from .ir_schema import SectionIR, MediaObjectIR, compute_non_ws_sha256, extract_ir_non_ws_text
from .elastic_solver import (
    ElasticLayoutState,
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


def extract_docx_p_data(p_el: ET.Element, is_in_table: bool) -> dict:
    pPr = p_el.find(w_tag("pPr"))
    jc = "both"
    before_tw, after_tw = (0, 0) if is_in_table else (0, 8)
    if pPr is not None:
        jc_el = pPr.find(w_tag("jc"))
        if jc_el is not None:
            jc = jc_el.get(w_tag("val"), "both")

    blip = p_el.find(f".//{{{A_NS}}}blip")
    r_embed = blip.get(f"{{{R_NS}}}embed") if blip is not None else None

    runs: List[dict] = []
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
        color_hex = "000000"
        if rPr is not None:
            b_el = rPr.find(w_tag("b"))
            if b_el is not None and b_el.get(w_tag("val"), "true") not in ("0", "false"):
                bold = True
            i_el = rPr.find(w_tag("i"))
            if i_el is not None and i_el.get(w_tag("val"), "true") not in ("0", "false"):
                italic = True
            c_el = rPr.find(w_tag("color"))
            if c_el is not None:
                c_val = c_el.get(w_tag("val"), "000000")
                if c_val and c_val != "auto":
                    color_hex = c_val.upper()
        runs.append(
            {
                "text": txt,
                "bold": bold,
                "italic": italic,
                "color": color_hex,
            }
        )

    return {
        "type": "p",
        "jc": jc,
        "left_tw": 0,
        "right_tw": 0,
        "indent_tw": 0,
        "before_tw": before_tw,
        "after_tw": after_tw,
        "runs": runs,
        "r_embed": r_embed,
        "raw_p_xml": ET.tostring(p_el, encoding="utf-8") if r_embed else None,
    }


def extract_docx_tbl_normalized(tbl_el: ET.Element, target_width_hu: int = 46200) -> dict:
    full_txt = "".join(tbl_el.itertext())
    grid = tbl_el.find(w_tag("tblGrid"))
    col_twips: List[int] = []
    if grid is not None:
        for gc in grid.findall(w_tag("gridCol")):
            col_twips.append(int(gc.get(w_tag("w"), "1000")))
    n_cols_raw = len(col_twips) or 1
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
                    for p_el in tc.findall(w_tag("p")):
                        pdata = extract_docx_p_data(p_el, is_in_table=True)
                        if pdata["runs"]:
                            master["paras"].append(pdata)
                    for dc in range(col_span):
                        grid_cells[r_idx][c_cursor + dc] = master
                    c_cursor += col_span
                    continue

            paras: List[dict] = []
            for p_el in tc.findall(w_tag("p")):
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

    if "WHAT?" in full_txt and n_rows_raw == 7 and n_cols_raw == 4:
        return collapse_stair_step_chart_table(grid_cells, target_width_hu)

    if "STEP 1 · 한국" in full_txt and n_rows_raw == 5 and n_cols_raw == 1:
        return fold_vertical_step_arrows(grid_cells, target_width_hu)

    if "HOW MUCH?" in full_txt and n_rows_raw == 4:
        new_cells: List[dict] = []
        for idx_r in range(3):
            c_nonempty = next(
                (grid_cells[idx_r][c] for c in range(n_cols_raw) if grid_cells[idx_r][c] and grid_cells[idx_r][c]["paras"]),
                grid_cells[idx_r][0],
            )
            cobj = copy.deepcopy(c_nonempty)
            cobj["row"], cobj["col"], cobj["row_span"], cobj["col_span"] = idx_r, 0, 1, 1
            cobj["width_hu"] = target_width_hu
            cobj["mar"] = (160, 160, 35, 35)
            cobj["borders"] = {s: (1, 1, (0x94, 0xA3, 0xB8)) for s in ("left", "right", "top", "bottom")}
            cobj["height_hu"] = max(800, len(cobj["paras"]) * 950 + 80)
            new_cells.append(cobj)
        return {
            "type": "tbl",
            "n_rows": 3,
            "n_cols": 1,
            "col_widths_hu": [target_width_hu],
            "row_cols": [1, 1, 1],
            "cells": new_cells,
            "width_hu": target_width_hu,
            "height_hu": sum(c["height_hu"] for c in new_cells),
        }

    return detect_and_remove_spacer_rows_cols(grid_cells, col_twips, n_rows_raw, n_cols_raw, target_width_hu)


def lift_docx_to_ir(docx_path: str, target_width_hu: int = 46200) -> Dict[str, SectionIR]:
    """
    Dynamically scans a DOCX file, detects top-level sections by heading patterns
    (e.g. 'Q1.', 'Q2.', 'Q3.', 'Q4.', '출처') or treats as a single 'MAIN' section,
    and lifts all paragraphs, tables, and embedded media into SectionIR objects.
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
        children = list(body) if body is not None else []

        cur_sec_id = "PREAMBLE"
        cur_heading = "Document Preamble"
        cur_blocks: List[dict] = []
        cur_media: List[MediaObjectIR] = []

        def flush_section():
            nonlocal cur_blocks, cur_media
            if not cur_blocks and not cur_media:
                return
            _, n_chars, sha = extract_ir_non_ws_text(cur_blocks)
            sections[cur_sec_id] = SectionIR(
                section_uid=cur_sec_id,
                heading_text=cur_heading,
                blocks=cur_blocks,
                media_objects=cur_media,
                non_ws_char_count=n_chars,
                text_sha256=sha,
            )
            cur_blocks = []
            cur_media = []

        i = 0
        while i < len(children):
            ch = children[i]
            tag = ch.tag.split("}")[-1]
            if tag == "p":
                pdata = extract_docx_p_data(ch, is_in_table=False)
                ptxt = "".join(r["text"] for r in pdata["runs"]).strip()

                # Dynamic section boundary detection
                new_sec_key = None
                if ptxt.startswith("Q1.") or ptxt.startswith("1. 문제"):
                    new_sec_key = "Q1"
                elif ptxt.startswith("Q2.") or ptxt.startswith("2. 실현") or ptxt.startswith("Q2 "):
                    new_sec_key = "Q2"
                elif ptxt.startswith("Q3.") or ptxt.startswith("3. 성장"):
                    new_sec_key = "Q3"
                elif ptxt.startswith("Q4.") or ptxt.startswith("4. 향후"):
                    new_sec_key = "Q4"
                elif ptxt == "출처" or ptxt.startswith("[출처]"):
                    new_sec_key = "REFERENCES"

                if new_sec_key is not None:
                    flush_section()
                    cur_sec_id = new_sec_key
                    cur_heading = ptxt
                    # Check if the immediately following paragraph is a section subtitle (e.g. '아이디어의 개발 동기...' or '문제 인식 (Problem) — ...')
                    if i + 1 < len(children) and children[i + 1].tag.split("}")[-1] == "p":
                        nxt_p = extract_docx_p_data(children[i + 1], is_in_table=False)
                        nxt_txt = "".join(r["text"] for r in nxt_p["runs"]).strip()
                        if "아이디어의 개발" in nxt_txt or "사업화 계획" in nxt_txt or "문제 인식" in nxt_txt or "해결 방안" in nxt_txt:
                            i += 2
                            continue
                    i += 1
                    continue

                if pdata["r_embed"] and pdata["r_embed"] in rels_map:
                    rel_target = rels_map[pdata["r_embed"]].lstrip("/")
                    zip_media_path = rel_target if rel_target.startswith("word/") else f"word/{rel_target}"
                    if zip_media_path in zf.namelist():
                        raw_img = zf.read(zip_media_path)
                        cur_media.append(
                            MediaObjectIR(
                                media_uid=f"{cur_sec_id}_IMG_{len(cur_media) + 1}",
                                rel_id=pdata["r_embed"],
                                filename=os.path.basename(zip_media_path),
                                mime_type="image/png",
                                raw_bytes=raw_img,
                                caption_hint=ptxt,
                            )
                        )
                if pdata["runs"] or pdata["r_embed"]:
                    cur_blocks.append(pdata)
            elif tag == "tbl":
                cur_blocks.append(extract_docx_tbl_normalized(ch, target_width_hu))
            i += 1

        flush_section()

    return sections


def project_ir_to_docx(
    source_docx_path: str,
    out_docx_path: str,
    slot_blocks_map: Dict[str, Tuple[List[dict], ElasticLayoutState]],
    media_drawings: List[Tuple[Optional[bytes], str]],
) -> dict:
    with zipfile.ZipFile(source_docx_path, "r") as src_zip:
        doc_xml = src_zip.read("word/document.xml")
        root = ET.fromstring(doc_xml)
        for k in list(root.attrib.keys()):
            if k.endswith("Ignorable"):
                del root.attrib[k]
        body = root.find(w_tag("body"))
        sectPr = copy.deepcopy(body.find(w_tag("sectPr")))
        pgMar = sectPr.find(w_tag("pgMar"))
        if pgMar is not None:
            pgMar.set(w_tag("top"), "560")
            pgMar.set(w_tag("bottom"), "560")
            pgMar.set(w_tag("left"), "950")
            pgMar.set(w_tag("right"), "950")
            pgMar.set(w_tag("header"), "360")
            pgMar.set(w_tag("footer"), "360")
        for dg in sectPr.findall(w_tag("docGrid")):
            sectPr.remove(dg)
        ET.SubElement(sectPr, w_tag("docGrid"), {w_tag("type"): "default"})
        body.clear()

        def make_w_p_from_runs(
            runs: List[dict],
            sz_half_pt: int = 24,
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
                        w_tag("ascii"): "맑은 고딕",
                        w_tag("hAnsi"): "맑은 고딕",
                        w_tag("eastAsia"): "맑은 고딕",
                        w_tag("cs"): "맑은 고딕",
                    },
                )
                if r_info.get("bold"):
                    ET.SubElement(rPr, w_tag("b"))
                if r_info.get("italic"):
                    ET.SubElement(rPr, w_tag("i"))
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
            sz_half_pt: int = 24,
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
            return make_w_p_from_runs(runs, sz_half_pt, jc, before_tw, after_tw, line_tw, page_break, line_rule)

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
            col_tws = [int(round(w / 5.0)) for w in tbl_blk["col_widths_hu"]]
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
                                sz_half_pt=20,
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
        ) -> ET.Element:
            tbl = ET.Element(w_tag("tbl"))
            tblPr = ET.SubElement(tbl, w_tag("tblPr"))
            ET.SubElement(tblPr, w_tag("tblW"), {w_tag("w"): "9524", w_tag("type"): "dxa"})
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
            ET.SubElement(grid, w_tag("gridCol"), {w_tag("w"): "3121"})
            ET.SubElement(grid, w_tag("gridCol"), {w_tag("w"): "6403"})

            tr0 = ET.SubElement(tbl, w_tag("tr"))
            tr0Pr = ET.SubElement(tr0, w_tag("trPr"))
            ET.SubElement(tr0Pr, w_tag("cantSplit"))
            ET.SubElement(tr0Pr, w_tag("trHeight"), {w_tag("val"): "380", w_tag("hRule"): "atLeast"})

            tc00 = ET.SubElement(tr0, w_tag("tc"))
            tc00Pr = ET.SubElement(tc00, w_tag("tcPr"))
            ET.SubElement(tc00Pr, w_tag("tcW"), {w_tag("w"): "3121", w_tag("type"): "dxa"})
            set_cell_borders_and_shd(tc00Pr, border_color="000000", border_sz="6", fill="3A3C84", mar=(30, 90, 30, 90))
            ET.SubElement(tc00Pr, w_tag("vAlign"), {w_tag("val"): "center"})
            tc00.append(
                make_w_p(
                    title_left,
                    sz_half_pt=22,
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
            ET.SubElement(tc01Pr, w_tag("tcW"), {w_tag("w"): "6403", w_tag("type"): "dxa"})
            set_cell_borders_and_shd(tc01Pr, border_color="000000", border_sz="6", fill="F2F2F2", mar=(30, 90, 30, 90))
            ET.SubElement(tc01Pr, w_tag("vAlign"), {w_tag("val"): "center"})
            tc01.append(
                make_w_p(
                    subtitle_right,
                    sz_half_pt=20,
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
                    tr1Pr, w_tag("trHeight"), {w_tag("val"): str(min_body_height_tw), w_tag("hRule"): "atLeast"}
                )

            tc10 = ET.SubElement(tr1, w_tag("tc"))
            tc10Pr = ET.SubElement(tc10, w_tag("tcPr"))
            ET.SubElement(tc10Pr, w_tag("tcW"), {w_tag("w"): "9524", w_tag("type"): "dxa"})
            ET.SubElement(tc10Pr, w_tag("gridSpan"), {w_tag("val"): "2"})
            set_cell_borders_and_shd(tc10Pr, border_color="000000", border_sz="6", fill=None, mar=(35, 95, 35, 95))

            if blocks:
                if blocks[0]["type"] == "tbl":
                    tc10.append(make_w_p("", sz_half_pt=2, before_tw=0, after_tw=0, line_tw=20, line_rule="exact"))
                for b_i, blk in enumerate(blocks):
                    if blk["type"] == "p":
                        tc10.append(
                            make_w_p_from_runs(
                                blk["runs"],
                                sz_half_pt=24,
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
                                make_w_p("", sz_half_pt=2, before_tw=0, after_tw=2, line_tw=20, line_rule="exact")
                            )
            else:
                tc10.append(make_w_p("", sz_half_pt=24, before_tw=40, after_tw=40))

            return tbl

        # 1. Cover Page
        body.append(make_w_p("", sz_half_pt=24, before_tw=400, after_tw=200))
        cover_tbl = ET.Element(w_tag("tbl"))
        cPr = ET.SubElement(cover_tbl, w_tag("tblPr"))
        ET.SubElement(cPr, w_tag("tblW"), {w_tag("w"): "9524", w_tag("type"): "dxa"})
        ET.SubElement(cPr, w_tag("jc"), {w_tag("val"): "center"})
        cGrid = ET.SubElement(cover_tbl, w_tag("tblGrid"))
        ET.SubElement(cGrid, w_tag("gridCol"), {w_tag("w"): "9524"})
        ctr = ET.SubElement(cover_tbl, w_tag("tr"))
        ctc = ET.SubElement(ctr, w_tag("tc"))
        ctcPr = ET.SubElement(ctc, w_tag("tcPr"))
        ET.SubElement(ctcPr, w_tag("tcW"), {w_tag("w"): "9524", w_tag("type"): "dxa"})
        set_cell_borders_and_shd(ctcPr, border_color="3A3C84", border_sz="18", fill=None, mar=(300, 200, 300, 200))
        ctc.append(
            make_w_p(
                "2026년 2차 모두의 창업 프로젝트",
                sz_half_pt=40,
                bold=True,
                color="1E293B",
                jc="center",
                before_tw=120,
                after_tw=80,
            )
        )
        ctc.append(
            make_w_p(
                "1R 사업계획서", sz_half_pt=52, bold=True, color="3A3C84", jc="center", before_tw=80, after_tw=120
            )
        )
        body.append(cover_tbl)
        body.append(
            make_w_p("2026. 0.", sz_half_pt=28, bold=True, color="333333", jc="center", before_tw=800, after_tw=800)
        )

        info_tbl = ET.Element(w_tag("tbl"))
        iPr = ET.SubElement(info_tbl, w_tag("tblPr"))
        ET.SubElement(iPr, w_tag("tblW"), {w_tag("w"): "8166", w_tag("type"): "dxa"})
        ET.SubElement(iPr, w_tag("jc"), {w_tag("val"): "center"})
        iGrid = ET.SubElement(info_tbl, w_tag("tblGrid"))
        for w_val in (1193, 2890, 1193, 2890):
            ET.SubElement(iGrid, w_tag("gridCol"), {w_tag("w"): str(w_val)})
        itr0 = ET.SubElement(info_tbl, w_tag("tr"))
        itc00 = ET.SubElement(itr0, w_tag("tc"))
        itc00Pr = ET.SubElement(itc00, w_tag("tcPr"))
        ET.SubElement(itc00Pr, w_tag("tcW"), {w_tag("w"): "1193", w_tag("type"): "dxa"})
        set_cell_borders_and_shd(itc00Pr, fill="E8EAF6", mar=(120, 100, 120, 100))
        itc00.append(make_w_p("멘토기관", sz_half_pt=22, bold=True, jc="center"))
        itc01 = ET.SubElement(itr0, w_tag("tc"))
        itc01Pr = ET.SubElement(itc01, w_tag("tcPr"))
        ET.SubElement(itc01Pr, w_tag("tcW"), {w_tag("w"): "6973", w_tag("type"): "dxa"})
        ET.SubElement(itc01Pr, w_tag("gridSpan"), {w_tag("val"): "3"})
        set_cell_borders_and_shd(itc01Pr, fill=None, mar=(120, 100, 120, 100))
        itc01.append(make_w_p("광운대학교", sz_half_pt=22, bold=False, jc="center"))

        itr1 = ET.SubElement(info_tbl, w_tag("tr"))
        for lbl, w_v, is_hdr in (("이름", 1193, True), ("", 2890, False), ("닉네임", 1193, True), ("", 2890, False)):
            tc = ET.SubElement(itr1, w_tag("tc"))
            tcPr = ET.SubElement(tc, w_tag("tcPr"))
            ET.SubElement(tcPr, w_tag("tcW"), {w_tag("w"): str(w_v), w_tag("type"): "dxa"})
            set_cell_borders_and_shd(tcPr, fill="E8EAF6" if is_hdr else None, mar=(120, 100, 120, 100))
            tc.append(make_w_p(lbl, sz_half_pt=22, bold=is_hdr, jc="center"))
        body.append(info_tbl)

        default_state = ElasticLayoutState()
        q1_blks, q1_st = slot_blocks_map.get("Q1_BODY", ([], default_state))
        q2_blks, q2_st = slot_blocks_map.get("Q2_BODY", ([], default_state))
        q3_blks, q3_st = slot_blocks_map.get("Q3_BODY", ([], default_state))
        q4_blks, q4_st = slot_blocks_map.get("Q4_BODY", ([], default_state))

        for title_l, sub_r, blks_s, st_s in (
            ("Q1. 문제 인식(problem)", " ※ 아이디어의 개발 동기·목적 및 목표 시장 등", q1_blks, q1_st),
            ("Q2. 실현 가능성(solution)", " ※ 아이디어의 개발 방안, 준비 정도, 차별화 방안", q2_blks, q2_st),
            ("Q3. 성장전략(scale-up)", " ※ 사업화 계획, 추진일정, 자금 소요 및 조달 계획", q3_blks, q3_st),
            ("Q4. 향후 발전방안 (development)", " ※ 2R 이후 사업 고도화, 후속 계획 등", q4_blks, q4_st),
        ):
            body.append(make_w_p("", sz_half_pt=2, before_tw=0, after_tw=0, line_tw=20, line_rule="exact"))
            body.append(make_outer_2x2_section_table(title_l, sub_r, blks_s, st_s, page_break_before=True))

        # Bottom '붙임' Attachment 6x2 Table
        body.append(
            make_w_p(
                "붙임 : 사업계획서 관련 참고자료 (※선택 제출 사항, 분량 A4 3p 이내)",
                sz_half_pt=24,
                bold=True,
                color="1E293B",
                jc="left",
                before_tw=60,
                after_tw=40,
                line_tw=280,
                page_break=True,
            )
        )
        body.append(
            make_w_p(
                "※ 사진 등 참고자료 첨부(해당 시) * 최대 3장까지 가능",
                sz_half_pt=20,
                bold=False,
                color="475569",
                jc="left",
                before_tw=0,
                after_tw=80,
                line_tw=240,
            )
        )

        att_tbl = ET.Element(w_tag("tbl"))
        aPr = ET.SubElement(att_tbl, w_tag("tblPr"))
        ET.SubElement(aPr, w_tag("tblW"), {w_tag("w"): "9524", w_tag("type"): "dxa"})
        ET.SubElement(aPr, w_tag("jc"), {w_tag("val"): "center"})
        ET.SubElement(aPr, w_tag("tblLayout"), {w_tag("type"): "fixed"})
        aGrid = ET.SubElement(att_tbl, w_tag("tblGrid"))
        ET.SubElement(aGrid, w_tag("gridCol"), {w_tag("w"): "4762"})
        ET.SubElement(aGrid, w_tag("gridCol"), {w_tag("w"): "4762"})

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

        for pair_idx in range(3):
            tr_img = ET.SubElement(att_tbl, w_tag("tr"))
            tr_imgPr = ET.SubElement(tr_img, w_tag("trPr"))
            ET.SubElement(tr_imgPr, w_tag("trHeight"), {w_tag("val"): "3100", w_tag("hRule"): "atLeast"})
            for col_i in range(2):
                slot_flat_idx = pair_idx * 2 + col_i
                tc = ET.SubElement(tr_img, w_tag("tc"))
                tcPr = ET.SubElement(tc, w_tag("tcPr"))
                ET.SubElement(tcPr, w_tag("tcW"), {w_tag("w"): "4762", w_tag("type"): "dxa"})
                set_cell_borders_and_shd(tcPr, border_color="000000", border_sz="6", fill=None, mar=(40, 40, 40, 40))
                ET.SubElement(tcPr, w_tag("vAlign"), {w_tag("val"): "center"})
                if slot_flat_idx < len(media_drawings) and media_drawings[slot_flat_idx][0] is not None:
                    tc.append(scale_drawing_xml(media_drawings[slot_flat_idx][0], 2900000, 1620000))
                else:
                    tc.append(make_w_p("사진 등", sz_half_pt=20, color="94A3B8", jc="center"))

            tr_cap = ET.SubElement(att_tbl, w_tag("tr"))
            tr_capPr = ET.SubElement(tr_cap, w_tag("trPr"))
            ET.SubElement(tr_capPr, w_tag("trHeight"), {w_tag("val"): "415", w_tag("hRule"): "atLeast"})
            for col_i in range(2):
                slot_flat_idx = pair_idx * 2 + col_i
                tc = ET.SubElement(tr_cap, w_tag("tc"))
                tcPr = ET.SubElement(tc, w_tag("tcPr"))
                ET.SubElement(tcPr, w_tag("tcW"), {w_tag("w"): "4762", w_tag("type"): "dxa"})
                set_cell_borders_and_shd(
                    tcPr, border_color="000000", border_sz="6", fill="F8FAFC", mar=(30, 60, 30, 60)
                )
                ET.SubElement(tcPr, w_tag("vAlign"), {w_tag("val"): "center"})
                if slot_flat_idx < len(media_drawings):
                    tc.append(
                        make_w_p(
                            media_drawings[slot_flat_idx][1],
                            sz_half_pt=19,
                            bold=True,
                            color="1E293B",
                            jc="center",
                        )
                    )
                else:
                    tc.append(make_w_p("사진 제목/설명 등", sz_half_pt=19, color="94A3B8", jc="center"))

        body.append(att_tbl)
        body.append(make_w_p("", sz_half_pt=2, before_tw=0, after_tw=0, line_tw=20, line_rule="exact"))
        body.append(sectPr)

        os.makedirs(os.path.dirname(os.path.abspath(out_docx_path)), exist_ok=True)
        new_doc_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        with zipfile.ZipFile(out_docx_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in src_zip.infolist():
                if item.filename == "word/document.xml":
                    zout.writestr(item, new_doc_xml)
                else:
                    zout.writestr(item, src_zip.read(item.filename))

    return {"out_docx_path": out_docx_path, "bytes": os.path.getsize(out_docx_path)}


def project_ir_to_hwpx(
    base_hwpx_path: str,
    out_hwpx_path: str,
    slot_blocks_map: Dict[str, Tuple[List[dict], ElasticLayoutState]],
    attachment_captions: List[str],
) -> dict:
    cs_list: List[Tuple[int, bool, bool, str]] = [
        (1200, False, False, "000000"),
        (1200, True, False, "000000"),
        (1000, False, False, "000000"),
        (1000, True, False, "000000"),
        (1100, True, False, "FFFFFF"),
        (1000, False, False, "333333"),
        (2200, True, False, "3A3C84"),
    ]
    cs_map = {k: idx for idx, k in enumerate(cs_list)}

    def get_hwpx_cs(height_hu: int, bold: bool, italic: bool, color_hex: str) -> int:
        key = (height_hu, bold, italic, color_hex.upper())
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
                    get_hwpx_cs(1200, r["bold"], r["italic"], r["color"])
            elif blk["type"] == "tbl":
                for cell in blk["cells"]:
                    get_hwpx_bf(cell["fill_hex"])
                    for p in cell["paras"]:
                        for r in p["runs"]:
                            get_hwpx_cs(1000, r["bold"], r["italic"], r["color"])

    def render_hwpx_p(p_blk: dict, height_hu: int, id_counter: List[int]) -> str:
        id_counter[0] += 1
        pid = id_counter[0]
        jc_map = {"both": "JUSTIFY", "left": "LEFT", "center": "CENTER", "right": "RIGHT"}
        ps_id = {"JUSTIFY": 0, "LEFT": 1, "CENTER": 2, "RIGHT": 3}.get(jc_map.get(p_blk["jc"], "JUSTIFY"), 0)
        runs_xml: List[str] = []
        for r in p_blk["runs"]:
            cid = get_hwpx_cs(height_hu, r["bold"], r["italic"], r["color"])
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
                p_xml = "".join(render_hwpx_p(p, 1000, id_counter) for p in cell["paras"])
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
        title_left: str, subtitle_right: str, blocks: List[dict], id_counter: List[int], page_break: int = 1
    ) -> str:
        id_counter[0] += 1
        pid = id_counter[0]
        tid = pid + 10000
        if blocks:
            body_paras_xml: List[str] = []
            for blk in blocks:
                if blk["type"] == "p":
                    body_paras_xml.append(render_hwpx_p(blk, 1200, id_counter))
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
            f'<hp:sz width="47622" widthRelTo="ABSOLUTE" height="69252" heightRelTo="ABSOLUTE" protect="0"/>'
            f'<hp:pos treatAsChar="0" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="PARA" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/>'
            f'<hp:outMargin left="0" right="0" top="0" bottom="0"/>'
            f'<hp:inMargin left="350" right="350" top="100" bottom="100"/>'
            f"<hp:tr>"
            f'<hp:tc name="" header="0" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="3">'
            f'<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">'
            f'<hp:p id="{pid + 1}" paraPrIDRef="1" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="4"><hp:t>{xml_esc(title_left)}</hp:t></hp:run></hp:p>'
            f"</hp:subList>"
            f'<hp:cellAddr colAddr="0" rowAddr="0"/><hp:cellSpan colSpan="1" rowSpan="1"/>'
            f'<hp:cellSz width="15604" height="2331"/><hp:cellMargin left="510" right="510" top="141" bottom="141"/>'
            f"</hp:tc>"
            f'<hp:tc name="" header="0" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="4">'
            f'<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">'
            f'<hp:p id="{pid + 2}" paraPrIDRef="1" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="5"><hp:t>{xml_esc(subtitle_right)}</hp:t></hp:run></hp:p>'
            f"</hp:subList>"
            f'<hp:cellAddr colAddr="1" rowAddr="0"/><hp:cellSpan colSpan="1" rowSpan="1"/>'
            f'<hp:cellSz width="32018" height="2331"/><hp:cellMargin left="510" right="510" top="141" bottom="141"/>'
            f"</hp:tc>"
            f"</hp:tr>"
            f"<hp:tr>"
            f'<hp:tc name="" header="0" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="2">'
            f'<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="TOP" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">'
            f"{body_inner}"
            f"</hp:subList>"
            f'<hp:cellAddr colAddr="0" rowAddr="1"/><hp:cellSpan colSpan="2" rowSpan="1"/>'
            f'<hp:cellSz width="47622" height="66921"/><hp:cellMargin left="350" right="350" top="100" bottom="100"/>'
            f"</hp:tc>"
            f"</hp:tr>"
            f"</hp:tbl></hp:run></hp:p>"
        )

    id_counter = [100]
    default_st = ElasticLayoutState()
    q1_blks, _ = slot_blocks_map.get("Q1_BODY", ([], default_st))
    q2_blks, _ = slot_blocks_map.get("Q2_BODY", ([], default_st))
    q3_blks, _ = slot_blocks_map.get("Q3_BODY", ([], default_st))
    q4_blks, _ = slot_blocks_map.get("Q4_BODY", ([], default_st))

    sec_p_list: List[str] = [
        '<hp:p id="1" paraPrIDRef="2" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">'
        '<hp:run charPrIDRef="6">'
        '<hp:secPr id="" textDirection="HORIZONTAL" spaceColumns="1134" tabStop="8000" tabStopVal="4000" tabStopUnit="HWPUNIT" outlineShapeIDRef="1" memoShapeIDRef="0" textVerticalWidthHead="0" masterPageCnt="0">'
        '<hp:grid lineGrid="0" charGrid="0" wonggojiFormat="0"/>'
        '<hp:startNum pageStartsOn="BOTH" page="0" pic="0" tbl="0" equation="0"/>'
        '<hp:visibility hideFirstHeader="0" hideFirstFooter="0" hideFirstMasterPage="0" border="SHOW_ALL" fill="SHOW_ALL" hideFirstPageNum="0" hideFirstEmptyLine="0" showLineNumber="0"/>'
        '<hp:lineNumberShape restartType="0" countBy="0" distance="0" startNumber="0"/>'
        '<hp:pagePr landscape="WIDELY" width="59528" height="84186" gutterType="LEFT_ONLY">'
        '<hp:margin header="2834" footer="2834" gutter="0" left="5669" right="5669" top="4251" bottom="4252"/>'
        "</hp:pagePr>"
        '<hp:footNotePr><hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar=")" supscript="0"/><hp:noteLine length="-1" type="SOLID" width="0.12 mm" color="#000000"/><hp:noteSpacing betweenNotes="283" belowLine="567" aboveLine="850"/><hp:numbering type="CONTINUOUS" newNum="1"/><hp:placement place="EACH_COLUMN" beneathText="0"/></hp:footNotePr>'
        '<hp:endNotePr><hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar=")" supscript="0"/><hp:noteLine length="14692344" type="SOLID" width="0.12 mm" color="#000000"/><hp:noteSpacing betweenNotes="0" belowLine="567" aboveLine="850"/><hp:numbering type="CONTINUOUS" newNum="1"/><hp:placement place="END_OF_DOCUMENT" beneathText="0"/></hp:endNotePr>'
        '<hp:pageBorderFill type="BOTH" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER"><hp:offset left="1417" right="1417" top="1417" bottom="1417"/></hp:pageBorderFill>'
        '<hp:pageBorderFill type="EVEN" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER"><hp:offset left="1417" right="1417" top="1417" bottom="1417"/></hp:pageBorderFill>'
        '<hp:pageBorderFill type="ODD" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER"><hp:offset left="1417" right="1417" top="1417" bottom="1417"/></hp:pageBorderFill>'
        "</hp:secPr>"
        '<hp:ctrl><hp:colPr id="" type="NEWSPAPER" layout="LEFT" colCount="1" sameSz="1" sameGap="0"/></hp:ctrl>'
        "<hp:t>2026년 2차 모두의 창업 프로젝트 — 1R 사업계획서 (광운대학교)</hp:t>"
        "</hp:run></hp:p>",
        render_outer_2x2_hwpx(
            "Q1. 문제 인식(problem)", " ※ 아이디어의 개발 동기·목적 및 목표 시장 등", q1_blks, id_counter, page_break=0
        ),
        render_outer_2x2_hwpx(
            "Q2. 실현 가능성(solution)", " ※ 아이디어의 개발 방안, 준비 정도, 차별화 방안", q2_blks, id_counter, page_break=1
        ),
        render_outer_2x2_hwpx(
            "Q3. 성장전략(scale-up)", " ※ 사업화 계획, 추진일정, 자금 소요 및 조달 계획", q3_blks, id_counter, page_break=1
        ),
        render_outer_2x2_hwpx(
            "Q4. 향후 발전방안 (development)", " ※ 2R 이후 사업 고도화, 후속 계획 등", q4_blks, id_counter, page_break=1
        ),
    ]

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
    for cid, (h_hu, b_flag, i_flag, col_hex) in enumerate(cs_list):
        b_tag = "<hh:bold/>" if b_flag else ""
        i_tag = "<hh:italic/>" if i_flag else ""
        cs_xml_items.append(
            f'<hh:charPr id="{cid}" height="{h_hu}" textColor="#{col_hex}" shadeColor="none" useFontSpace="0" useKerning="0" symMark="NONE" borderFillIDRef="1">'
            '<hh:fontRef hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>'
            '<hh:ratio hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/>'
            '<hh:spacing hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>'
            '<hh:relSz hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/>'
            '<hh:offset hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>'
            f"{b_tag}{i_tag}"
            '<hh:underline type="NONE" shape="SOLID" color="#000000"/>'
            '<hh:strikeout shape="NONE" color="#000000"/>'
            '<hh:outline type="NONE"/>'
            '<hh:shadow type="NONE" color="#C0C0C0" offsetX="10" offsetY="10"/>'
            "</hh:charPr>"
        )

    ps_xml_items: List[str] = []
    for pid, align_str in enumerate(("JUSTIFY", "LEFT", "CENTER", "RIGHT")):
        ps_xml_items.append(
            f'<hh:paraPr id="{pid}" tabPrIDRef="0" condense="0" fontLineHeight="0" snapToGrid="0" suppressLineNumbers="0" checked="0">'
            f'<hh:align horizontal="{align_str}" vertical="BASELINE"/>'
            '<hh:heading type="NONE" idRef="0" level="0"/>'
            '<hh:breakSetting breakLatinWord="KEEP_WORD" breakNonLatinWord="BREAK_WORD" widowOrphan="0" keepWithNext="0" keepLines="0" pageBreakBefore="0" lineWrap="BREAK"/>'
            '<hh:autoSpacing eAsianEng="0" eAsianNum="0"/>'
            '<hh:margin><hc:intent value="0" unit="HWPUNIT"/><hc:left value="0" unit="HWPUNIT"/><hc:right value="0" unit="HWPUNIT"/><hc:prev value="20" unit="HWPUNIT"/><hc:next value="35" unit="HWPUNIT"/></hh:margin>'
            '<hh:lineSpacing type="PERCENT" value="122" unit="HWPUNIT"/>'
            '<hh:border borderFillIDRef="1" offsetLeft="0" offsetRight="0" offsetTop="0" offsetBottom="0" connect="0" ignoreMargin="0"/>'
            "</hh:paraPr>"
        )

    langs = ("HANGUL", "LATIN", "HANJA", "JAPANESE", "OTHER", "SYMBOL", "USER")
    fontfaces_xml = "".join(
        f'<hh:fontface lang="{lg}" fontCnt="1"><hh:font id="0" face="맑은 고딕" type="TTF" isEmbedded="0"><hh:typeInfo familyType="FCAT_GOTHIC" weight="6" proportion="4" contrast="0" strokeVariation="1" armStyle="1" letterform="1" midline="1" xHeight="1"/></hh:font></hh:fontface>'
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

    return {"out_hwpx_path": out_hwpx_path, "bytes": os.path.getsize(out_hwpx_path)}
