"""
Markup (.md, .html, .txt), Tabular (.csv, .tsv, .xlsx), and Native PDF 1.4 (.pdf)
Lifters & Projectors for T-DIR v4.0.
Enables N-to-M conversion across Markdown, HTML5 DOM, CSV/TSV, OpenXML Spreadsheet (.xlsx),
Deterministic Multi-Page PDF 1.4 (.pdf), and HWP v5.0 / OWPML (.hwpx) / OOXML (.docx).
"""

import os
import re
import csv
import math
import base64
import zipfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape as xml_esc
from html.parser import HTMLParser
from typing import List, Dict, Tuple, Optional

from .ir_schema import (
    SectionIR,
    MediaObjectIR,
    TargetDocumentSpecIR,
    extract_ir_non_ws_text,
    extract_ordinal_index,
    detect_image_format,
    compute_weighted_text_units,
    strip_illegal_xml_control_chars,
    validate_safe_zip_archive,
    safe_parse_xml_bytes,
    sanitize_spreadsheet_cell_formula,
    parse_latex_to_equation_ast,
    compile_equation_ast_to_latex,
    compile_equation_ast_to_linear_text,
)
from .elastic_solver import ElasticLayoutState


def _split_markdown_runs_with_equations(text: str, is_bold: bool = False) -> List[dict]:
    cleaned = strip_illegal_xml_control_chars(text)
    if "$" not in cleaned:
        return [{"text": cleaned, "bold": is_bold, "italic": False, "color": "000000"}] if cleaned else []

    runs: List[dict] = []
    pattern = re.compile(r"(\$\$[^$]+\$\$|\$[^$]+\$)")
    pos = 0
    for m in pattern.finditer(cleaned):
        if m.start() > pos:
            plain = cleaned[pos : m.start()]
            if plain:
                runs.append({"text": plain, "bold": is_bold, "italic": False, "color": "000000"})
        math_token = m.group(1)
        eq_ast = parse_latex_to_equation_ast(math_token)
        lin_txt = compile_equation_ast_to_linear_text(eq_ast)
        if lin_txt:
            runs.append(
                {
                    "text": lin_txt,
                    "bold": is_bold,
                    "italic": False,
                    "color": "000000",
                    "equation_ast": eq_ast,
                }
            )
        pos = m.end()
    if pos < len(cleaned):
        tail = cleaned[pos:]
        if tail:
            runs.append({"text": tail, "bold": is_bold, "italic": False, "color": "000000"})
    return runs


def lift_markdown_to_ir(md_path: str, target_width_hu: int = 46200) -> Dict[str, SectionIR]:
    with open(md_path, "r", encoding="utf-8") as f:
        lines = f.read().splitlines()

    sections: Dict[str, SectionIR] = {}
    cur_sec_id = "SEC_01"
    cur_heading = "Markdown Root"
    cur_ord = 1
    cur_blocks: List[dict] = []
    table_buffer: List[List[str]] = []

    def flush_table() -> None:
        nonlocal table_buffer
        if not table_buffer:
            return
        n_rows = len(table_buffer)
        n_cols = max(len(r) for r in table_buffer)
        w_col = target_width_hu // max(1, n_cols)
        col_w = [w_col] * n_cols
        col_w[-1] += target_width_hu - sum(col_w)
        cells: List[dict] = []
        for r_i, row_vals in enumerate(table_buffer):
            for c_i in range(n_cols):
                val = row_vals[c_i].strip() if c_i < len(row_vals) else ""
                cell_runs = _split_markdown_runs_with_equations(val, is_bold=(r_i == 0))
                if not cell_runs:
                    cell_runs = [{"text": "", "bold": r_i == 0, "italic": False, "color": "000000"}]
                cells.append(
                    {
                        "row": r_i,
                        "col": c_i,
                        "row_span": 1,
                        "col_span": 1,
                        "width_hu": col_w[c_i],
                        "height_hu": 900,
                        "mar": (140, 140, 35, 35),
                        "borders": {s: (1, 1, (0x94, 0xA3, 0xB8)) for s in ("left", "right", "top", "bottom")},
                        "fill_hex": "F1F5F9" if r_i == 0 else None,
                        "paras": [
                            {
                                "type": "p",
                                "jc": "center" if r_i == 0 else "left",
                                "left_tw": 0,
                                "right_tw": 0,
                                "indent_tw": 0,
                                "before_tw": 0,
                                "after_tw": 0,
                                "runs": cell_runs,
                                "r_embed": None,
                            }
                        ],
                    }
                )
        cur_blocks.append(
            {
                "type": "tbl",
                "n_rows": n_rows,
                "n_cols": n_cols,
                "col_widths_hu": col_w,
                "row_cols": [n_cols] * n_rows,
                "cells": cells,
                "width_hu": target_width_hu,
                "height_hu": n_rows * 900,
            }
        )
        table_buffer = []

    def flush_sec() -> None:
        nonlocal cur_blocks
        flush_table()
        if not cur_blocks:
            return
        _, n_c, sha = extract_ir_non_ws_text(cur_blocks)
        sections[cur_sec_id] = SectionIR(
            section_uid=cur_sec_id,
            heading_text=cur_heading,
            blocks=cur_blocks,
            non_ws_char_count=n_c,
            text_sha256=sha,
            ordinal_index=cur_ord,
            section_role="BODY",
        )
        cur_blocks = []

    in_code_block = False
    for line in lines:
        stripped = line.strip()
        if stripped.startswith("```") or stripped.startswith("~~~"):
            flush_table()
            in_code_block = not in_code_block
            continue
        if in_code_block:
            if line:
                cur_blocks.append(
                    {
                        "type": "p",
                        "jc": "left",
                        "left_tw": 0,
                        "right_tw": 0,
                        "indent_tw": 0,
                        "before_tw": 0,
                        "after_tw": 4,
                        "runs": [{"text": strip_illegal_xml_control_chars(line), "bold": False, "italic": False, "color": "000000"}],
                        "r_embed": None,
                    }
                )
            continue
        if stripped.startswith("|") and stripped.endswith("|") and len(stripped) >= 2:
            raw_inner = stripped[1:-1]
            cells_raw = [c.replace(r"\|", "|").strip() for c in re.split(r"(?<!\\)\|", raw_inner)]
            if all(set(c) <= {"-", ":", " "} for c in cells_raw):
                continue
            table_buffer.append(cells_raw)
            continue
        flush_table()
        if not stripped:
            continue
        if stripped.startswith("# "):
            flush_sec()
            cur_heading = strip_illegal_xml_control_chars(stripped[2:].strip())
            sec_seq = len(sections) + 1
            cur_sec_id = f"SEC_{sec_seq:02d}"
            cur_ord = extract_ordinal_index(cur_heading) or sec_seq
            continue
        is_bold_heading = stripped.startswith("##")
        clean_txt = stripped.lstrip("#").strip() if is_bold_heading else stripped
        p_runs = _split_markdown_runs_with_equations(clean_txt, is_bold=is_bold_heading)
        if p_runs:
            cur_blocks.append(
                {
                    "type": "p",
                    "jc": "both",
                    "left_tw": 0,
                    "right_tw": 0,
                    "indent_tw": 0,
                    "before_tw": 0,
                    "after_tw": 8,
                    "runs": p_runs,
                    "r_embed": None,
                }
            )

    flush_sec()
    return sections


def project_ir_to_markdown(
    sections: Dict[str, SectionIR],
    out_md_path: str,
    emit_latex_delimiters: bool = False,
) -> dict:
    def _render_run_md(r: dict) -> str:
        eq_ast = r.get("equation_ast")
        if emit_latex_delimiters and eq_ast:
            return f"${compile_equation_ast_to_latex(eq_ast)}$"
        return strip_illegal_xml_control_chars(str(r.get("text", "")))

    out_lines: List[str] = []
    for sec in sections.values():
        out_lines.append(f"# {strip_illegal_xml_control_chars(sec.heading_text)}")
        out_lines.append("")
        for blk in sec.blocks:
            if blk["type"] == "p":
                txt = "".join(_render_run_md(r) for r in blk["runs"])
                out_lines.append(txt)
                out_lines.append("")
            elif blk["type"] == "tbl":
                rows_by_r: Dict[int, List[dict]] = {}
                for c in blk["cells"]:
                    rows_by_r.setdefault(c["row"], []).append(c)
                for r_i in range(blk["n_rows"]):
                    row_cells = sorted(rows_by_r.get(r_i, []), key=lambda x: x["col"])
                    cell_strs = [
                        " ".join("".join(_render_run_md(r) for r in p["runs"]) for p in rc["paras"])
                        for rc in row_cells
                    ]
                    out_lines.append("| " + " | ".join(cell_strs) + " |")
                    if r_i == 0:
                        out_lines.append("| " + " | ".join(["---"] * max(1, len(cell_strs))) + " |")
                out_lines.append("")

    os.makedirs(os.path.dirname(os.path.abspath(out_md_path)), exist_ok=True)
    content = "\n".join(out_lines)
    with open(out_md_path, "w", encoding="utf-8") as f:
        f.write(content)
    return {"out_md_path": out_md_path, "bytes": os.path.getsize(out_md_path)}


def lift_csv_to_ir(csv_path: str, delimiter: str = ",", target_width_hu: int = 46200) -> Dict[str, SectionIR]:
    with open(csv_path, "r", encoding="utf-8-sig") as f:
        rows = list(csv.reader(f, delimiter=delimiter))
    if not rows:
        return {}
    n_rows = len(rows)
    n_cols = max(len(r) for r in rows)
    w_col = target_width_hu // max(1, n_cols)
    col_w = [w_col] * n_cols
    col_w[-1] += target_width_hu - sum(col_w)
    cells: List[dict] = []
    for r_i, row_vals in enumerate(rows):
        for c_i in range(n_cols):
            val = strip_illegal_xml_control_chars(row_vals[c_i].strip()) if c_i < len(row_vals) else ""
            cells.append(
                {
                    "row": r_i,
                    "col": c_i,
                    "row_span": 1,
                    "col_span": 1,
                    "width_hu": col_w[c_i],
                    "height_hu": 900,
                    "mar": (140, 140, 35, 35),
                    "borders": {s: (1, 1, (0x94, 0xA3, 0xB8)) for s in ("left", "right", "top", "bottom")},
                    "fill_hex": "F1F5F9" if r_i == 0 else None,
                    "paras": [
                        {
                            "type": "p",
                            "jc": "center" if r_i == 0 else "left",
                            "left_tw": 0,
                            "right_tw": 0,
                            "indent_tw": 0,
                            "before_tw": 0,
                            "after_tw": 0,
                            "runs": [{"text": val, "bold": r_i == 0, "italic": False, "color": "000000"}],
                            "r_embed": None,
                        }
                    ],
                }
            )
    tbl_blk = {
        "type": "tbl",
        "n_rows": n_rows,
        "n_cols": n_cols,
        "col_widths_hu": col_w,
        "row_cols": [n_cols] * n_rows,
        "cells": cells,
        "width_hu": target_width_hu,
        "height_hu": n_rows * 900,
    }
    _, n_c, sha = extract_ir_non_ws_text([tbl_blk])
    return {
        "SEC_01": SectionIR(
            section_uid="SEC_01",
            heading_text=os.path.basename(csv_path),
            blocks=[tbl_blk],
            non_ws_char_count=n_c,
            text_sha256=sha,
            ordinal_index=1,
            section_role="BODY",
        )
    }


def project_ir_to_csv(
    sections: Dict[str, SectionIR],
    out_csv_path: str,
    delimiter: str = ",",
    neutralize_dde: bool = True,
) -> dict:
    rows_out: List[List[str]] = []
    for sec in sections.values():
        for blk in sec.blocks:
            if blk["type"] == "tbl":
                rows_by_r: Dict[int, List[dict]] = {}
                for c in blk["cells"]:
                    rows_by_r.setdefault(c["row"], []).append(c)
                for r_i in range(blk["n_rows"]):
                    row_cells = sorted(rows_by_r.get(r_i, []), key=lambda x: x["col"])
                    rows_out.append(
                        [
                            sanitize_spreadsheet_cell_formula(
                                " ".join("".join(r["text"] for r in p["runs"]) for p in rc["paras"]),
                                neutralize_dde=neutralize_dde,
                            )
                            for rc in row_cells
                        ]
                    )
            elif blk["type"] == "p":
                txt = "".join(r["text"] for r in blk["runs"]).strip()
                if txt:
                    rows_out.append(
                        [
                            sec.section_uid,
                            sanitize_spreadsheet_cell_formula(txt, neutralize_dde=neutralize_dde),
                        ]
                    )
    os.makedirs(os.path.dirname(os.path.abspath(out_csv_path)), exist_ok=True)
    with open(out_csv_path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f, delimiter=delimiter)
        writer.writerows(rows_out)
    return {"out_csv_path": out_csv_path, "bytes": os.path.getsize(out_csv_path), "rows": len(rows_out)}


# ==============================================================================
# 2. Bidirectional OpenXML Spreadsheet (.xlsx) Lifter & Projector ([P3-01], [SEC-05])
# ==============================================================================

XLSX_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def _col_idx_to_excel_letters(col_idx_0: int) -> str:
    n = col_idx_0 + 1
    chars: List[str] = []
    while n > 0:
        n, rem = divmod(n - 1, 26)
        chars.append(chr(65 + rem))
    return "".join(reversed(chars))


def _excel_ref_to_row_col(cell_ref: str) -> Tuple[int, int]:
    m = re.match(r"^([A-Za-z]+)(\d+)$", cell_ref.strip())
    if not m:
        return 0, 0
    col_str, row_str = m.group(1).upper(), m.group(2)
    col_val = 0
    for ch in col_str:
        col_val = col_val * 26 + (ord(ch) - 64)
    return max(0, int(row_str) - 1), max(0, col_val - 1)


def project_ir_to_xlsx(
    sections: Dict[str, SectionIR],
    out_xlsx_path: str,
    neutralize_dde: bool = True,
) -> dict:
    """
    Projects T-DIR `SectionIR` blocks into a native OpenXML Spreadsheet (.xlsx) archive
    with shared strings (`xl/sharedStrings.xml`), styles (`xl/styles.xml`), column widths,
    and merged cell ranges (`<mergeCells>`) ([P3-01], [SEC-05]).
    """
    sst_list: List[str] = []
    sst_map: Dict[str, int] = {}

    def _get_sst_id(raw_val: str) -> int:
        safe_val = sanitize_spreadsheet_cell_formula(raw_val, neutralize_dde=neutralize_dde)
        if safe_val not in sst_map:
            sst_map[safe_val] = len(sst_list)
            sst_list.append(safe_val)
        return sst_map[safe_val]

    sheet_rows_xml: List[str] = []
    merge_refs: List[str] = []
    cur_excel_row = 1
    max_cols = 2

    for sec in sections.values():
        # Emit section header row
        s_uid_id = _get_sst_id(f"__SEC_HDR__:{sec.section_uid}")
        s_title_id = _get_sst_id(sec.heading_text)
        r_ref1 = f"A{cur_excel_row}"
        r_ref2 = f"B{cur_excel_row}"
        sheet_rows_xml.append(
            f'<row r="{cur_excel_row}" ht="22" customHeight="1">'
            f'<c r="{r_ref1}" t="s" s="2"><v>{s_uid_id}</v></c>'
            f'<c r="{r_ref2}" t="s" s="2"><v>{s_title_id}</v></c>'
            f"</row>"
        )
        cur_excel_row += 1

        for blk in sec.blocks:
            if blk["type"] == "p":
                txt = "".join(r.get("text", "") for r in blk.get("runs", [])).strip()
                if not txt:
                    continue
                p_marker_id = _get_sst_id("__PARA__")
                p_txt_id = _get_sst_id(txt)
                c1 = f"A{cur_excel_row}"
                c2 = f"B{cur_excel_row}"
                sheet_rows_xml.append(
                    f'<row r="{cur_excel_row}">'
                    f'<c r="{c1}" t="s" s="0"><v>{p_marker_id}</v></c>'
                    f'<c r="{c2}" t="s" s="0"><v>{p_txt_id}</v></c>'
                    f"</row>"
                )
                cur_excel_row += 1
            elif blk["type"] == "tbl":
                n_r = blk.get("n_rows", 1)
                n_c = blk.get("n_cols", 1)
                max_cols = max(max_cols, n_c)
                t_marker_id = _get_sst_id(f"__TBL__:{n_r}:{n_c}")
                sheet_rows_xml.append(
                    f'<row r="{cur_excel_row}"><c r="A{cur_excel_row}" t="s" s="1"><v>{t_marker_id}</v></c></row>'
                )
                cur_excel_row += 1

                base_tbl_row = cur_excel_row
                rows_by_r: Dict[int, List[dict]] = {}
                for cell in blk.get("cells", []):
                    rows_by_r.setdefault(cell["row"], []).append(cell)

                for r_i in range(n_r):
                    excel_r = base_tbl_row + r_i
                    row_cells = sorted(rows_by_r.get(r_i, []), key=lambda x: x["col"])
                    c_xml_parts: List[str] = []
                    for cell in row_cells:
                        c_i = cell["col"]
                        r_sp = cell.get("row_span", 1)
                        c_sp = cell.get("col_span", 1)
                        c_let = _col_idx_to_excel_letters(c_i)
                        c_ref = f"{c_let}{excel_r}"
                        cell_txt = " ".join(
                            "".join(r.get("text", "") for r in p.get("runs", []))
                            for p in cell.get("paras", [])
                        )
                        sid = _get_sst_id(cell_txt)
                        style_id = 1 if (r_i == 0 or cell.get("fill_hex")) else 0
                        c_xml_parts.append(f'<c r="{c_ref}" t="s" s="{style_id}"><v>{sid}</v></c>')
                        if r_sp > 1 or c_sp > 1:
                            end_let = _col_idx_to_excel_letters(c_i + c_sp - 1)
                            end_r = excel_r + r_sp - 1
                            merge_refs.append(f"{c_ref}:{end_let}{end_r}")
                    sheet_rows_xml.append(f'<row r="{excel_r}">{"".join(c_xml_parts)}</row>')
                cur_excel_row += n_r

    cols_xml = f'<cols><col min="1" max="{max(2, max_cols)}" width="24" customWidth="1"/></cols>'
    merge_xml = (
        f'<mergeCells count="{len(merge_refs)}">'
        + "".join(f'<mergeCell ref="{mref}"/>' for mref in merge_refs)
        + "</mergeCells>"
        if merge_refs
        else ""
    )

    sheet1_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<worksheet xmlns="{XLSX_NS}">'
        f"{cols_xml}"
        f'<sheetData>{"".join(sheet_rows_xml)}</sheetData>'
        f"{merge_xml}"
        "</worksheet>"
    )

    sst_items_xml = "".join(f'<si><t xml:space="preserve">{xml_esc(s)}</t></si>' for s in sst_list)
    shared_strings_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<sst xmlns="{XLSX_NS}" count="{len(sst_list)}" uniqueCount="{len(sst_list)}">'
        f"{sst_items_xml}"
        "</sst>"
    )

    styles_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<styleSheet xmlns="{XLSX_NS}">'
        '<fonts count="3">'
        '<font><sz val="10"/><name val="Malgun Gothic"/></font>'
        '<font><b/><sz val="10"/><name val="Malgun Gothic"/></font>'
        '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Malgun Gothic"/></font>'
        "</fonts>"
        '<fills count="4">'
        '<fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="gray125"/></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFF1F5F9"/></patternFill></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FF3A3C84"/></patternFill></fill>'
        "</fills>"
        '<borders count="2">'
        "<border><left/><right/><top/><bottom/><diagonal/></border>"
        '<border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/><diagonal/></border>'
        "</borders>"
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="3">'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>'
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>'
        '<xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>'
        "</cellXfs>"
        "</styleSheet>"
    )

    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<workbook xmlns="{XLSX_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="TDIR_Sheet1" sheetId="1" r:id="rId1"/></sheets>'
        "</workbook>"
    )

    wb_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
        "</Relationships>"
    )

    root_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        "</Relationships>"
    )

    content_types_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
        "</Types>"
    )

    os.makedirs(os.path.dirname(os.path.abspath(out_xlsx_path)), exist_ok=True)
    with zipfile.ZipFile(out_xlsx_path, "w", zipfile.ZIP_DEFLATED) as zout:
        zout.writestr("[Content_Types].xml", content_types_xml.encode("utf-8"))
        zout.writestr("_rels/.rels", root_rels_xml.encode("utf-8"))
        zout.writestr("xl/workbook.xml", workbook_xml.encode("utf-8"))
        zout.writestr("xl/_rels/workbook.xml.rels", wb_rels_xml.encode("utf-8"))
        zout.writestr("xl/styles.xml", styles_xml.encode("utf-8"))
        zout.writestr("xl/sharedStrings.xml", shared_strings_xml.encode("utf-8"))
        zout.writestr("xl/worksheets/sheet1.xml", sheet1_xml.encode("utf-8"))

    return {
        "out_xlsx_path": out_xlsx_path,
        "bytes": os.path.getsize(out_xlsx_path),
        "rows_emitted": cur_excel_row - 1,
        "merged_ranges_count": len(merge_refs),
        "shared_strings_count": len(sst_list),
    }


def lift_xlsx_to_ir(xlsx_path: str, target_width_hu: int = 46200) -> Dict[str, SectionIR]:
    """
    Lifts an OpenXML Spreadsheet (.xlsx) archive into canonical `SectionIR` objects,
    preserving merged cell spans (`<mergeCells>`), shared strings, and multi-section markers ([P3-01]).
    """
    with zipfile.ZipFile(xlsx_path, "r") as zf:
        validate_safe_zip_archive(zf)
        sst: List[str] = []
        if "xl/sharedStrings.xml" in zf.namelist():
            sst_root = safe_parse_xml_bytes(zf.read("xl/sharedStrings.xml"))
            for si in sst_root.findall(f".//{{{XLSX_NS}}}si"):
                txt = "".join(t.text or "" for t in si.findall(f".//{{{XLSX_NS}}}t"))
                sst.append(strip_illegal_xml_control_chars(txt))

        sheet_files = sorted(
            nm for nm in zf.namelist() if nm.startswith("xl/worksheets/sheet") and nm.endswith(".xml")
        )
        if not sheet_files:
            return {}

        sheet_root = safe_parse_xml_bytes(zf.read(sheet_files[0]))
        merge_span_map: Dict[Tuple[int, int], Tuple[int, int]] = {}
        covered_coords: set = set()
        for mc in sheet_root.findall(f".//{{{XLSX_NS}}}mergeCell"):
            ref = mc.get("ref", "")
            if ":" in ref:
                start_ref, end_ref = ref.split(":", 1)
                r0, c0 = _excel_ref_to_row_col(start_ref)
                r1, c1 = _excel_ref_to_row_col(end_ref)
                r_span = max(1, r1 - r0 + 1)
                c_span = max(1, c1 - c0 + 1)
                merge_span_map[(r0, c0)] = (r_span, c_span)
                for rr in range(r0, r1 + 1):
                    for cc in range(c0, c1 + 1):
                        if (rr, cc) != (r0, c0):
                            covered_coords.add((rr, cc))

        rows_data: Dict[int, Dict[int, Tuple[str, int]]] = {}
        for row_el in sheet_root.findall(f".//{{{XLSX_NS}}}row"):
            r_num = int(row_el.get("r", "1") or 1) - 1
            row_dict: Dict[int, Tuple[str, int]] = {}
            for c_el in row_el.findall(f"{{{XLSX_NS}}}c"):
                c_ref = c_el.get("r", "A1")
                _, c_idx = _excel_ref_to_row_col(c_ref)
                c_type = c_el.get("t", "")
                s_idx = int(c_el.get("s", "0") or 0)
                val_str = ""
                if c_type == "s":
                    v_el = c_el.find(f"{{{XLSX_NS}}}v")
                    if v_el is not None and v_el.text and v_el.text.isdigit():
                        sid = int(v_el.text)
                        if 0 <= sid < len(sst):
                            val_str = sst[sid]
                elif c_type == "inlineStr":
                    val_str = "".join(t.text or "" for t in c_el.findall(f".//{{{XLSX_NS}}}t"))
                else:
                    v_el = c_el.find(f"{{{XLSX_NS}}}v")
                    if v_el is not None and v_el.text:
                        val_str = v_el.text
                row_dict[c_idx] = (strip_illegal_xml_control_chars(val_str), s_idx)
            rows_data[r_num] = row_dict

    if not rows_data:
        return {}

    sorted_r_indices = sorted(rows_data.keys())
    has_structured_markers = any(
        rows_data[r_idx].get(0, ("", 0))[0].startswith(("__SEC_HDR__:", "__PARA__", "__TBL__:"))
        for r_idx in sorted_r_indices
    )

    sections: Dict[str, SectionIR] = {}
    if has_structured_markers:
        cur_sec_uid = "SEC_01"
        cur_heading = "Spreadsheet Section"
        cur_ord = 1
        cur_blocks: List[dict] = []

        def _flush_xlsx_sec() -> None:
            nonlocal cur_blocks
            if not cur_blocks:
                return
            _, n_c, sha = extract_ir_non_ws_text(cur_blocks)
            sections[cur_sec_uid] = SectionIR(
                section_uid=cur_sec_uid,
                heading_text=cur_heading,
                blocks=cur_blocks,
                non_ws_char_count=n_c,
                text_sha256=sha,
                ordinal_index=cur_ord,
                section_role="BODY",
            )
            cur_blocks = []

        pos = 0
        while pos < len(sorted_r_indices):
            r_idx = sorted_r_indices[pos]
            r_map = rows_data[r_idx]
            c0_val = r_map.get(0, ("", 0))[0]
            if c0_val.startswith("__SEC_HDR__:"):
                _flush_xlsx_sec()
                cur_sec_uid = c0_val.split(":", 1)[1] or f"SEC_{len(sections) + 1:02d}"
                cur_heading = r_map.get(1, (cur_sec_uid, 0))[0]
                cur_ord = extract_ordinal_index(cur_heading) or extract_ordinal_index(cur_sec_uid) or (len(sections) + 1)
                pos += 1
                continue
            if c0_val == "__PARA__":
                p_txt = r_map.get(1, ("", 0))[0]
                if p_txt:
                    cur_blocks.append(
                        {
                            "type": "p",
                            "jc": "both",
                            "left_tw": 0,
                            "right_tw": 0,
                            "indent_tw": 0,
                            "before_tw": 0,
                            "after_tw": 8,
                            "runs": [{"text": p_txt, "bold": False, "italic": False, "color": "000000"}],
                            "r_embed": None,
                        }
                    )
                pos += 1
                continue
            if c0_val.startswith("__TBL__:"):
                parts = c0_val.split(":")
                n_r = int(parts[1]) if len(parts) > 1 and parts[1].isdigit() else 1
                n_c = int(parts[2]) if len(parts) > 2 and parts[2].isdigit() else 1
                base_r = r_idx + 1
                w_col = target_width_hu // max(1, n_c)
                col_w = [w_col] * n_c
                col_w[-1] += target_width_hu - sum(col_w)
                cells: List[dict] = []
                row_cols: List[int] = []
                for tr_i in range(n_r):
                    abs_r = base_r + tr_i
                    tr_map = rows_data.get(abs_r, {})
                    rc_count = 0
                    for tc_i in range(n_c):
                        if (abs_r, tc_i) in covered_coords:
                            continue
                        r_sp, c_sp = merge_span_map.get((abs_r, tc_i), (1, 1))
                        cell_txt, s_idx = tr_map.get(tc_i, ("", 0))
                        cw_hu = sum(col_w[tc_i : tc_i + c_sp])
                        cells.append(
                            {
                                "row": tr_i,
                                "col": tc_i,
                                "row_span": r_sp,
                                "col_span": c_sp,
                                "width_hu": cw_hu,
                                "height_hu": 900 * r_sp,
                                "mar": (140, 140, 35, 35),
                                "borders": {s: (1, 1, (0x94, 0xA3, 0xB8)) for s in ("left", "right", "top", "bottom")},
                                "fill_hex": "F1F5F9" if s_idx == 1 else None,
                                "paras": [
                                    {
                                        "type": "p",
                                        "jc": "center" if tr_i == 0 else "left",
                                        "left_tw": 0,
                                        "right_tw": 0,
                                        "indent_tw": 0,
                                        "before_tw": 0,
                                        "after_tw": 0,
                                        "runs": [
                                            {
                                                "text": cell_txt,
                                                "bold": s_idx == 1,
                                                "italic": False,
                                                "color": "000000",
                                            }
                                        ],
                                        "r_embed": None,
                                    }
                                ],
                            }
                        )
                        rc_count += 1
                    row_cols.append(max(1, rc_count))
                cur_blocks.append(
                    {
                        "type": "tbl",
                        "n_rows": n_r,
                        "n_cols": n_c,
                        "col_widths_hu": col_w,
                        "row_cols": row_cols,
                        "cells": cells,
                        "width_hu": target_width_hu,
                        "height_hu": n_r * 900,
                    }
                )
                pos += 1 + n_r
                continue
            pos += 1
        _flush_xlsx_sec()
        return sections

    # Unstructured general spreadsheet grid
    max_r = max(sorted_r_indices) + 1
    max_c = max((max(rm.keys()) + 1 for rm in rows_data.values() if rm), default=1)
    w_col = target_width_hu // max(1, max_c)
    col_w = [w_col] * max_c
    col_w[-1] += target_width_hu - sum(col_w)
    cells_all: List[dict] = []
    row_cols_all: List[int] = []
    for r_i in range(max_r):
        rm = rows_data.get(r_i, {})
        rc_cnt = 0
        for c_i in range(max_c):
            if (r_i, c_i) in covered_coords:
                continue
            r_sp, c_sp = merge_span_map.get((r_i, c_i), (1, 1))
            val, s_idx = rm.get(c_i, ("", 0))
            cells_all.append(
                {
                    "row": r_i,
                    "col": c_i,
                    "row_span": r_sp,
                    "col_span": c_sp,
                    "width_hu": sum(col_w[c_i : c_i + c_sp]),
                    "height_hu": 900 * r_sp,
                    "mar": (140, 140, 35, 35),
                    "borders": {s: (1, 1, (0x94, 0xA3, 0xB8)) for s in ("left", "right", "top", "bottom")},
                    "fill_hex": "F1F5F9" if (r_i == 0 or s_idx == 1) else None,
                    "paras": [
                        {
                            "type": "p",
                            "jc": "center" if r_i == 0 else "left",
                            "left_tw": 0,
                            "right_tw": 0,
                            "indent_tw": 0,
                            "before_tw": 0,
                            "after_tw": 0,
                            "runs": [{"text": val, "bold": r_i == 0 or s_idx == 1, "italic": False, "color": "000000"}],
                            "r_embed": None,
                        }
                    ],
                }
            )
            rc_cnt += 1
        row_cols_all.append(max(1, rc_cnt))

    tbl_blk = {
        "type": "tbl",
        "n_rows": max_r,
        "n_cols": max_c,
        "col_widths_hu": col_w,
        "row_cols": row_cols_all,
        "cells": cells_all,
        "width_hu": target_width_hu,
        "height_hu": max_r * 900,
    }
    _, n_c, sha = extract_ir_non_ws_text([tbl_blk])
    return {
        "SEC_01": SectionIR(
            section_uid="SEC_01",
            heading_text=os.path.basename(xlsx_path),
            blocks=[tbl_blk],
            non_ws_char_count=n_c,
            text_sha256=sha,
            ordinal_index=1,
            section_role="BODY",
        )
    }


def extract_xlsx_all_text(xlsx_path: str) -> str:
    sections = lift_xlsx_to_ir(xlsx_path)
    all_blocks = [b for s in sections.values() for b in s.blocks]
    norm_txt, _, _ = extract_ir_non_ws_text(all_blocks)
    return norm_txt


# ==============================================================================
# 3. Bidirectional Semantic HTML5 DOM (.html) Lifter & Projector ([P3-01], [P2-02])
# ==============================================================================


def project_ir_to_html(
    sections: Dict[str, SectionIR],
    out_html_path: str,
    target_spec: Optional[TargetDocumentSpecIR] = None,
    attachments: Optional[List[Tuple[bytes, str]]] = None,
) -> dict:
    """
    Projects T-DIR `SectionIR` blocks into a semantic HTML5 document with `@page` CSS
    geometry, `rowspan`/`colspan` tables, `<math data-latex="...">` equations, and
    base64-encoded media ([P3-01], [P2-02]).
    """
    _ = attachments
    font_family = (target_spec.body_font_family if target_spec and target_spec.body_font_family else "Malgun Gothic")
    body_pt = round((target_spec.body_font_size_hu if target_spec and target_spec.body_font_size_hu > 0 else 1000) / 100.0, 2)
    tbl_pt = round((target_spec.table_font_size_hu if target_spec and target_spec.table_font_size_hu > 0 else 950) / 100.0, 2)

    sec_html_list: List[str] = []
    for sec in sections.values():
        blk_html_list: List[str] = [
            f'<h1 class="tdir-sec-heading">{xml_esc(strip_illegal_xml_control_chars(sec.heading_text))}</h1>'
        ]
        if sec.subtitle_text:
            blk_html_list.append(
                f'<div class="tdir-sec-subtitle">{xml_esc(strip_illegal_xml_control_chars(sec.subtitle_text))}</div>'
            )

        for blk in sec.blocks:
            if blk["type"] == "p":
                jc = blk.get("jc", "both")
                align_css = {"both": "justify", "left": "left", "center": "center", "right": "right"}.get(jc, "justify")
                run_spans: List[str] = []
                for r in blk.get("runs", []):
                    txt = strip_illegal_xml_control_chars(str(r.get("text", "")))
                    eq_ast = r.get("equation_ast")
                    if eq_ast and not txt:
                        txt = compile_equation_ast_to_linear_text(eq_ast)
                    if not txt:
                        continue
                    if eq_ast:
                        latex_attr = xml_esc(compile_equation_ast_to_latex(eq_ast))
                        run_spans.append(f'<math class="tdir-eq" data-latex="{latex_attr}">{xml_esc(txt)}</math>')
                    else:
                        styles: List[str] = []
                        if r.get("bold"):
                            styles.append("font-weight:bold")
                        if r.get("italic"):
                            styles.append("font-style:italic")
                        if r.get("underline"):
                            styles.append("text-decoration:underline")
                        col = r.get("color", "000000")
                        if col and col != "000000":
                            styles.append(f"color:#{col}")
                        st_attr = f' style="{";".join(styles)}"' if styles else ""
                        run_spans.append(f"<span{st_attr}>{xml_esc(txt)}</span>")
                if run_spans:
                    blk_html_list.append(f'<p style="text-align:{align_css}">{"".join(run_spans)}</p>')
            elif blk["type"] == "tbl":
                rows_by_r: Dict[int, List[dict]] = {}
                for c in blk.get("cells", []):
                    rows_by_r.setdefault(c["row"], []).append(c)
                tr_list: List[str] = []
                for r_i in range(blk.get("n_rows", 1)):
                    row_cells = sorted(rows_by_r.get(r_i, []), key=lambda x: x["col"])
                    td_list: List[str] = []
                    for cell in row_cells:
                        r_sp = cell.get("row_span", 1)
                        c_sp = cell.get("col_span", 1)
                        fill = cell.get("fill_hex")
                        bg_style = f"background-color:#{fill};" if fill else ""
                        p_in_cell: List[str] = []
                        for p in cell.get("paras", []):
                            p_txt = "".join(
                                (
                                    f'<math class="tdir-eq" data-latex="{xml_esc(compile_equation_ast_to_latex(r["equation_ast"]))}">{xml_esc(r.get("text", ""))}</math>'
                                    if r.get("equation_ast")
                                    else xml_esc(strip_illegal_xml_control_chars(str(r.get("text", ""))))
                                )
                                for r in p.get("runs", [])
                            )
                            p_in_cell.append(f"<div>{p_txt}</div>")
                        td_list.append(
                            f'<td data-row="{cell["row"]}" data-col="{cell["col"]}" '
                            f'rowspan="{r_sp}" colspan="{c_sp}" style="{bg_style}">'
                            f'{"".join(p_in_cell)}</td>'
                        )
                    tr_list.append(f'<tr>{"".join(td_list)}</tr>')
                blk_html_list.append(
                    f'<table data-nrows="{blk.get("n_rows", 1)}" data-ncols="{blk.get("n_cols", 1)}">'
                    f'<tbody>{"".join(tr_list)}</tbody></table>'
                )

        for m_obj in sec.media_objects:
            b64 = base64.b64encode(m_obj.raw_bytes).decode("ascii")
            blk_html_list.append(
                f'<figure class="tdir-media"><img src="data:{m_obj.mime_type};base64,{b64}" '
                f'alt="{xml_esc(strip_illegal_xml_control_chars(m_obj.caption_hint))}"/></figure>'
            )

        sec_html_list.append(
            f'<section data-section-uid="{xml_esc(sec.section_uid)}" '
            f'data-ordinal="{sec.ordinal_index or 1}">'
            f'{"".join(blk_html_list)}</section>'
        )

    html_doc = (
        "<!DOCTYPE html>\n"
        '<html lang="ko">\n<head>\n<meta charset="UTF-8"/>\n'
        "<title>T-DIR Document</title>\n"
        "<style>\n"
        f"body {{ font-family: '{xml_esc(font_family)}', sans-serif; font-size: {body_pt}pt; line-height: 1.35; }}\n"
        f"table {{ border-collapse: collapse; width: 100%; font-size: {tbl_pt}pt; margin: 6px 0; }}\n"
        "td, th {{ border: 1px solid #cbd5e1; padding: 4px 6px; vertical-align: middle; }}\n"
        "</style>\n</head>\n<body>\n"
        + "\n".join(sec_html_list)
        + "\n</body>\n</html>\n"
    )

    os.makedirs(os.path.dirname(os.path.abspath(out_html_path)), exist_ok=True)
    with open(out_html_path, "w", encoding="utf-8") as f:
        f.write(html_doc)

    return {
        "out_html_path": out_html_path,
        "bytes": os.path.getsize(out_html_path),
        "sections_count": len(sections),
    }


class _TdirHtmlDOMParser(HTMLParser):
    def __init__(self, target_width_hu: int = 46200) -> None:
        super().__init__()
        self.target_width_hu = target_width_hu
        self.sections: Dict[str, SectionIR] = {}
        self.cur_sec_uid = "SEC_01"
        self.cur_heading = "HTML Section 1"
        self.cur_subtitle = ""
        self.cur_ord = 1
        self.cur_blocks: List[dict] = []
        self.cur_media: List[MediaObjectIR] = []

        self.in_style_or_script = False
        self.in_heading = False
        self.in_subtitle = False
        self.in_para = False
        self.in_table = False
        self.in_td = False
        self.in_math_latex: Optional[str] = None
        self.cur_span_bold = False
        self.cur_span_italic = False
        self.cur_span_underline = False
        self.cur_span_color = "000000"

        self.cur_p_jc = "both"
        self.cur_p_runs: List[dict] = []
        self.cur_tbl_nrows = 0
        self.cur_tbl_ncols = 0
        self.cur_tr_idx = -1
        self.cur_td_col_cursor = 0
        self.cur_tbl_cells: List[dict] = []
        self.cur_td_meta: Optional[dict] = None
        self.cur_td_runs: List[dict] = []
        self.occupied_grid: set = set()
        self.heading_buf: List[str] = []
        self.subtitle_buf: List[str] = []

    def flush_section(self) -> None:
        if not self.cur_blocks and not self.cur_media:
            return
        _, n_c, sha = extract_ir_non_ws_text(self.cur_blocks)
        self.sections[self.cur_sec_uid] = SectionIR(
            section_uid=self.cur_sec_uid,
            heading_text=self.cur_heading,
            subtitle_text=self.cur_subtitle,
            blocks=self.cur_blocks,
            media_objects=self.cur_media,
            non_ws_char_count=n_c,
            text_sha256=sha,
            ordinal_index=self.cur_ord,
            section_role="BODY",
        )
        self.cur_blocks = []
        self.cur_media = []
        self.cur_subtitle = ""

    def handle_starttag(self, tag: str, attrs: List[Tuple[str, Optional[str]]]) -> None:
        attr_map = {k.lower(): (v or "") for k, v in attrs}
        t = tag.lower()
        if t in ("style", "script", "title"):
            self.in_style_or_script = True
            return
        if t == "section":
            self.flush_section()
            self.cur_sec_uid = attr_map.get("data-section-uid") or f"SEC_{len(self.sections) + 1:02d}"
            ord_s = attr_map.get("data-ordinal", "")
            self.cur_ord = int(ord_s) if ord_s.isdigit() else (len(self.sections) + 1)
            return
        if t in ("h1", "h2") and "tdir-sec-heading" in attr_map.get("class", ""):
            self.in_heading = True
            self.heading_buf = []
            return
        if t == "div" and "tdir-sec-subtitle" in attr_map.get("class", ""):
            self.in_subtitle = True
            self.subtitle_buf = []
            return
        if t in ("h1", "h2", "h3") and not self.in_table:
            self.flush_section()
            self.cur_sec_uid = f"SEC_{len(self.sections) + 1:02d}"
            self.in_heading = True
            self.heading_buf = []
            return
        if t == "p" and not self.in_td:
            self.in_para = True
            self.cur_p_runs = []
            st = attr_map.get("style", "")
            self.cur_p_jc = "center" if "center" in st else ("right" if "right" in st else ("left" if "left" in st else "both"))
            return
        if t == "span":
            st = attr_map.get("style", "").lower()
            self.cur_span_bold = "bold" in st
            self.cur_span_italic = "italic" in st
            self.cur_span_underline = "underline" in st
            m_col = re.search(r"color:\s*#([0-9a-fA-F]{6})", st)
            self.cur_span_color = m_col.group(1).upper() if m_col else "000000"
            return
        if t in ("b", "strong"):
            self.cur_span_bold = True
            return
        if t in ("i", "em"):
            self.cur_span_italic = True
            return
        if t == "math":
            self.in_math_latex = attr_map.get("data-latex", "")
            return
        if t == "table":
            self.in_table = True
            self.cur_tbl_nrows = int(attr_map.get("data-nrows", "0") or 0)
            self.cur_tbl_ncols = int(attr_map.get("data-ncols", "0") or 0)
            self.cur_tr_idx = -1
            self.cur_tbl_cells = []
            self.occupied_grid = set()
            return
        if t == "tr" and self.in_table:
            self.cur_tr_idx += 1
            self.cur_td_col_cursor = 0
            return
        if t in ("td", "th") and self.in_table:
            self.in_td = True
            self.cur_td_runs = []
            r_sp = int(attr_map.get("rowspan", "1") or 1)
            c_sp = int(attr_map.get("colspan", "1") or 1)
            if "data-row" in attr_map and "data-col" in attr_map:
                r_addr = int(attr_map["data-row"])
                c_addr = int(attr_map["data-col"])
            else:
                r_addr = max(0, self.cur_tr_idx)
                while (r_addr, self.cur_td_col_cursor) in self.occupied_grid:
                    self.cur_td_col_cursor += 1
                c_addr = self.cur_td_col_cursor
                for dr in range(r_sp):
                    for dc in range(c_sp):
                        self.occupied_grid.add((r_addr + dr, c_addr + dc))
                self.cur_td_col_cursor += c_sp
            st = attr_map.get("style", "")
            m_bg = re.search(r"background-color:\s*#([0-9a-fA-F]{6})", st)
            fill_hex = m_bg.group(1).upper() if m_bg else None
            self.cur_td_meta = {
                "row": r_addr,
                "col": c_addr,
                "row_span": r_sp,
                "col_span": c_sp,
                "fill_hex": fill_hex,
            }
            return
        if t == "img":
            src = attr_map.get("src", "")
            alt = attr_map.get("alt", "")
            if src.startswith("data:") and ";base64," in src:
                header_part, b64_part = src.split(";base64,", 1)
                mime_t = header_part[5:] or "image/png"
                raw_bytes = base64.b64decode(b64_part)
                self.cur_media.append(
                    MediaObjectIR(
                        media_uid=f"{self.cur_sec_uid}_HTML_IMG_{len(self.cur_media) + 1}",
                        rel_id=f"html_img_{len(self.cur_media) + 1}",
                        filename=f"html_img_{len(self.cur_media) + 1}.png",
                        mime_type=mime_t,
                        raw_bytes=raw_bytes,
                        caption_hint=alt,
                    )
                )

    def handle_endtag(self, tag: str) -> None:
        t = tag.lower()
        if t in ("style", "script", "title"):
            self.in_style_or_script = False
            return
        if t in ("h1", "h2", "h3") and self.in_heading:
            self.in_heading = False
            self.cur_heading = "".join(self.heading_buf).strip() or self.cur_sec_uid
            self.cur_ord = extract_ordinal_index(self.cur_heading) or self.cur_ord
            return
        if t == "div" and self.in_subtitle:
            self.in_subtitle = False
            self.cur_subtitle = "".join(self.subtitle_buf).strip()
            return
        if t == "span":
            self.cur_span_bold = False
            self.cur_span_italic = False
            self.cur_span_underline = False
            self.cur_span_color = "000000"
            return
        if t in ("b", "strong"):
            self.cur_span_bold = False
            return
        if t in ("i", "em"):
            self.cur_span_italic = False
            return
        if t == "math":
            self.in_math_latex = None
            return
        if t == "p" and self.in_para and not self.in_td:
            self.in_para = False
            if self.cur_p_runs:
                self.cur_blocks.append(
                    {
                        "type": "p",
                        "jc": self.cur_p_jc,
                        "left_tw": 0,
                        "right_tw": 0,
                        "indent_tw": 0,
                        "before_tw": 0,
                        "after_tw": 8,
                        "runs": self.cur_p_runs,
                        "r_embed": None,
                    }
                )
            self.cur_p_runs = []
            return
        if t in ("td", "th") and self.in_td and self.cur_td_meta is not None:
            self.in_td = False
            meta = self.cur_td_meta
            self.cur_tbl_cells.append(
                {
                    "row": meta["row"],
                    "col": meta["col"],
                    "row_span": meta["row_span"],
                    "col_span": meta["col_span"],
                    "width_hu": 4000,
                    "height_hu": 900 * meta["row_span"],
                    "mar": (140, 140, 35, 35),
                    "borders": {s: (1, 1, (0x94, 0xA3, 0xB8)) for s in ("left", "right", "top", "bottom")},
                    "fill_hex": meta["fill_hex"],
                    "paras": [
                        {
                            "type": "p",
                            "jc": "left",
                            "left_tw": 0,
                            "right_tw": 0,
                            "indent_tw": 0,
                            "before_tw": 0,
                            "after_tw": 0,
                            "runs": self.cur_td_runs
                            if self.cur_td_runs
                            else [{"text": "", "bold": False, "italic": False, "color": "000000"}],
                            "r_embed": None,
                        }
                    ],
                }
            )
            self.cur_td_meta = None
            self.cur_td_runs = []
            return
        if t == "table" and self.in_table:
            self.in_table = False
            if self.cur_tbl_cells:
                n_r = self.cur_tbl_nrows or (max(c["row"] + c["row_span"] for c in self.cur_tbl_cells))
                n_c = self.cur_tbl_ncols or (max(c["col"] + c["col_span"] for c in self.cur_tbl_cells))
                w_col = self.target_width_hu // max(1, n_c)
                col_w = [w_col] * n_c
                col_w[-1] += self.target_width_hu - sum(col_w)
                rows_cnt_map: Dict[int, int] = {}
                for c in self.cur_tbl_cells:
                    c["width_hu"] = sum(col_w[c["col"] : c["col"] + c["col_span"]])
                    rows_cnt_map[c["row"]] = rows_cnt_map.get(c["row"], 0) + 1
                row_cols = [rows_cnt_map.get(r_i, n_c) for r_i in range(n_r)]
                self.cur_blocks.append(
                    {
                        "type": "tbl",
                        "n_rows": n_r,
                        "n_cols": n_c,
                        "col_widths_hu": col_w,
                        "row_cols": row_cols,
                        "cells": self.cur_tbl_cells,
                        "width_hu": self.target_width_hu,
                        "height_hu": n_r * 900,
                    }
                )
            self.cur_tbl_cells = []
            return
        if t == "section":
            self.flush_section()

    def handle_data(self, data: str) -> None:
        if self.in_style_or_script:
            return
        cleaned = strip_illegal_xml_control_chars(data)
        if not cleaned:
            return
        if self.in_heading:
            self.heading_buf.append(cleaned)
            return
        if self.in_subtitle:
            self.subtitle_buf.append(cleaned)
            return
        if self.in_td or self.in_para:
            if not cleaned.strip() and not (self.cur_td_runs if self.in_td else self.cur_p_runs):
                return
            r_dict: dict = {
                "text": cleaned,
                "bold": self.cur_span_bold,
                "italic": self.cur_span_italic,
                "underline": self.cur_span_underline,
                "color": self.cur_span_color,
            }
            if self.in_math_latex:
                eq_ast = parse_latex_to_equation_ast(self.in_math_latex)
                r_dict["equation_ast"] = eq_ast
                r_dict["text"] = compile_equation_ast_to_linear_text(eq_ast)
            if self.in_td:
                self.cur_td_runs.append(r_dict)
            else:
                self.cur_p_runs.append(r_dict)


def lift_html_to_ir(html_path: str, target_width_hu: int = 46200) -> Dict[str, SectionIR]:
    with open(html_path, "r", encoding="utf-8") as f:
        raw_html = f.read()
    parser = _TdirHtmlDOMParser(target_width_hu=target_width_hu)
    parser.feed(raw_html)
    parser.flush_section()
    return parser.sections


def extract_html_body_text(html_path: str) -> str:
    sections = lift_html_to_ir(html_path)
    all_blocks = [b for s in sections.values() for b in s.blocks]
    norm_txt, _, _ = extract_ir_non_ws_text(all_blocks)
    return norm_txt


# ==============================================================================
# 4. Deterministic Pure-Python Multi-Page PDF 1.4 Projector & Verifier ([P3-02])
# ==============================================================================


def _pdf_utf16be_hex(text: str) -> str:
    cleaned = strip_illegal_xml_control_chars(text)
    u16_be = b"\xfe\xff" + cleaned.encode("utf-16-be")
    return "<" + u16_be.hex().upper() + ">"


def project_ir_to_pdf(
    sections: Dict[str, SectionIR],
    out_pdf_path: str,
    target_spec: Optional[TargetDocumentSpecIR] = None,
    slot_blocks_map: Optional[Dict[str, Tuple[List[dict], ElasticLayoutState]]] = None,
    attachments: Optional[List[Tuple[bytes, str]]] = None,
    expected_total_pages: Optional[int] = None,
    clear_unmapped_slots: bool = False,
    preserve_outer_tables: bool = True,
) -> dict:
    """
    Deterministic pure-Python multi-page PDF 1.4 renderer ([P3-02]).
    Computes exact box-model page geometry in PDF points (1 pt = 100 HU) from `target_spec`,
    paginates preamble cover, section slots, and attachment media grids using the solved
    `ElasticLayoutState`, and encodes all CJK/Unicode text losslessly as UTF-16BE (`<FEFF...>`)
    strings inside `/Contents` streams.
    """
    eff_spec = target_spec or TargetDocumentSpecIR(template_path="", target_format="PDF")
    pw_pt = round((eff_spec.page_width_hu if eff_spec.page_width_hu > 0 else 59528) / 100.0, 2)
    ph_pt = round((eff_spec.page_height_hu if eff_spec.page_height_hu > 0 else 84188) / 100.0, 2)
    ml_pt = round((eff_spec.margin_left_hu if eff_spec.margin_left_hu > 0 else 5669) / 100.0, 2)
    mr_pt = round((eff_spec.margin_right_hu if eff_spec.margin_right_hu > 0 else 5669) / 100.0, 2)
    mt_pt = round((eff_spec.margin_top_hu if eff_spec.margin_top_hu > 0 else 4251) / 100.0, 2)
    mb_pt = round((eff_spec.margin_bottom_hu if eff_spec.margin_bottom_hu > 0 else 4252) / 100.0, 2)

    usable_w_pt = max(180.0, pw_pt - ml_pt - mr_pt)
    usable_h_pt = max(240.0, ph_pt - mt_pt - mb_pt)

    pages_ops: List[List[str]] = []

    def _new_page() -> List[str]:
        ops: List[str] = [
            f"0.22 0.24 0.52 RG 0.8 w {ml_pt:.2f} {mb_pt:.2f} {usable_w_pt:.2f} {usable_h_pt:.2f} re S"
        ]
        pages_ops.append(ops)
        return ops

    def _emit_text_line(ops: List[str], x_pt: float, y_pt: float, font_sz_pt: float, txt: str) -> None:
        if not txt:
            return
        hex_str = _pdf_utf16be_hex(txt)
        ops.append(f"BT /F2 {font_sz_pt:.2f} Tf {x_pt:.2f} {y_pt:.2f} Td {hex_str} Tj ET")

    # 1. Preamble Cover Page (if preamble_blocks exist)
    if eff_spec.preamble_blocks and preserve_outer_tables:
        ops_cov = _new_page()
        y_cur = ph_pt - mt_pt - 60.0
        for pb in eff_spec.preamble_blocks:
            pb_txt = str(pb.get("text", "")).strip()
            if pb_txt:
                pb_pt = round(pb.get("height_hu", 1400) / 100.0, 2)
                _emit_text_line(ops_cov, ml_pt + 24.0, y_cur, pb_pt, pb_txt)
                y_cur -= pb_pt * 1.8

    # 2. Section Slots Pagination
    default_st = ElasticLayoutState()
    eff_slot_map: Dict[str, Tuple[List[dict], ElasticLayoutState]] = dict(slot_blocks_map or {})
    if not eff_slot_map and sections:
        for idx_s, sec in enumerate(sections.values()):
            sl_uid = eff_spec.section_slots[idx_s].slot_uid if idx_s < len(eff_spec.section_slots) else sec.section_uid
            eff_slot_map[sl_uid] = (sec.blocks, default_st)

    slots_to_render = eff_spec.section_slots
    for s_slot in slots_to_render:
        if (clear_unmapped_slots or not preserve_outer_tables) and s_slot.slot_uid not in eff_slot_map:
            continue
        if s_slot.slot_uid not in eff_slot_map:
            ops_empty = _new_page()
            y_e = ph_pt - mt_pt - 18.0
            if s_slot.discovered_title_text:
                _emit_text_line(ops_empty, ml_pt + 8.0, y_e, 11.0, s_slot.discovered_title_text)
            continue

        blocks, el_state = eff_slot_map.get(s_slot.slot_uid, ([], default_st))
        body_sz_pt = round((el_state.body_font_size_hu or eff_spec.body_font_size_hu or 1000) / 100.0, 2)
        tbl_sz_pt = round((el_state.table_font_size_hu or eff_spec.table_font_size_hu or 950) / 100.0, 2)
        body_lh_pt = max(9.0, body_sz_pt * (el_state.body_line_spacing_pct / 100.0))
        tbl_lh_pt = max(8.5, tbl_sz_pt * (el_state.table_line_spacing_pct / 100.0))

        slot_max_pages = max(1, int(round(s_slot.max_pages))) if s_slot.max_pages > 0 else 2
        slot_pages_used = 1
        ops_cur = _new_page()
        y_top = ph_pt - mt_pt - 18.0
        y_bot = mb_pt + 18.0
        y_cur = y_top

        if s_slot.discovered_title_text:
            _emit_text_line(ops_cur, ml_pt + 8.0, y_cur, body_sz_pt, s_slot.discovered_title_text)
            y_cur -= body_lh_pt * 1.15
        if s_slot.discovered_subtitle_text:
            _emit_text_line(ops_cur, ml_pt + 8.0, y_cur, tbl_sz_pt, s_slot.discovered_subtitle_text)
            y_cur -= tbl_lh_pt * 1.15

        chars_per_line = max(25.0, (usable_w_pt - 16.0) / max(4.0, body_sz_pt))

        for blk in blocks:
            if blk["type"] == "p":
                p_txt = "".join(
                    (
                        r.get("text", "")
                        or (compile_equation_ast_to_linear_text(r["equation_ast"]) if r.get("equation_ast") else "")
                    )
                    for r in blk.get("runs", [])
                )
                if not p_txt:
                    continue
                w_units = compute_weighted_text_units(p_txt)
                n_lines = max(1, math.ceil(w_units / chars_per_line))
                needed_h = n_lines * body_lh_pt + 2.0
                if y_cur - needed_h < y_bot and slot_pages_used < slot_max_pages:
                    ops_cur = _new_page()
                    slot_pages_used += 1
                    y_cur = y_top
                _emit_text_line(ops_cur, ml_pt + 8.0, max(y_bot, y_cur), body_sz_pt, p_txt)
                y_cur -= needed_h
            elif blk["type"] == "tbl":
                rows_by_r: Dict[int, List[dict]] = {}
                for c in blk.get("cells", []):
                    rows_by_r.setdefault(c["row"], []).append(c)
                for r_i in range(blk.get("n_rows", 1)):
                    row_cells = sorted(rows_by_r.get(r_i, []), key=lambda x: x["col"])
                    row_str = "   ".join(
                        " ".join(
                            "".join(
                                (
                                    r.get("text", "")
                                    or (
                                        compile_equation_ast_to_linear_text(r["equation_ast"])
                                        if r.get("equation_ast")
                                        else ""
                                    )
                                )
                                for r in p.get("runs", [])
                            )
                            for p in rc.get("paras", [])
                        )
                        for rc in row_cells
                    )
                    w_units = compute_weighted_text_units(row_str)
                    n_lines = max(1, math.ceil(w_units / max(28.0, (usable_w_pt - 16.0) / max(4.0, tbl_sz_pt))))
                    row_h = n_lines * tbl_lh_pt + 3.0
                    if y_cur - row_h < y_bot and slot_pages_used < slot_max_pages:
                        ops_cur = _new_page()
                        slot_pages_used += 1
                        y_cur = y_top
                    _emit_text_line(ops_cur, ml_pt + 8.0, max(y_bot, y_cur), tbl_sz_pt, row_str)
                    y_cur -= row_h

        while expected_total_pages is not None and slot_pages_used < slot_max_pages:
            _new_page()
            slot_pages_used += 1

    # 3. Attachment Media Grid Page (if attachment_slots or attachments exist)
    if (eff_spec.attachment_slots and preserve_outer_tables) or attachments:
        ops_att = _new_page()
        y_cur = ph_pt - mt_pt - 24.0
        for ah_p in eff_spec.attachment_header_paras:
            ah_txt = str(ah_p.get("text", "")).strip()
            if ah_txt:
                _emit_text_line(ops_att, ml_pt + 8.0, y_cur, 11.0, ah_txt)
                y_cur -= 18.0
        for _, cap_s in (attachments or []):
            if cap_s:
                _emit_text_line(ops_att, ml_pt + 12.0, y_cur, 9.5, cap_s)
                y_cur -= 15.0

    if not pages_ops:
        _new_page()

    # Assemble PDF 1.4 objects & cross-reference table
    n_pages = len(pages_ops)
    first_page_obj_id = 5
    page_obj_ids = [first_page_obj_id + i * 2 for i in range(n_pages)]
    content_obj_ids = [first_page_obj_id + i * 2 + 1 for i in range(n_pages)]

    objects: List[Tuple[int, bytes]] = []
    objects.append((1, b"<< /Type /Catalog /Pages 2 0 R >>"))
    kids_str = " ".join(f"{pid} 0 R" for pid in page_obj_ids)
    objects.append((2, f"<< /Type /Pages /Kids [ {kids_str} ] /Count {n_pages} >>".encode("ascii")))
    objects.append((3, b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"))
    objects.append(
        (
            4,
            b"<< /Type /Font /Subtype /Type0 /BaseFont /Dotum /Encoding /Identity-H "
            b"/DescendantFonts [ << /Type /Font /Subtype /CIDFontType2 /BaseFont /Dotum "
            b"/CIDSystemInfo << /Registry (Adobe) /Ordering (Korea1) /Supplement 2 >> /DW 1000 >> ] >>",
        )
    )

    for idx_p, ops in enumerate(pages_ops):
        p_oid = page_obj_ids[idx_p]
        c_oid = content_obj_ids[idx_p]
        stream_bytes = "\n".join(ops).encode("ascii")
        page_dict = (
            f"<< /Type /Page /Parent 2 0 R /MediaBox [ 0 0 {pw_pt:.2f} {ph_pt:.2f} ] "
            f"/Resources << /Font << /F1 3 0 R /F2 4 0 R >> >> /Contents {c_oid} 0 R >>"
        ).encode("ascii")
        content_dict = f"<< /Length {len(stream_bytes)} >>\nstream\n".encode("ascii") + stream_bytes + b"\nendstream"
        objects.append((p_oid, page_dict))
        objects.append((c_oid, content_dict))

    pdf_buf = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
    offsets: Dict[int, int] = {}
    for oid, obj_body in objects:
        offsets[oid] = len(pdf_buf)
        pdf_buf.extend(f"{oid} 0 obj\n".encode("ascii"))
        pdf_buf.extend(obj_body)
        pdf_buf.extend(b"\nendobj\n")

    xref_offset = len(pdf_buf)
    total_objs = len(objects) + 1
    pdf_buf.extend(f"xref\n0 {total_objs}\n0000000000 65535 f \n".encode("ascii"))
    for oid in range(1, total_objs):
        pdf_buf.extend(f"{offsets[oid]:010d} 00000 n \n".encode("ascii"))
    pdf_buf.extend(
        f"trailer\n<< /Size {total_objs} /Root 1 0 R >>\nstartxref\n{xref_offset}\n%%EOF\n".encode("ascii")
    )

    os.makedirs(os.path.dirname(os.path.abspath(out_pdf_path)), exist_ok=True)
    with open(out_pdf_path, "wb") as f:
        f.write(pdf_buf)

    return {
        "out_pdf_path": out_pdf_path,
        "bytes": len(pdf_buf),
        "actual_pdf_pages": n_pages,
        "page_count": n_pages,
        "page_width_pt": pw_pt,
        "page_height_pt": ph_pt,
    }


def inspect_pdf_file(pdf_path: str) -> dict:
    """
    Extracts exact page count (`/Type /Page` and `/Count N`) and UTF-16BE (`<FEFF...>`)
    text streams from a PDF 1.4 file ([P3-02]).
    """
    with open(pdf_path, "rb") as f:
        raw = f.read()
    page_objs = re.findall(rb"/Type\s*/Page\b(?!s)", raw)
    if page_objs:
        page_count = len(page_objs)
    else:
        m_cnt = re.search(rb"/Type\s*/Pages\b[^>]*?/Count\s+(\d+)", raw, re.DOTALL)
        if not m_cnt:
            m_cnt = re.search(rb"/Count\s+(\d+)", raw)
        page_count = int(m_cnt.group(1)) if m_cnt else 0

    decoded_chunks: List[str] = []
    for m_hex in re.finditer(rb"<FEFF([0-9A-Fa-f]+)>", raw):
        hex_payload = m_hex.group(1).decode("ascii")
        if len(hex_payload) % 2 == 0:
            decoded_chunks.append(bytes.fromhex(hex_payload).decode("utf-16-be", errors="replace"))

    extracted_text = "\n".join(decoded_chunks)
    non_ws = re.sub(r"\s+", "", extracted_text)
    return {
        "pdf_path": pdf_path,
        "out_pdf_path": pdf_path,
        "bytes": len(raw),
        "actual_pdf_pages": page_count,
        "page_count": page_count,
        "extracted_text": extracted_text,
        "non_ws_chars": len(non_ws),
    }

