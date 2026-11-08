"""
Markup (.md, .html, .txt) and Tabular (.csv, .tsv, .xlsx) Lifters & Projectors for T-DIR v3.0.
Enables N-to-M conversion across Markdown, HTML, CSV/TSV, XLSX, and HWP/HWPX/DOCX.
"""

import os
import re
import csv
from html.parser import HTMLParser
from typing import List, Dict

from .ir_schema import SectionIR, extract_ir_non_ws_text, extract_ordinal_index


def lift_markdown_to_ir(md_path: str, target_width_hu: int = 46200) -> Dict[str, SectionIR]:
    with open(md_path, "r", encoding="utf-8") as f:
        lines = f.read().splitlines()

    sections: Dict[str, SectionIR] = {}
    cur_sec_id = "SEC_1"
    cur_heading = "Markdown Root"
    cur_blocks: List[dict] = []
    table_buffer: List[List[str]] = []

    def flush_table():
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

    def flush_sec():
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
                        "runs": [{"text": line, "bold": False, "italic": False, "color": "000000"}],
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
            cur_heading = stripped[2:].strip()
            ord_idx = extract_ordinal_index(cur_heading)
            if cur_heading.startswith(("Q", "q")) and ord_idx is not None:
                cur_sec_id = f"Q{ord_idx}"
            else:
                cur_sec_id = f"SEC_{len(sections) + 1}"
            continue
        is_bold_heading = stripped.startswith("##")
        clean_txt = stripped.lstrip("#").strip() if is_bold_heading else stripped
        cur_blocks.append(
            {
                "type": "p",
                "jc": "both",
                "left_tw": 0,
                "right_tw": 0,
                "indent_tw": 0,
                "before_tw": 0,
                "after_tw": 8,
                "runs": [{"text": clean_txt, "bold": is_bold_heading, "italic": False, "color": "000000"}],
                "r_embed": None,
            }
        )

    flush_sec()
    return sections


def project_ir_to_markdown(sections: Dict[str, SectionIR], out_md_path: str) -> dict:
    out_lines: List[str] = []
    for sec in sections.values():
        out_lines.append(f"# {sec.heading_text}")
        out_lines.append("")
        for blk in sec.blocks:
            if blk["type"] == "p":
                txt = "".join(r["text"] for r in blk["runs"])
                out_lines.append(txt)
                out_lines.append("")
            elif blk["type"] == "tbl":
                rows_by_r: Dict[int, List[dict]] = {}
                for c in blk["cells"]:
                    rows_by_r.setdefault(c["row"], []).append(c)
                for r_i in range(blk["n_rows"]):
                    row_cells = sorted(rows_by_r.get(r_i, []), key=lambda x: x["col"])
                    cell_strs = [
                        " ".join("".join(r["text"] for r in p["runs"]) for p in rc["paras"]) for rc in row_cells
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
            val = row_vals[c_i].strip() if c_i < len(row_vals) else ""
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
        "TABULAR_MAIN": SectionIR(
            section_uid="TABULAR_MAIN",
            heading_text=os.path.basename(csv_path),
            blocks=[tbl_blk],
            non_ws_char_count=n_c,
            text_sha256=sha,
        )
    }
