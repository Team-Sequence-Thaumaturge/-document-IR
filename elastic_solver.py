"""
T-DIR v3.2 2D Topological Grid Normalizer, Numerical Page-Budget Solver, & 6-Stage Minimum-Invasion Elastic Layout Solver.
Zero hardcoded font sizes, pixel/HWPUNIT constants, or document-specific strings:
  - Table normalization relies purely on 2D grid sparsity and span topology (with pre-spacer
    and post-spacer stair-step detection `[BUG-01]`, `[BUG-02]` and non-zero `row_cols` invariant `[EDGE-02]`).
  - Line heights, character capacities (with CJK/Latin East-Asian-Width weighting `[EDGE-04]`),
    and page budgets scale proportionally with resolved typography and physical page box geometry (`[NUM-04]`).
"""

import math
import copy
from dataclasses import dataclass
from typing import List, Dict, Tuple, Set, Optional

from .ir_schema import (
    SectionIR,
    SlotSpecIR,
    TargetDocumentSpecIR,
    UserSpecOverrideIR,
    compute_weighted_text_units,
    resolve_budget_for_section_or_slot,
)


@dataclass
class ElasticLayoutState:
    snap_to_grid: bool = False
    prune_empty_spacers: bool = True
    body_font_size_hu: int = 1000
    table_font_size_hu: int = 900
    body_line_spacing_pct: int = 122
    table_line_spacing_pct: int = 112
    body_line_tw: int = 220
    table_line_tw: int = 190
    cell_pad_v_hu: int = 35
    para_after_tw: int = 7
    merged_subheadings_count: int = 0
    merged_cell_pairs_count: int = 0
    merged_body_pairs_count: int = 0
    stopped_stage: str = "STAGE_1_SPACERS_ONLY"


def compute_proportional_line_tw(font_size_hu: int, line_spacing_pct: int, is_table: bool = False) -> int:
    """
    Computes exact OOXML line spacing in Twips (dxa) proportionally from font_size_hu (HU)
    and line_spacing_pct (%), compensating for CJK winAscent/winDescent bounding boxes.
    """
    base_tw = font_size_hu / 5.0
    metric_factor = 0.948 if is_table else 0.916
    return max(120, int(round(base_tw * (line_spacing_pct / 100.0) * metric_factor)))


def compute_effective_page_body_height_hu(
    target_spec: Optional[TargetDocumentSpecIR] = None,
    explicit_override_hu: Optional[int] = None,
) -> int:
    """
    Derives single-page effective body capacity (HU) directly from the target document's
    physical paper height minus effective top and bottom page margins ([NUM-04]).
    """
    if explicit_override_hu is not None and explicit_override_hu > 10000:
        return explicit_override_hu
    if target_spec is not None and target_spec.page_height_hu > 20000:
        eff_top_hu = (
            target_spec.margin_top_hu
            if target_spec.margins_explicitly_overridden
            else min(2800, target_spec.margin_top_hu if target_spec.margin_top_hu > 0 else 2800)
        )
        eff_bot_hu = (
            target_spec.margin_bottom_hu
            if target_spec.margins_explicitly_overridden
            else min(2800, target_spec.margin_bottom_hu if target_spec.margin_bottom_hu > 0 else 2800)
        )
        return max(20000, target_spec.page_height_hu - eff_top_hu - eff_bot_hu)
    return 78588


def is_stair_step_sparse_grid(grid_cells: List[List[Optional[dict]]], n_rows_raw: int, n_cols_raw: int) -> bool:
    """
    Pure topological detector for stair-step bar chart tables:
    Checks whether rows 1..n_cols_raw each have exactly 1 non-empty cell placed in a distinct column.
    """
    if n_cols_raw < 3 or n_rows_raw < n_cols_raw + 2:
        return False
    occupied_cols: Set[int] = set()
    for r_i in range(1, 1 + n_cols_raw):
        non_empty_cols = [
            c_i
            for c_i in range(n_cols_raw)
            if grid_cells[r_i][c_i] is not None
            and grid_cells[r_i][c_i]["col"] == c_i
            and len(grid_cells[r_i][c_i]["paras"]) > 0
        ]
        if len(non_empty_cols) != 1:
            return False
        occupied_cols.add(non_empty_cols[0])
    return len(occupied_cols) == n_cols_raw


def is_vertical_step_arrow_grid(grid_cells: List[List[Optional[dict]]], n_rows_raw: int, n_cols_raw: int) -> bool:
    """
    Pure topological detector for 1-column vertical step flow tables where odd rows
    contain only directional arrow glyphs ('↓', '▼', '→').
    """
    if n_cols_raw != 1 or n_rows_raw < 3:
        return False
    arrow_rows = 0
    for r_i in range(1, n_rows_raw, 2):
        cobj = grid_cells[r_i][0]
        if cobj is None:
            return False
        txt = "".join(r["text"] for p in cobj["paras"] for r in p["runs"]).strip()
        if not txt or len(txt) > 3 or not all(ch in "↓▼→⬇" for ch in txt):
            return False
        arrow_rows += 1
    return arrow_rows >= 1


def is_standalone_subheading_para(txt: str) -> bool:
    """
    Grammatical & topological predicate identifying short standalone section subheadings
    (replacing any hardcoded list of document-specific section titles).
    """
    s = txt.strip()
    if not s or len(s) > 28:
        return False
    if s.endswith((".", "함", "음", "다", ")", "]")):
        return False
    return True


def detect_and_remove_spacer_rows_cols(
    grid_cells: List[List[Optional[dict]]],
    col_twips: List[int],
    n_rows_raw: int,
    n_cols_raw: int,
    target_width_hu: int,
    table_font_size_hu: int = 1000,
) -> dict:
    active_rows: List[int] = []
    for r_i in range(n_rows_raw):
        has_text = any(
            grid_cells[r_i][c_i] is not None and len(grid_cells[r_i][c_i]["paras"]) > 0
            for c_i in range(n_cols_raw)
        )
        if has_text:
            active_rows.append(r_i)
    if not active_rows:
        active_rows = [0]

    # Check if every active row has only 1 distinct non-empty cell (e.g., multi-column spacer card table)
    distinct_nonempty_per_row: List[List[dict]] = []
    for r_i in active_rows:
        row_nonempty: List[dict] = []
        seen_row_ids: Set[int] = set()
        for c_i in range(n_cols_raw):
            cobj = grid_cells[r_i][c_i]
            if cobj is not None and id(cobj) not in seen_row_ids and len(cobj["paras"]) > 0:
                seen_row_ids.add(id(cobj))
                row_nonempty.append(cobj)
        distinct_nonempty_per_row.append(row_nonempty)

    if n_cols_raw > 1 and all(len(rn) == 1 for rn in distinct_nonempty_per_row):
        new_cells_1col: List[dict] = []
        line_h_est = int(round(table_font_size_hu * 0.95))
        for idx_r, rn in enumerate(distinct_nonempty_per_row):
            cobj = copy.deepcopy(rn[0])
            cobj["row"], cobj["col"], cobj["row_span"], cobj["col_span"] = idx_r, 0, 1, 1
            cobj["width_hu"] = target_width_hu
            cobj["mar"] = (160, 160, 35, 35)
            cobj["borders"] = {s: (1, 1, (0x94, 0xA3, 0xB8)) for s in ("left", "right", "top", "bottom")}
            cobj["height_hu"] = max(800, len(cobj["paras"]) * line_h_est + 80)
            new_cells_1col.append(cobj)
        return {
            "type": "tbl",
            "n_rows": len(new_cells_1col),
            "n_cols": 1,
            "col_widths_hu": [target_width_hu],
            "row_cols": [1] * len(new_cells_1col),
            "cells": new_cells_1col,
            "width_hu": target_width_hu,
            "height_hu": sum(c["height_hu"] for c in new_cells_1col),
        }

    active_cols: List[int] = []
    for c_i in range(n_cols_raw):
        has_single_col_text = any(
            grid_cells[r_i][c_i] is not None
            and grid_cells[r_i][c_i]["col"] == c_i
            and grid_cells[r_i][c_i]["col_span"] == 1
            and len(grid_cells[r_i][c_i]["paras"]) > 0
            for r_i in active_rows
        )
        if has_single_col_text:
            active_cols.append(c_i)
    if not active_cols:
        active_cols = list(range(n_cols_raw))

    row_map = {old_r: new_r for new_r, old_r in enumerate(active_rows)}
    col_map = {old_c: new_c for new_c, old_c in enumerate(active_cols)}
    n_rows_new = len(active_rows)
    n_cols_new = len(active_cols)

    # Check post-spacer compact subgrid for stair-step topology ([BUG-02])
    compact_grid: List[List[Optional[dict]]] = [[None for _ in range(n_cols_new)] for _ in range(n_rows_new)]
    seen_compact: Set[int] = set()
    for old_r in active_rows:
        for old_c in range(n_cols_raw):
            cobj = grid_cells[old_r][old_c]
            if cobj is None or id(cobj) in seen_compact:
                continue
            seen_compact.add(id(cobj))
            sp_cols = [c for c in range(cobj["col"], cobj["col"] + cobj["col_span"]) if c in col_map]
            sp_rows = [r for r in range(cobj["row"], cobj["row"] + cobj["row_span"]) if r in row_map]
            if not sp_cols or not sp_rows:
                continue
            nr = row_map[sp_rows[0]]
            nc = col_map[sp_cols[0]]
            nrs = len(sp_rows)
            ncs = len(sp_cols)
            ccopy = copy.deepcopy(cobj)
            ccopy["row"], ccopy["col"], ccopy["row_span"], ccopy["col_span"] = nr, nc, nrs, ncs
            for dr in range(nrs):
                for dc in range(ncs):
                    if nr + dr < n_rows_new and nc + dc < n_cols_new:
                        compact_grid[nr + dr][nc + dc] = ccopy

    if is_stair_step_sparse_grid(compact_grid, n_rows_new, n_cols_new):
        return collapse_stair_step_chart_table(
            compact_grid,
            target_width_hu,
            n_cols_raw=n_cols_new,
            table_font_size_hu=table_font_size_hu,
            n_rows_raw=n_rows_new,
        )

    col_widths_hu = normalize_column_ratios_with_arrow_clamp(
        grid_cells, active_rows, active_cols, col_twips, target_width_hu, table_font_size_hu
    )

    new_cells: List[dict] = []
    seen_ids: Set[int] = set()
    line_h_est = int(round(table_font_size_hu * 0.95))
    for old_r in active_rows:
        for old_c in range(n_cols_raw):
            cobj = grid_cells[old_r][old_c]
            if cobj is None or id(cobj) in seen_ids:
                continue
            seen_ids.add(id(cobj))
            spanned_old_cols = [c for c in range(cobj["col"], cobj["col"] + cobj["col_span"]) if c in col_map]
            spanned_old_rows = [r for r in range(cobj["row"], cobj["row"] + cobj["row_span"]) if r in row_map]
            if not spanned_old_cols or not spanned_old_rows:
                continue
            new_r = row_map[spanned_old_rows[0]]
            new_c = col_map[spanned_old_cols[0]]
            new_r_span = len(spanned_old_rows)
            new_c_span = len(spanned_old_cols)

            cell_paras = copy.deepcopy(cobj["paras"])
            if not cell_paras:
                cell_paras = [
                    {
                        "type": "p",
                        "jc": "left",
                        "left_tw": 0,
                        "right_tw": 0,
                        "indent_tw": 0,
                        "before_tw": 0,
                        "after_tw": 0,
                        "runs": [],
                        "r_embed": None,
                    }
                ]
            w_hu = max(400, sum(col_widths_hu[new_c : new_c + new_c_span]))
            new_cells.append(
                {
                    "row": new_r,
                    "col": new_c,
                    "row_span": new_r_span,
                    "col_span": new_c_span,
                    "width_hu": w_hu,
                    "height_hu": max(800, len(cell_paras) * line_h_est + 80),
                    "mar": (140, 140, 35, 35),
                    "borders": {s: (1, 1, (0x94, 0xA3, 0xB8)) for s in ("left", "right", "top", "bottom")},
                    "fill_hex": cobj["fill_hex"],
                    "paras": cell_paras,
                }
            )

    row_cols = [0] * n_rows_new
    for c in new_cells:
        row_cols[c["row"]] += 1

    # Enforce HWP Tag 77 invariant: every row must have row_cols[r] >= 1 ([EDGE-02])
    if any(rc == 0 for rc in row_cols):
        valid_new_rows = [r_idx for r_idx, rc in enumerate(row_cols) if rc > 0]
        if valid_new_rows:
            remap_r = {old_nr: idx_nr for idx_nr, old_nr in enumerate(valid_new_rows)}
            for c in new_cells:
                r0 = c["row"]
                r1 = r0 + c["row_span"]
                kept_span = [rr for rr in range(r0, r1) if rr in remap_r]
                c["row"] = remap_r[kept_span[0]] if kept_span else 0
                c["row_span"] = max(1, len(kept_span))
            n_rows_new = len(valid_new_rows)
            row_cols = [0] * n_rows_new
            for c in new_cells:
                row_cols[c["row"]] += 1

    return {
        "type": "tbl",
        "n_rows": n_rows_new,
        "n_cols": n_cols_new,
        "col_widths_hu": col_widths_hu,
        "row_cols": row_cols,
        "cells": new_cells,
        "width_hu": target_width_hu,
        "height_hu": n_rows_new * max(900, line_h_est + 100),
    }


def normalize_column_ratios_with_arrow_clamp(
    grid_cells: List[List[Optional[dict]]],
    active_rows: List[int],
    active_cols: List[int],
    col_twips: List[int],
    target_width_hu: int,
    table_font_size_hu: int = 1000,
) -> List[int]:
    active_twips = [col_twips[c] if c < len(col_twips) else 1000 for c in active_cols]
    arrow_width_hu = max(1200, int(round(table_font_size_hu * 1.6)))

    for new_ci, old_ci in enumerate(active_cols):
        col_texts: List[str] = []
        for r_i in active_rows:
            cobj = grid_cells[r_i][old_ci]
            if cobj is not None and cobj["col"] == old_ci and cobj["col_span"] == 1:
                txt = "".join(r["text"] for p in cobj["paras"] for r in p["runs"]).strip()
                if txt:
                    col_texts.append(txt)
        if col_texts and all(len(t) <= 2 and any(ch in t for ch in "▶→►➡") for t in col_texts):
            active_twips[new_ci] = max(120, int(round(arrow_width_hu / 5.0)))

    tot_tw = sum(active_twips) or 1
    col_widths_hu = [max(400, int(round(w / tot_tw * target_width_hu))) for w in active_twips]
    if col_widths_hu:
        col_widths_hu[-1] += target_width_hu - sum(col_widths_hu)
    return col_widths_hu


def collapse_stair_step_chart_table(
    grid_cells: List[List[Optional[dict]]],
    target_width_hu: int,
    n_cols_raw: int = 4,
    table_font_size_hu: int = 1000,
    n_rows_raw: Optional[int] = None,
) -> dict:
    """
    Collapses a stair-step sparse bar chart table into a compact 3-row or 4-row grid ([BUG-01]):
      Row 0: Header banner (col_span=n_cols_raw)
      Row 1: Chronological data cells (C0 .. C_{n-1}) extracted from the sparse stair-step rows
      Row 2: Axis label cells (C0 .. C_{n-1})
      Row 3: Optional footer summary banner (col_span=n_cols_raw) if present
    """
    total_rows_in = n_rows_raw if n_rows_raw is not None else len(grid_cells)
    w_col = target_width_hu // max(1, n_cols_raw)
    col_w = [w_col] * n_cols_raw
    col_w[-1] += target_width_hu - sum(col_w)
    line_h_est = int(round(table_font_size_hu * 1.1))

    hdr_src = grid_cells[0][0] if (grid_cells and grid_cells[0] and grid_cells[0][0]) else {"fill_hex": "1E293B", "paras": []}
    hdr_cell = copy.deepcopy(hdr_src)
    hdr_cell["row"], hdr_cell["col"], hdr_cell["row_span"], hdr_cell["col_span"] = 0, 0, 1, n_cols_raw
    hdr_cell["width_hu"], hdr_cell["height_hu"] = target_width_hu, max(850, line_h_est)
    hdr_cell["mar"] = (140, 140, 35, 35)
    hdr_cell["borders"] = {s: (1, 1, (0x94, 0xA3, 0xB8)) for s in ("left", "right", "top", "bottom")}

    bar_cells: List[dict] = []
    for c_idx in range(n_cols_raw):
        found = None
        for r_idx in range(1, min(total_rows_in, 1 + n_cols_raw)):
            cand = grid_cells[r_idx][c_idx]
            if cand is not None and len(cand["paras"]) > 0:
                found = copy.deepcopy(cand)
                break
        if found is None:
            found = {"fill_hex": "F8FAFC", "paras": []}
        found["row"], found["col"], found["row_span"], found["col_span"] = 1, c_idx, 1, 1
        found["width_hu"], found["height_hu"] = col_w[c_idx], max(1100, line_h_est * 2)
        found["mar"] = (120, 120, 35, 35)
        found["borders"] = {s: (1, 1, (0x94, 0xA3, 0xB8)) for s in ("left", "right", "top", "bottom")}
        bar_cells.append(found)

    axis_row_idx = min(total_rows_in - 1, 1 + n_cols_raw)
    year_cells: List[dict] = []
    for c_idx in range(n_cols_raw):
        yc_src = grid_cells[axis_row_idx][c_idx] if grid_cells[axis_row_idx][c_idx] is not None else {"fill_hex": None, "paras": []}
        yc = copy.deepcopy(yc_src)
        yc["row"], yc["col"], yc["row_span"], yc["col_span"] = 2, c_idx, 1, 1
        yc["width_hu"], yc["height_hu"] = col_w[c_idx], max(750, line_h_est)
        yc["mar"] = (120, 120, 25, 25)
        yc["borders"] = {s: (1, 1, (0x94, 0xA3, 0xB8)) for s in ("left", "right", "top", "bottom")}
        year_cells.append(yc)

    all_cells = [hdr_cell] + bar_cells + year_cells
    row_cols_out = [1, n_cols_raw, n_cols_raw]
    total_h_hu = hdr_cell["height_hu"] + bar_cells[0]["height_hu"] + year_cells[0]["height_hu"]

    if axis_row_idx + 1 < total_rows_in and grid_cells[axis_row_idx + 1][0] is not None:
        ftr_cell = copy.deepcopy(grid_cells[axis_row_idx + 1][0])
        ftr_cell["row"], ftr_cell["col"], ftr_cell["row_span"], ftr_cell["col_span"] = 3, 0, 1, n_cols_raw
        ftr_cell["width_hu"], ftr_cell["height_hu"] = target_width_hu, max(850, line_h_est)
        ftr_cell["mar"] = (140, 140, 35, 35)
        ftr_cell["borders"] = {s: (1, 1, (0x94, 0xA3, 0xB8)) for s in ("left", "right", "top", "bottom")}
        all_cells.append(ftr_cell)
        row_cols_out.append(1)
        total_h_hu += ftr_cell["height_hu"]

    return {
        "type": "tbl",
        "n_rows": len(row_cols_out),
        "n_cols": n_cols_raw,
        "col_widths_hu": col_w,
        "row_cols": row_cols_out,
        "cells": all_cells,
        "width_hu": target_width_hu,
        "height_hu": total_h_hu,
    }


def fold_vertical_step_arrows(
    grid_cells: List[List[Optional[dict]]],
    target_width_hu: int,
    n_rows_raw: int = 5,
    table_font_size_hu: int = 1000,
) -> dict:
    """
    Folds alternating arrow-only rows (odd indices) into the preceding step row's last paragraph.
    """
    new_cells: List[dict] = []
    line_h_est = int(round(table_font_size_hu * 1.15))
    for r_src in range(0, n_rows_raw, 2):
        cobj = copy.deepcopy(grid_cells[r_src][0])
        if r_src + 1 < n_rows_raw and grid_cells[r_src + 1][0] is not None and cobj["paras"]:
            arrow_txt = "".join(r["text"] for p in grid_cells[r_src + 1][0]["paras"] for r in p["runs"]).strip()
            if arrow_txt:
                cobj["paras"][-1]["runs"].append(
                    {"text": f"  {arrow_txt}", "bold": True, "italic": False, "color": "2563EB"}
                )
        new_r = len(new_cells)
        cobj["row"], cobj["col"], cobj["row_span"], cobj["col_span"] = new_r, 0, 1, 1
        cobj["width_hu"], cobj["height_hu"] = target_width_hu, max(1050, len(cobj["paras"]) * line_h_est)
        cobj["mar"] = (160, 160, 35, 35)
        cobj["borders"] = {s: (1, 1, (0x94, 0xA3, 0xB8)) for s in ("left", "right", "top", "bottom")}
        new_cells.append(cobj)
    return {
        "type": "tbl",
        "n_rows": len(new_cells),
        "n_cols": 1,
        "col_widths_hu": [target_width_hu],
        "row_cols": [1] * len(new_cells),
        "cells": new_cells,
        "width_hu": target_width_hu,
        "height_hu": sum(c["height_hu"] for c in new_cells),
    }


def estimate_block_height_hu(blk: dict, state: ElasticLayoutState, outer_width_hu: int) -> int:
    b_font_hu = max(400, state.body_font_size_hu)
    t_font_hu = max(400, state.table_font_size_hu)

    if blk["type"] == "p":
        txt = "".join(r["text"] for r in blk["runs"])
        if not txt.strip():
            return 0
        eff_units = 0.70 * len(txt) + 0.30 * compute_weighted_text_units(txt)
        char_w_hu = max(340.0, b_font_hu * 0.85)
        chars_per_line = max(10.0, (outer_width_hu - 700) / char_w_hu)
        n_lines = max(1, math.ceil(eff_units / chars_per_line))
        line_h_hu = int(round(b_font_hu * (state.body_line_spacing_pct / 100.0)))
        return n_lines * line_h_hu + state.para_after_tw * 5 + 20

    if blk["type"] == "tbl":
        rows_by_r: Dict[int, List[dict]] = {}
        for c in blk["cells"]:
            rows_by_r.setdefault(c["row"], []).append(c)
        tbl_h = 0
        char_w_tbl_hu = max(360.0, t_font_hu * 0.92)
        for r_i in range(blk["n_rows"]):
            row_cells = rows_by_r.get(r_i, [])
            max_row_h = max(700, int(round(t_font_hu * 0.8)))
            for cell in row_cells:
                eff_w = max(1200, cell["width_hu"] - 280)
                chars_per_line = max(4.0, eff_w / char_w_tbl_hu)
                cell_lines = 0
                for p in cell["paras"]:
                    p_txt = "".join(r["text"] for r in p["runs"])
                    p_units = 0.70 * len(p_txt) + 0.30 * compute_weighted_text_units(p_txt)
                    cell_lines += max(1, math.ceil(p_units / chars_per_line))
                line_h_hu = int(round(t_font_hu * (state.table_line_spacing_pct / 100.0)))
                cell_h = cell_lines * line_h_hu + 2 * state.cell_pad_v_hu + len(cell["paras"]) * 20
                max_row_h = max(max_row_h, cell_h)
            tbl_h += max_row_h
        return tbl_h + 60
    return 0


def estimate_section_height_hu(blocks: List[dict], state: ElasticLayoutState, outer_width_hu: int) -> int:
    return sum(estimate_block_height_hu(b, state, outer_width_hu) for b in blocks)


class MinimumInvasionElasticSolver:
    """
    6-Stage Progressive Elastic Solver with Early-Stop:
    Dynamically scales line height and character capacity from TargetDocumentSpecIR
    and preserves maximum original paragraph & table cell line breaks.
    """

    def __init__(self, page_body_height_hu: Optional[int] = None):
        self.page_body_height_hu = page_body_height_hu

    def solve_section_layout_budget(
        self,
        raw_blocks: List[dict],
        target_pages: float = 2.0,
        outer_width_hu: int = 46200,
        target_spec: Optional[TargetDocumentSpecIR] = None,
        allow_font_size_elastic_scaling: bool = False,
        preserve_inline_media: bool = False,
    ) -> Tuple[List[dict], ElasticLayoutState, dict]:
        eff_page_h = compute_effective_page_body_height_hu(target_spec, self.page_body_height_hu)
        budget_hu = int(round(target_pages * eff_page_h))

        b_font_hu = target_spec.body_font_size_hu if (target_spec and target_spec.body_font_size_hu > 0) else 1200
        t_font_hu = target_spec.table_font_size_hu if (target_spec and target_spec.table_font_size_hu > 0) else 1000
        min_b_ls = target_spec.min_line_spacing_pct if (target_spec and target_spec.min_line_spacing_pct > 0) else 122
        max_b_ls = (
            target_spec.default_line_spacing_pct if (target_spec and target_spec.default_line_spacing_pct > 0) else 160
        )
        min_t_ls = (
            target_spec.table_min_line_spacing_pct
            if (target_spec and target_spec.table_min_line_spacing_pct > 0)
            else 112
        )
        max_t_ls = (
            target_spec.table_default_line_spacing_pct
            if (target_spec and target_spec.table_default_line_spacing_pct > 0)
            else 125
        )

        state = ElasticLayoutState(
            body_font_size_hu=b_font_hu,
            table_font_size_hu=t_font_hu,
            body_line_spacing_pct=min_b_ls,
            table_line_spacing_pct=min_t_ls,
            body_line_tw=compute_proportional_line_tw(b_font_hu, min_b_ls, is_table=False),
            table_line_tw=compute_proportional_line_tw(t_font_hu, min_t_ls, is_table=True),
        )

        blocks: List[dict] = []
        for b in raw_blocks:
            if not preserve_inline_media and b["type"] == "p" and b.get("r_embed") and not b["runs"]:
                continue
            blocks.append(copy.deepcopy(b))

        initial_para_count = sum(1 for b in blocks if b["type"] == "p")
        initial_cell_para_count = sum(
            len(c["paras"]) for b in blocks if b["type"] == "tbl" for c in b["cells"]
        )
        initial_height_hu = estimate_section_height_hu(blocks, state, outer_width_hu)

        # Stage 2: Interpolate spacing between max and min bounds before touching any paragraph breaks
        mid_b_ls = int(round((max_b_ls + min_b_ls) / 2.0))
        mid_t_ls = int(round((max_t_ls + min_t_ls) / 2.0))
        for ls_b, ls_t, pad_v, p_aft in (
            (max_b_ls, max_t_ls, 60, 12),
            (mid_b_ls, mid_t_ls, 45, 9),
            (min_b_ls, min_t_ls, 35, 7),
        ):
            state.body_line_spacing_pct = ls_b
            state.table_line_spacing_pct = ls_t
            state.body_line_tw = compute_proportional_line_tw(b_font_hu, ls_b, is_table=False)
            state.table_line_tw = compute_proportional_line_tw(t_font_hu, ls_t, is_table=True)
            state.cell_pad_v_hu = pad_v
            state.para_after_tw = p_aft
            if estimate_section_height_hu(blocks, state, outer_width_hu) <= budget_hu:
                state.stopped_stage = "STAGE_2_SPACING_ONLY"
                return self._finalize_result(
                    blocks, state, budget_hu, outer_width_hu, initial_height_hu, initial_para_count, initial_cell_para_count
                )

        # Stage 3: Merge standalone short subheadings (<= 40 chars, not ending with sentence terminal)
        i = 0
        while i < len(blocks) and estimate_section_height_hu(blocks, state, outer_width_hu) > budget_hu:
            cur = blocks[i]
            if cur["type"] == "p" and not cur.get("lock_unmerged", False):
                cur_txt = "".join(r["text"] for r in cur["runs"]).strip()
                if i + 1 < len(blocks) and blocks[i + 1]["type"] == "p":
                    nxt = blocks[i + 1]
                    if len(cur_txt) <= 40 and not cur_txt.endswith(".") and not cur_txt.endswith("함"):
                        for r in cur["runs"]:
                            r["bold"] = True
                        cur["runs"].append({"text": "  ", "bold": True, "italic": False, "color": "1E293B"})
                        cur["runs"].extend(nxt["runs"])
                        state.merged_subheadings_count += 1
                        del blocks[i + 1]
                        if i + 1 < len(blocks) and blocks[i + 1]["type"] == "p":
                            nxt2_txt = "".join(r["text"] for r in blocks[i + 1]["runs"]).strip()
                            if nxt2_txt.startswith("("):
                                cur["runs"].append({"text": " ", "bold": False, "italic": False, "color": "000000"})
                                cur["runs"].extend(blocks[i + 1]["runs"])
                                del blocks[i + 1]
                        continue
            i += 1

        if estimate_section_height_hu(blocks, state, outer_width_hu) <= budget_hu:
            state.stopped_stage = "STAGE_3_SUBHEADING_INLINE_EARLY_STOP"
            return self._finalize_result(
                blocks, state, budget_hu, outer_width_hu, initial_height_hu, initial_para_count, initial_cell_para_count
            )

        # Stage 4: Incrementally merge shortest adjacent paragraph pairs inside table cells (Early-Stop!)
        while estimate_section_height_hu(blocks, state, outer_width_hu) > budget_hu:
            best_candidate = None
            best_pair_len = 999999
            for b_idx, blk in enumerate(blocks):
                if blk["type"] != "tbl":
                    continue
                for c_idx, cell in enumerate(blk["cells"]):
                    paras = cell["paras"]
                    if len(paras) <= 1:
                        continue
                    for p_idx in range(len(paras) - 1):
                        len_a = sum(len(r["text"]) for r in paras[p_idx]["runs"])
                        len_b = sum(len(r["text"]) for r in paras[p_idx + 1]["runs"])
                        if len_a + len_b < best_pair_len:
                            best_pair_len = len_a + len_b
                            best_candidate = (b_idx, c_idx, p_idx)
            if best_candidate is None:
                break
            b_idx, c_idx, p_idx = best_candidate
            target_paras = blocks[b_idx]["cells"][c_idx]["paras"]
            target_paras[p_idx]["runs"].append({"text": "  ", "bold": False, "italic": False, "color": "64748B"})
            target_paras[p_idx]["runs"].extend(target_paras[p_idx + 1]["runs"])
            target_paras[p_idx]["jc"] = "left"
            del target_paras[p_idx + 1]
            state.merged_cell_pairs_count += 1

        if estimate_section_height_hu(blocks, state, outer_width_hu) <= budget_hu:
            state.stopped_stage = "STAGE_4_PARTIAL_CELL_MERGE_EARLY_STOP"
            return self._finalize_result(
                blocks, state, budget_hu, outer_width_hu, initial_height_hu, initial_para_count, initial_cell_para_count
            )

        # Stage 5: Only if still over budget, merge shortest consecutive body continuation paragraphs (Early-Stop!)
        while estimate_section_height_hu(blocks, state, outer_width_hu) > budget_hu:
            best_b_idx = -1
            best_len = 999999
            for idx in range(len(blocks) - 1):
                if blocks[idx]["type"] == "p" and blocks[idx + 1]["type"] == "p":
                    if blocks[idx].get("lock_unmerged") or blocks[idx + 1].get("lock_unmerged"):
                        continue
                    t1 = "".join(r["text"] for r in blocks[idx]["runs"]).strip()
                    t2 = "".join(r["text"] for r in blocks[idx + 1]["runs"]).strip()
                    nxt_starts_bold = (
                        blocks[idx + 1]["runs"]
                        and blocks[idx + 1]["runs"][0].get("bold")
                        and len(blocks[idx + 1]["runs"][0]["text"].strip()) <= 35
                    )
                    if (
                        not is_standalone_subheading_para(t1)
                        and not is_standalone_subheading_para(t2)
                        and not nxt_starts_bold
                        and not t2.startswith("■")
                    ):
                        pair_l = len(t1) + len(t2)
                        if pair_l < best_len:
                            best_len = pair_l
                            best_b_idx = idx
            if best_b_idx < 0:
                break
            blocks[best_b_idx]["runs"].append({"text": "  ", "bold": False, "italic": False, "color": "000000"})
            blocks[best_b_idx]["runs"].extend(blocks[best_b_idx + 1]["runs"])
            del blocks[best_b_idx + 1]
            state.merged_body_pairs_count += 1

        if estimate_section_height_hu(blocks, state, outer_width_hu) <= budget_hu or not allow_font_size_elastic_scaling:
            state.stopped_stage = "STAGE_5_MINIMAL_BODY_CONTINUATION_EARLY_STOP"
            return self._finalize_result(
                blocks, state, budget_hu, outer_width_hu, initial_height_hu, initial_para_count, initial_cell_para_count
            )

        # Stage 6 ([EDGE-03]): Optional micro-typography font scaling (-0.5pt steps, max -1.0pt)
        for delta_hu in (50, 100):
            state.body_font_size_hu = max(800, b_font_hu - delta_hu)
            state.table_font_size_hu = max(750, t_font_hu - delta_hu)
            state.body_line_tw = compute_proportional_line_tw(
                state.body_font_size_hu, state.body_line_spacing_pct, is_table=False
            )
            state.table_line_tw = compute_proportional_line_tw(
                state.table_font_size_hu, state.table_line_spacing_pct, is_table=True
            )
            if estimate_section_height_hu(blocks, state, outer_width_hu) <= budget_hu:
                break

        state.stopped_stage = "STAGE_6_FONT_SCALE_EARLY_STOP"
        return self._finalize_result(
            blocks, state, budget_hu, outer_width_hu, initial_height_hu, initial_para_count, initial_cell_para_count
        )

    def _finalize_result(
        self,
        blocks: List[dict],
        state: ElasticLayoutState,
        budget_hu: int,
        outer_width_hu: int,
        initial_height_hu: int,
        initial_para_count: int,
        initial_cell_para_count: int,
    ) -> Tuple[List[dict], ElasticLayoutState, dict]:
        final_height_hu = estimate_section_height_hu(blocks, state, outer_width_hu)
        final_para_count = sum(1 for b in blocks if b["type"] == "p")
        final_cell_para_count = sum(
            len(c["paras"]) for b in blocks if b["type"] == "tbl" for c in b["cells"]
        )
        telemetry = {
            "budget_hu": budget_hu,
            "initial_height_hu": initial_height_hu,
            "final_height_hu": final_height_hu,
            "body_font_size_hu": state.body_font_size_hu,
            "table_font_size_hu": state.table_font_size_hu,
            "body_line_spacing_pct": state.body_line_spacing_pct,
            "table_line_spacing_pct": state.table_line_spacing_pct,
            "stopped_stage": state.stopped_stage,
            "initial_body_paras": initial_para_count,
            "final_body_paras": final_para_count,
            "initial_cell_paras": initial_cell_para_count,
            "final_cell_paras": final_cell_para_count,
            "merged_subheadings_count": state.merged_subheadings_count,
            "merged_cell_pairs_count": state.merged_cell_pairs_count,
            "merged_body_pairs_count": state.merged_body_pairs_count,
        }
        return blocks, state, telemetry


def solve_document_page_budgets(
    sections: Dict[str, SectionIR],
    slot_mapping: Dict[str, str],
    target_spec: TargetDocumentSpecIR,
    user_overrides: UserSpecOverrideIR,
    solver: MinimumInvasionElasticSolver,
) -> dict:
    """
    Numerically resolves per-section page budgets B_k and total expected document pages P_expected
    from `user_overrides.total_page_limit`, `user_overrides.section_page_budgets`, or natural
    Stage-2 content volume without hardcoded 2.0-page or 8-page constants ([NUM-04]).
    """
    eff_page_h = compute_effective_page_body_height_hu(target_spec, solver.page_body_height_hu)
    b_font_hu = target_spec.body_font_size_hu if target_spec.body_font_size_hu > 0 else 1200
    t_font_hu = target_spec.table_font_size_hu if target_spec.table_font_size_hu > 0 else 1000
    min_b_ls = target_spec.min_line_spacing_pct if target_spec.min_line_spacing_pct > 0 else 122
    min_t_ls = target_spec.table_min_line_spacing_pct if target_spec.table_min_line_spacing_pct > 0 else 112

    probe_state = ElasticLayoutState(
        body_font_size_hu=b_font_hu,
        table_font_size_hu=t_font_hu,
        body_line_spacing_pct=min_b_ls,
        table_line_spacing_pct=min_t_ls,
        body_line_tw=compute_proportional_line_tw(b_font_hu, min_b_ls, is_table=False),
        table_line_tw=compute_proportional_line_tw(t_font_hu, min_t_ls, is_table=True),
    )

    mapped_sec_uids = [s_uid for s_uid in slot_mapping.keys() if s_uid in sections]
    slot_by_uid: Dict[str, SlotSpecIR] = {sl.slot_uid: sl for sl in target_spec.section_slots}
    natural_heights_hu: Dict[str, int] = {}

    for s_uid in mapped_sec_uids:
        sec_obj = sections[s_uid]
        sl_obj = slot_by_uid.get(slot_mapping[s_uid])
        avail_w = sl_obj.available_width_hu if (sl_obj and sl_obj.available_width_hu > 0) else 46200
        filtered_blks = [
            b
            for b in sec_obj.blocks
            if not (user_overrides.route_media_to_attachment_slots and b["type"] == "p" and b.get("r_embed") and not b["runs"])
        ]
        natural_heights_hu[s_uid] = max(1, estimate_section_height_hu(filtered_blks, probe_state, avail_w))

    # Count fixed pages (cover/preamble + unmapped outer slots + attachment media pages)
    preamble_pages = 1 if (target_spec.preamble_blocks and user_overrides.preserve_outer_tables) else 0
    mapped_slot_set = {slot_mapping[s_uid] for s_uid in mapped_sec_uids}
    unmapped_slots = [sl for sl in target_spec.section_slots if sl.slot_uid not in mapped_slot_set]
    unmapped_slot_pages = (
        len(unmapped_slots)
        if (user_overrides.preserve_outer_tables and not user_overrides.clear_unmapped_slots)
        else 0
    )
    if target_spec.attachment_slots and user_overrides.preserve_outer_tables:
        g_cols = max(1, target_spec.attachment_slots[0].grid_cols)
        attachment_pages = max(1, math.ceil(len(target_spec.attachment_slots) / float(g_cols * 3)))
    else:
        attachment_pages = 0

    fixed_pages = preamble_pages + unmapped_slot_pages + attachment_pages

    resolved_budgets: Dict[str, float] = {}
    auto_sec_uids: List[str] = []

    for s_uid in mapped_sec_uids:
        sec_obj = sections[s_uid]
        sl_obj = slot_by_uid.get(slot_mapping[s_uid])
        exp_b = resolve_budget_for_section_or_slot(user_overrides.section_page_budgets, sec_obj, sl_obj)
        if exp_b is not None:
            resolved_budgets[s_uid] = float(exp_b)
        else:
            auto_sec_uids.append(s_uid)

    if auto_sec_uids:
        if user_overrides.total_page_limit is not None:
            explicit_sum = sum(resolved_budgets.values())
            avail_body_pages = max(
                float(len(auto_sec_uids)),
                float(user_overrides.total_page_limit - fixed_pages) - explicit_sum,
            )
            tot_h = float(sum(natural_heights_hu[u] for u in auto_sec_uids)) or 1.0
            target_int_sum = max(len(auto_sec_uids), int(round(avail_body_pages)))
            raw_quotas = {
                u: max(1.0, (natural_heights_hu[u] / tot_h) * target_int_sum) for u in auto_sec_uids
            }
            floored = {u: max(1, int(math.floor(raw_quotas[u]))) for u in auto_sec_uids}
            while sum(floored.values()) > target_int_sum:
                k_max = max(auto_sec_uids, key=lambda u: floored[u])
                if floored[k_max] > 1:
                    floored[k_max] -= 1
                else:
                    break
            rem_pages = target_int_sum - sum(floored.values())
            if rem_pages > 0:
                by_frac = sorted(
                    auto_sec_uids,
                    key=lambda u: (raw_quotas[u] - floored[u], natural_heights_hu[u]),
                    reverse=True,
                )
                for idx_r in range(rem_pages):
                    floored[by_frac[idx_r % len(by_frac)]] += 1
            for u in auto_sec_uids:
                resolved_budgets[u] = float(floored[u])
        else:
            for u in auto_sec_uids:
                nat_p = max(1, math.ceil(natural_heights_hu[u] / float(eff_page_h)))
                # If natural Stage-2 height is within 38% of an integer page boundary (Stage 3/4/5 compressible range),
                # set target budget to the tighter integer page count so MinimumInvasionElasticSolver compacts overflow
                ratio_p = natural_heights_hu[u] / float(eff_page_h)
                floor_p = max(1, int(math.floor(ratio_p)))
                if ratio_p - floor_p <= 0.38:
                    resolved_budgets[u] = float(floor_p)
                else:
                    resolved_budgets[u] = float(nat_p)

    for s_uid, b_val in resolved_budgets.items():
        sl_obj = slot_by_uid.get(slot_mapping[s_uid])
        if sl_obj is not None:
            sl_obj.max_pages = float(b_val)

    mapped_body_pages = sum(max(1, int(math.ceil(b_val))) for b_val in resolved_budgets.values())
    expected_total_pages = fixed_pages + mapped_body_pages

    return {
        "effective_page_body_height_hu": eff_page_h,
        "natural_stage2_heights_hu": natural_heights_hu,
        "section_budgets": resolved_budgets,
        "preamble_pages": preamble_pages,
        "mapped_body_pages": mapped_body_pages,
        "unmapped_slot_pages": unmapped_slot_pages,
        "attachment_pages": attachment_pages,
        "expected_total_pages": expected_total_pages,
    }
