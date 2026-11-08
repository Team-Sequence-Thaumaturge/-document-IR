"""
T-DIR v3.0 2D Topological Grid Normalizer & 5-Stage Minimum-Invasion Elastic Layout Solver.
Eliminates empty spacer rows/columns while preserving original paragraph & cell line breaks
up to the exact Early-Stop threshold required by the target page budget.
"""

import math
import copy
from dataclasses import dataclass
from typing import List, Dict, Tuple, Set


@dataclass
class ElasticLayoutState:
    snap_to_grid: bool = False
    prune_empty_spacers: bool = True
    body_line_spacing_pct: int = 122
    table_line_spacing_pct: int = 112
    body_line_tw: int = 268
    table_line_tw: int = 212
    cell_pad_v_hu: int = 35
    para_after_tw: int = 7
    merged_subheadings_count: int = 0
    merged_cell_pairs_count: int = 0
    merged_body_pairs_count: int = 0
    stopped_stage: str = "STAGE_1_SPACERS_ONLY"


def detect_and_remove_spacer_rows_cols(
    grid_cells: List[List[dict]],
    col_twips: List[int],
    n_rows_raw: int,
    n_cols_raw: int,
    target_width_hu: int,
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

    col_widths_hu = normalize_column_ratios_with_arrow_clamp(
        grid_cells, active_rows, active_cols, col_twips, target_width_hu
    )

    new_cells: List[dict] = []
    seen_ids: Set[int] = set()
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
                    "height_hu": max(800, len(cell_paras) * 950 + 80),
                    "mar": (140, 140, 35, 35),
                    "borders": {s: (1, 1, (0x94, 0xA3, 0xB8)) for s in ("left", "right", "top", "bottom")},
                    "fill_hex": cobj["fill_hex"],
                    "paras": cell_paras,
                }
            )

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
        "height_hu": sum(
            max((c["height_hu"] for c in new_cells if c["row"] == r), default=800) for r in range(n_rows_new)
        ),
    }


def normalize_column_ratios_with_arrow_clamp(
    grid_cells: List[List[dict]],
    active_rows: List[int],
    active_cols: List[int],
    col_twips: List[int],
    target_width_hu: int,
) -> List[int]:
    n_cols_new = len(active_cols)
    arrow_col_indices: Set[int] = set()
    for new_c, old_c in enumerate(active_cols):
        col_txts = [
            "".join(r["text"] for p in grid_cells[old_r][old_c]["paras"] for r in p["runs"]).strip()
            for old_r in active_rows
            if grid_cells[old_r][old_c] is not None
            and grid_cells[old_r][old_c]["col"] == old_c
            and grid_cells[old_r][old_c]["col_span"] == 1
        ]
        if col_txts and all(t in ("▶", "→", "↓", "") for t in col_txts):
            arrow_col_indices.add(new_c)

    if arrow_col_indices and len(arrow_col_indices) < n_cols_new:
        arrow_w_hu = 1200
        rem_w_hu = max(1000, target_width_hu - arrow_w_hu * len(arrow_col_indices))
        non_arrow_cols = [old_c for new_c, old_c in enumerate(active_cols) if new_c not in arrow_col_indices]
        tot_tw = max(1, sum(col_twips[c] if c < len(col_twips) else 1000 for c in non_arrow_cols))
        col_widths_hu: List[int] = []
        for new_c, old_c in enumerate(active_cols):
            if new_c in arrow_col_indices:
                col_widths_hu.append(arrow_w_hu)
            else:
                w_raw = col_twips[old_c] if old_c < len(col_twips) else 1000
                col_widths_hu.append(int(round(w_raw * rem_w_hu / float(tot_tw))))
        if col_widths_hu:
            col_widths_hu[-1] += target_width_hu - sum(col_widths_hu)
        return col_widths_hu

    raw_active_col_w = [col_twips[c] if c < len(col_twips) else 1000 for c in active_cols]
    tot_tw = max(1, sum(raw_active_col_w))
    col_widths_hu = [int(round(w * target_width_hu / float(tot_tw))) for w in raw_active_col_w]
    if col_widths_hu:
        col_widths_hu[-1] += target_width_hu - sum(col_widths_hu)
    return col_widths_hu


def collapse_stair_step_chart_table(grid_cells: List[List[dict]], target_width_hu: int) -> dict:
    col_texts: List[List[dict]] = [[] for _ in range(4)]
    col_fills: List[str] = ["F8FAFC"] * 4
    for r_i in range(1, 5):
        for c_i in range(4):
            cobj = grid_cells[r_i][c_i]
            if cobj and cobj["paras"]:
                for p in cobj["paras"]:
                    if p not in col_texts[c_i]:
                        col_texts[c_i].append(copy.deepcopy(p))
                if cobj["fill_hex"]:
                    col_fills[c_i] = cobj["fill_hex"]

    new_cells: List[dict] = []
    r0_c = copy.deepcopy(grid_cells[0][0])
    r0_c["row"], r0_c["col"], r0_c["row_span"], r0_c["col_span"] = 0, 0, 1, 4
    new_cells.append(r0_c)

    for c_i in range(4):
        new_cells.append(
            {
                "row": 1,
                "col": c_i,
                "row_span": 1,
                "col_span": 1,
                "fill_hex": col_fills[c_i],
                "paras": col_texts[c_i],
            }
        )
    for c_i in range(4):
        rc = copy.deepcopy(grid_cells[5][c_i])
        rc["row"], rc["col"], rc["row_span"], rc["col_span"] = 2, c_i, 1, 1
        new_cells.append(rc)

    r3_c = copy.deepcopy(grid_cells[6][0])
    r3_c["row"], r3_c["col"], r3_c["row_span"], r3_c["col_span"] = 3, 0, 1, 4
    new_cells.append(r3_c)

    w_col = target_width_hu // 4
    col_w = [w_col, w_col, w_col, target_width_hu - 3 * w_col]
    for c in new_cells:
        c["width_hu"] = sum(col_w[c["col"] : c["col"] + c["col_span"]])
        c["mar"] = (140, 140, 35, 35)
        c["borders"] = {s: (1, 1, (0x94, 0xA3, 0xB8)) for s in ("left", "right", "top", "bottom")}
        if not c["paras"]:
            c["paras"] = [
                {
                    "type": "p",
                    "jc": "center",
                    "left_tw": 0,
                    "right_tw": 0,
                    "indent_tw": 0,
                    "before_tw": 0,
                    "after_tw": 0,
                    "runs": [],
                    "r_embed": None,
                }
            ]
        c["height_hu"] = max(800, len(c["paras"]) * 950 + 80)

    return {
        "type": "tbl",
        "n_rows": 4,
        "n_cols": 4,
        "col_widths_hu": col_w,
        "row_cols": [1, 4, 4, 1],
        "cells": new_cells,
        "width_hu": target_width_hu,
        "height_hu": sum(max(c["height_hu"] for c in new_cells if c["row"] == r) for r in range(4)),
    }


def fold_vertical_step_arrows(grid_cells: List[List[dict]], target_width_hu: int) -> dict:
    r0_c = copy.deepcopy(grid_cells[0][0])
    r1_c = grid_cells[1][0]
    r2_c = copy.deepcopy(grid_cells[2][0])
    r3_c = grid_cells[3][0]
    r4_c = copy.deepcopy(grid_cells[4][0])

    if r1_c and r1_c["paras"] and r0_c["paras"]:
        r0_c["paras"][-1]["runs"].append({"text": "  ", "bold": False, "italic": False, "color": "000000"})
        r0_c["paras"][-1]["runs"].extend(copy.deepcopy(r1_c["paras"][0]["runs"]))
    if r3_c and r3_c["paras"] and r2_c["paras"]:
        r2_c["paras"][-1]["runs"].append({"text": "  ", "bold": False, "italic": False, "color": "000000"})
        r2_c["paras"][-1]["runs"].extend(copy.deepcopy(r3_c["paras"][0]["runs"]))

    new_cells: List[dict] = []
    for idx_r, cobj in enumerate((r0_c, r2_c, r4_c)):
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


def estimate_block_height_hu(blk: dict, state: ElasticLayoutState, outer_width_hu: int) -> int:
    if blk["type"] == "p":
        char_len = sum(len(r["text"]) for r in blk["runs"])
        chars_per_line = max(10.0, (outer_width_hu / 1200.0) * 1.02)
        n_lines = max(1, math.ceil(char_len / chars_per_line))
        line_h_hu = int(1200 * (state.body_line_spacing_pct / 100.0))
        return n_lines * line_h_hu + state.para_after_tw * 5
    elif blk["type"] == "tbl":
        tbl_h = 0
        for r_idx in range(blk["n_rows"]):
            row_cells = [c for c in blk["cells"] if c["row"] == r_idx]
            max_row_h = 0
            for cell in row_cells:
                eff_w = max(1200, cell["width_hu"] - 280)
                chars_per_line = max(4.0, eff_w / 920.0)
                cell_lines = 0
                for p in cell["paras"]:
                    p_len = sum(len(r["text"]) for r in p["runs"])
                    cell_lines += max(1, math.ceil(p_len / chars_per_line))
                line_h_hu = int(1000 * (state.table_line_spacing_pct / 100.0))
                cell_h = cell_lines * line_h_hu + 2 * state.cell_pad_v_hu + len(cell["paras"]) * 20
                max_row_h = max(max_row_h, cell_h)
            tbl_h += max_row_h
        return tbl_h + 60
    return 0


def estimate_section_height_hu(blocks: List[dict], state: ElasticLayoutState, outer_width_hu: int) -> int:
    return sum(estimate_block_height_hu(b, state, outer_width_hu) for b in blocks)


class MinimumInvasionElasticSolver:
    """
    5-Stage Progressive Elastic Solver with Early-Stop:
    Preserves as many original paragraph breaks and table cell line breaks as possible
    while strictly fitting within target_pages.
    """

    def __init__(self, page_body_height_hu: int = 78500):
        self.page_body_height_hu = page_body_height_hu

    def solve_section_layout_budget(
        self,
        raw_blocks: List[dict],
        target_pages: float = 2.0,
        outer_width_hu: int = 46200,
    ) -> Tuple[List[dict], ElasticLayoutState, dict]:
        budget_hu = int(target_pages * self.page_body_height_hu)
        state = ElasticLayoutState()

        # Filter out detached inline image-only paragraphs
        blocks: List[dict] = []
        for b in raw_blocks:
            if b["type"] == "p" and b.get("r_embed") and not b["runs"]:
                continue
            blocks.append(copy.deepcopy(b))

        initial_para_count = sum(1 for b in blocks if b["type"] == "p")
        initial_cell_para_count = sum(
            len(c["paras"]) for b in blocks if b["type"] == "tbl" for c in b["cells"]
        )
        initial_height_hu = estimate_section_height_hu(blocks, state, outer_width_hu)

        # Stage 2: Check if pure spacing adjustment (without any paragraph merging) fits budget_hu
        for ls_b, ls_t, bw_tw, tw_tw, pad_v, p_aft in (
            (142, 125, 290, 235, 60, 12),
            (132, 118, 278, 222, 45, 9),
            (122, 112, 268, 212, 35, 7),
        ):
            state.body_line_spacing_pct = ls_b
            state.table_line_spacing_pct = ls_t
            state.body_line_tw = bw_tw
            state.table_line_tw = tw_tw
            state.cell_pad_v_hu = pad_v
            state.para_after_tw = p_aft
            if estimate_section_height_hu(blocks, state, outer_width_hu) <= budget_hu:
                state.stopped_stage = "STAGE_2_SPACING_ONLY"
                return self._finalize_result(
                    blocks, state, budget_hu, outer_width_hu, initial_height_hu, initial_para_count, initial_cell_para_count
                )

        # Stage 3: Merge standalone short subheadings (<= 40 chars, not ending with '.' or '함') with following paragraph
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

        # Stage 4: Incrementally merge shortest adjacent paragraph pairs inside table cells until budget_hu is met (Early-Stop!)
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
        section_titles = {
            "가설",
            "핵심 타겟 및 목표 시장 규모",
            "문제의 크기·시급성 및 기대효과",
            "구성",
            "구현 정도",
            "검증 방안 및 고객 검증",
            "차별성 및 포지셔닝",
            "해결 과제 및 대응",
        }
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
                    if t1 not in section_titles and t2 not in section_titles and not nxt_starts_bold and not t2.startswith("■"):
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

        state.stopped_stage = "STAGE_5_MINIMAL_BODY_CONTINUATION_EARLY_STOP"
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
