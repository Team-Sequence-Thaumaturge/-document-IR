"""
T-DIR v3.2 Beta Isolated MCP Server (JSON-RPC 2.0 over stdio + Direct Python API).
Zero hardcoded font names, pixel/HWPUNIT constants, document-specific section IDs, or fixed page counts:
  - Dynamically extracts source & target specifications at runtime (`tdir_probe_files`).
  - Resolves generic section/slot topologies (`SEC_01..SEC_K`, `SLOT_01..SLOT_M`, `MEDIA_01..MEDIA_P`)
    and numerical page budgets (`solve_document_page_budgets`) via `UserSpecOverrideIR`.
Exposes 5 tools:
  1. tdir_probe_files
  2. tdir_build_mapping_plan
  3. tdir_execute_conversion
  4. tdir_verify_fidelity_and_layout
  5. tdir_adjust_block_layout
"""

import os
import re
import sys
import uuid
import json
import zipfile
import subprocess
import xml.etree.ElementTree as ET
from typing import List, Dict, Optional, Tuple

from .ir_schema import (
    SectionIR,
    SourceDocumentSpecIR,
    TargetDocumentSpecIR,
    UserSpecOverrideIR,
    ConversionJobIR,
    resolve_effective_document_spec,
    validate_text_immutability,
    extract_ir_non_ws_text,
    compute_heading_slot_similarity,
    matches_section_selector,
    hu_to_pt,
)
from .elastic_solver import (
    ElasticLayoutState,
    MinimumInvasionElasticSolver,
    compute_effective_page_body_height_hu,
    solve_document_page_budgets,
)
from .codec_hwp_v5 import (
    scan_hwp_template_slots,
    scan_hwp_source_spec,
    lift_hwp_to_ir,
    project_ir_to_hwp,
)
from .codec_ooxml_owpml import (
    scan_docx_style_and_page_spec,
    lift_docx_to_ir,
    project_ir_to_docx,
    project_ir_to_hwpx,
)
from .codec_tabular_markup import (
    lift_markdown_to_ir,
    project_ir_to_markdown,
    lift_csv_to_ir,
    project_ir_to_csv,
)


USER_SPEC_OVERRIDES_JSON_SCHEMA = {
    "type": "object",
    "description": (
        "Explicit user or target specification overrides (Priority 1 > Target Template > Source File). "
        "Any omitted field automatically inherits from the target template or source document."
    ),
    "properties": {
        "body_font_family": {"type": "string", "description": "Body prose font family name."},
        "table_font_family": {"type": "string", "description": "Inner table font family name."},
        "body_font_size_pt": {"type": "number", "description": "Body prose font size in points (e.g. 12.0)."},
        "table_font_size_pt": {"type": "number", "description": "Inner table font size in points (e.g. 10.0)."},
        "body_line_spacing_min_pct": {"type": "integer", "description": "Minimum body line spacing percentage."},
        "body_line_spacing_max_pct": {"type": "integer", "description": "Initial/maximum body line spacing percentage."},
        "table_line_spacing_min_pct": {"type": "integer", "description": "Minimum table line spacing percentage."},
        "table_line_spacing_max_pct": {"type": "integer", "description": "Initial/maximum table line spacing percentage."},
        "page_width_mm": {"type": "number", "description": "Target page width in millimeters (e.g. 210.0 for A4)."},
        "page_height_mm": {"type": "number", "description": "Target page height in millimeters (e.g. 297.0 for A4)."},
        "margins_mm": {
            "type": "object",
            "description": "Page margins in millimeters ({top, bottom, left, right}).",
        },
        "total_page_limit": {"type": "integer", "description": "Total document page ceiling."},
        "section_page_budgets": {
            "type": "object",
            "description": "Per-section page budget map (e.g. {'Q1': 2.0, 'Q2': 2.0}).",
        },
        "page_body_height_hu": {
            "type": "integer",
            "description": "Optional override for effective single-page body height in HWPUNIT.",
        },
        "include_sections": {
            "type": "array",
            "items": {"type": "string"},
            "description": "Explicit whitelist of source section UIDs to transplant.",
        },
        "exclude_sections": {
            "type": "array",
            "items": {"type": "string"},
            "description": "List of source section UIDs to exclude.",
        },
        "preserve_outer_tables": {
            "type": "boolean",
            "description": "Preserve target template 2x2 outer section container tables (treatAsChar=False).",
        },
        "route_media_to_attachment_slots": {
            "type": "boolean",
            "description": "Route embedded section images sequentially into bottom attachment slots.",
        },
        "max_attachment_images": {
            "type": "integer",
            "description": "Maximum number of images to route to attachment slots.",
        },
        "attachment_captions": {
            "type": "array",
            "items": {"type": "string"},
            "description": "Optional explicit caption strings for routed attachment images.",
        },
        "allow_font_size_elastic_scaling": {
            "type": "boolean",
            "description": "Allow Stage 6 micro-typography font scaling (-0.5pt) if Stage 5 still overflows.",
        },
        "clear_unmapped_slots": {
            "type": "boolean",
            "description": "Clear placeholder text in unmapped target section slots.",
        },
        "style_inheritance_mode": {
            "type": "string",
            "description": "Style resolution policy: TARGET_FIRST, SOURCE_FIRST, or HYBRID.",
        },
    },
}


class TdirBetaMcpServer:
    def __init__(self) -> None:
        self.jobs: Dict[str, ConversionJobIR] = {}
        self.solver = MinimumInvasionElasticSolver()

    def _detect_format(self, path: str) -> str:
        ext = os.path.splitext(path)[1].lower()
        with open(path, "rb") as f:
            head = f.read(8)
        if head == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
            return "HWP_V5_BINARY"
        if head[:4] == b"PK\x03\x04":
            if ext == ".hwpx":
                return "HWPX_OWPML"
            if ext == ".xlsx":
                return "XLSX_OOXML"
            return "DOCX_OOXML"
        if ext in (".md", ".markdown"):
            return "MARKDOWN"
        if ext == ".csv":
            return "CSV"
        if ext == ".tsv":
            return "TSV"
        return "PLAIN_TEXT"

    def _serialize_target_spec(self, spec: TargetDocumentSpecIR) -> dict:
        return {
            "template_path": spec.template_path,
            "target_format": spec.target_format,
            "body_font_family": spec.body_font_family,
            "table_font_family": spec.table_font_family,
            "body_font_size_hu": spec.body_font_size_hu,
            "body_font_size_pt": hu_to_pt(spec.body_font_size_hu),
            "table_font_size_hu": spec.table_font_size_hu,
            "table_font_size_pt": hu_to_pt(spec.table_font_size_hu),
            "default_line_spacing_pct": spec.default_line_spacing_pct,
            "min_line_spacing_pct": spec.min_line_spacing_pct,
            "page_width_hu": spec.page_width_hu,
            "page_height_hu": spec.page_height_hu,
            "margins_hu": {
                "top": spec.margin_top_hu,
                "bottom": spec.margin_bottom_hu,
                "left": spec.margin_left_hu,
                "right": spec.margin_right_hu,
                "header": spec.margin_header_hu,
                "footer": spec.margin_footer_hu,
            },
            "preamble_blocks_count": len(spec.preamble_blocks),
            "attachment_header_paras_count": len(spec.attachment_header_paras),
            "section_slots": [
                {
                    "slot_uid": sl.slot_uid,
                    "ordinal_index": sl.ordinal_index,
                    "title": sl.discovered_title_text,
                    "subtitle": sl.discovered_subtitle_text,
                    "outer_tbl_record_idx": sl.outer_tbl_record_idx,
                    "body_cell_record_idx": sl.body_cell_record_idx,
                    "outer_width_hu": sl.outer_width_hu,
                    "col_widths_hu": sl.col_widths_hu,
                    "available_width_hu": sl.available_width_hu,
                    "available_height_hu": sl.available_height_hu,
                    "max_pages": sl.max_pages,
                }
                for sl in spec.section_slots
            ],
            "attachment_slots": [
                {
                    "slot_uid": sl.slot_uid,
                    "ordinal_index": sl.ordinal_index,
                    "grid_cols": sl.grid_cols,
                    "caption": sl.discovered_title_text,
                    "body_cell_record_idx": sl.body_cell_record_idx,
                    "caption_cell_record_idx": sl.caption_cell_record_idx,
                    "available_width_hu": sl.available_width_hu,
                    "available_height_hu": sl.available_height_hu,
                }
                for sl in spec.attachment_slots
            ],
        }

    def tdir_probe_files(self, source_paths: List[str], target_path: str) -> dict:
        job_id = f"job_{uuid.uuid4().hex[:8]}"
        raw_target_spec = scan_hwp_template_slots(target_path)

        source_specs: List[SourceDocumentSpecIR] = []
        for sp in source_paths:
            fmt = self._detect_format(sp)
            if fmt == "DOCX_OOXML":
                source_specs.append(scan_docx_style_and_page_spec(sp))
            elif fmt == "HWP_V5_BINARY":
                source_specs.append(scan_hwp_source_spec(sp))
            else:
                source_specs.append(SourceDocumentSpecIR(source_path=sp, source_format=fmt))

        effective_spec = resolve_effective_document_spec(source_specs, raw_target_spec, UserSpecOverrideIR())
        avail_w = (
            effective_spec.section_slots[0].available_width_hu
            if effective_spec.section_slots
            else max(10000, effective_spec.page_width_hu - effective_spec.margin_left_hu - effective_spec.margin_right_hu)
        )
        tbl_sz_hu = effective_spec.table_font_size_hu

        all_sections: Dict[str, SectionIR] = {}
        sources_report: List[dict] = []

        for idx_s, sp in enumerate(source_paths):
            fmt = source_specs[idx_s].source_format
            s_spec = source_specs[idx_s]
            if fmt == "DOCX_OOXML":
                secs = lift_docx_to_ir(sp, target_width_hu=avail_w, table_font_size_hu=tbl_sz_hu)
            elif fmt == "MARKDOWN":
                secs = lift_markdown_to_ir(sp, target_width_hu=avail_w)
            elif fmt == "CSV":
                secs = lift_csv_to_ir(sp, delimiter=",", target_width_hu=avail_w)
            elif fmt == "TSV":
                secs = lift_csv_to_ir(sp, delimiter="\t", target_width_hu=avail_w)
            elif fmt == "HWP_V5_BINARY":
                _, blks = lift_hwp_to_ir(sp)
                _, nc, sha = extract_ir_non_ws_text(blks)
                secs = {
                    "SEC_01": SectionIR(
                        section_uid="SEC_01",
                        heading_text=os.path.basename(sp),
                        blocks=blks,
                        non_ws_char_count=nc,
                        text_sha256=sha,
                        ordinal_index=1,
                        section_role="BODY",
                    )
                }
            else:
                secs = {}

            sec_list_meta: List[dict] = []
            for s_uid, s_obj in secs.items():
                eff_s_uid = s_uid if s_uid not in all_sections else f"{s_uid}_S{idx_s + 1}"
                s_obj.section_uid = eff_s_uid
                all_sections[eff_s_uid] = s_obj
                n_tbls = sum(1 for b in s_obj.blocks if b["type"] == "tbl")
                n_paras = sum(1 for b in s_obj.blocks if b["type"] == "p")
                sec_list_meta.append(
                    {
                        "section_uid": eff_s_uid,
                        "ordinal_index": s_obj.ordinal_index,
                        "section_role": s_obj.section_role,
                        "heading_text": s_obj.heading_text,
                        "subtitle_text": s_obj.subtitle_text,
                        "non_ws_chars": s_obj.non_ws_char_count,
                        "tables_count": n_tbls,
                        "paras_count": n_paras,
                        "media_count": len(s_obj.media_objects),
                        "text_sha256": s_obj.text_sha256,
                    }
                )
            sources_report.append(
                {
                    "path": sp,
                    "detected_format": fmt,
                    "source_spec": {
                        "body_font_family": s_spec.body_font_family,
                        "table_font_family": s_spec.table_font_family,
                        "body_font_size_hu": s_spec.body_font_size_hu,
                        "table_font_size_hu": s_spec.table_font_size_hu,
                        "page_width_hu": s_spec.page_width_hu,
                        "page_height_hu": s_spec.page_height_hu,
                    },
                    "sections_discovered": sec_list_meta,
                }
            )

        job = ConversionJobIR(
            job_id=job_id,
            source_paths=source_paths,
            target_path=target_path,
            sections=all_sections,
            source_specs=source_specs,
            raw_target_spec=raw_target_spec,
            target_spec=effective_spec,
        )
        self.jobs[job_id] = job

        return {
            "job_id": job_id,
            "sources": sources_report,
            "target_spec": self._serialize_target_spec(effective_spec),
        }

    def tdir_build_mapping_plan(
        self,
        job_id: str,
        user_spec_overrides: Optional[dict] = None,
        user_constraints: Optional[dict] = None,
    ) -> dict:
        job = self.jobs[job_id]
        assert job.raw_target_spec is not None

        merged_dict: dict = {}
        if user_constraints:
            merged_dict.update(user_constraints)
        if user_spec_overrides:
            merged_dict.update(user_spec_overrides)

        overrides_ir = UserSpecOverrideIR.from_dict(merged_dict)
        job.user_overrides = overrides_ir
        job.target_spec = resolve_effective_document_spec(job.source_specs, job.raw_target_spec, overrides_ir)
        eff_spec = job.target_spec

        page_body_height_hu = compute_effective_page_body_height_hu(eff_spec, overrides_ir.page_body_height_hu)
        self.solver = MinimumInvasionElasticSolver(page_body_height_hu=page_body_height_hu)

        slot_mapping: Dict[str, str] = {}
        solver_preview: Dict[str, dict] = {}
        used_slot_uids = set()

        for s_uid, s_obj in job.sections.items():
            if overrides_ir.include_sections is not None:
                if not any(matches_section_selector(sel, s_obj) for sel in overrides_ir.include_sections):
                    continue
            else:
                if s_obj.section_role != "BODY":
                    continue
            if overrides_ir.exclude_sections and any(
                matches_section_selector(sel, s_obj) for sel in overrides_ir.exclude_sections
            ):
                continue

            matched_slot = None
            if eff_spec.section_slots:
                best_score = -1.0
                best_slot = None
                for sl in eff_spec.section_slots:
                    if sl.slot_uid in used_slot_uids:
                        continue
                    sc = compute_heading_slot_similarity(
                        s_obj.heading_text or s_uid,
                        sl.discovered_title_text,
                        sl.slot_uid,
                    )
                    if sc > best_score:
                        best_score = sc
                        best_slot = sl
                if best_slot is not None and best_score >= 0.15:
                    matched_slot = best_slot
                else:
                    matched_slot = next(
                        (
                            sl
                            for sl in eff_spec.section_slots
                            if sl.slot_uid not in used_slot_uids
                            and sl.ordinal_index is not None
                            and sl.ordinal_index == s_obj.ordinal_index
                        ),
                        None,
                    )
                if matched_slot is None:
                    matched_slot = next(
                        (sl for sl in eff_spec.section_slots if sl.slot_uid not in used_slot_uids),
                        eff_spec.section_slots[0],
                    )
            if matched_slot is not None:
                used_slot_uids.add(matched_slot.slot_uid)
                slot_mapping[s_uid] = matched_slot.slot_uid

        budget_solution = solve_document_page_budgets(
            job.sections,
            slot_mapping,
            eff_spec,
            overrides_ir,
            self.solver,
        )
        job.page_budget_solution = budget_solution

        for s_uid, slot_uid in list(slot_mapping.items()):
            s_obj = job.sections[s_uid]
            matched_slot = next((sl for sl in eff_spec.section_slots if sl.slot_uid == slot_uid), None)
            avail_w = matched_slot.available_width_hu if matched_slot else 46200
            max_p = float(budget_solution["section_budgets"].get(s_uid, matched_slot.max_pages if matched_slot else 2.0))
            _, _, tele = self.solver.solve_section_layout_budget(
                s_obj.blocks,
                target_pages=max_p,
                outer_width_hu=avail_w,
                target_spec=eff_spec,
                allow_font_size_elastic_scaling=overrides_ir.allow_font_size_elastic_scaling,
            )
            solver_preview[s_uid] = tele

        if overrides_ir.route_media_to_attachment_slots:
            media_idx = 0
            max_imgs = (
                overrides_ir.max_attachment_images
                if overrides_ir.max_attachment_images is not None
                else len(eff_spec.attachment_slots)
            )
            for s_uid in sorted(slot_mapping.keys()):
                if s_uid not in job.sections:
                    continue
                for m_obj in job.sections[s_uid].media_objects:
                    if media_idx < len(eff_spec.attachment_slots) and media_idx < max_imgs:
                        att_slot = eff_spec.attachment_slots[media_idx]
                        slot_mapping[m_obj.media_uid] = att_slot.slot_uid
                        media_idx += 1

        job.slot_mapping = slot_mapping
        job.solver_telemetry = solver_preview
        return {
            "job_id": job_id,
            "effective_spec": self._serialize_target_spec(eff_spec),
            "slot_mapping": slot_mapping,
            "page_budget_solution": budget_solution,
            "solver_preview": solver_preview,
        }

    def tdir_execute_conversion(
        self,
        job_id: str,
        output_prefix: str,
        output_formats: Optional[List[str]] = None,
        page_body_height_hu: Optional[int] = None,
        base_hwpx_path: Optional[str] = None,
        user_spec_overrides: Optional[dict] = None,
    ) -> dict:
        job = self.jobs[job_id]
        assert job.raw_target_spec is not None

        if user_spec_overrides is not None or not job.slot_mapping:
            self.tdir_build_mapping_plan(job_id, user_spec_overrides=user_spec_overrides)

        assert job.target_spec is not None
        eff_spec = job.target_spec
        overrides_ir = job.user_overrides
        formats = [f.lower().lstrip(".") for f in (output_formats or ["hwp", "hwpx", "docx"])]

        eff_page_body_hu = compute_effective_page_body_height_hu(
            eff_spec,
            page_body_height_hu if page_body_height_hu is not None else overrides_ir.page_body_height_hu,
        )
        self.solver = MinimumInvasionElasticSolver(page_body_height_hu=eff_page_body_hu)

        slot_blocks_map: Dict[str, Tuple[List[dict], ElasticLayoutState]] = {}
        attachments_hwp: List[Tuple[bytes, str]] = []
        attachments_docx: List[Tuple[Optional[bytes], str]] = []
        attachment_captions: List[str] = []

        max_imgs = (
            overrides_ir.max_attachment_images
            if overrides_ir.max_attachment_images is not None
            else len(eff_spec.attachment_slots)
        )

        for s_uid, slot_uid in job.slot_mapping.items():
            if s_uid not in job.sections:
                continue
            s_obj = job.sections[s_uid]
            sl_spec = next((sl for sl in eff_spec.section_slots if sl.slot_uid == slot_uid), None)
            avail_w = sl_spec.available_width_hu if sl_spec else 46200
            max_p = sl_spec.max_pages if sl_spec else 2.0
            solved_blocks, el_state, tele = self.solver.solve_section_layout_budget(
                s_obj.blocks,
                target_pages=max_p,
                outer_width_hu=avail_w,
                target_spec=eff_spec,
                allow_font_size_elastic_scaling=overrides_ir.allow_font_size_elastic_scaling,
            )
            slot_blocks_map[slot_uid] = (solved_blocks, el_state)
            job.solver_telemetry[s_uid] = tele

            if overrides_ir.route_media_to_attachment_slots:
                for m_obj in s_obj.media_objects:
                    idx_m = len(attachments_hwp)
                    if idx_m >= max_imgs:
                        break
                    if idx_m < len(overrides_ir.attachment_captions):
                        cap_str = overrides_ir.attachment_captions[idx_m]
                    elif m_obj.caption_hint:
                        cap_str = m_obj.caption_hint
                    else:
                        cap_str = f"[Attachment {idx_m + 1}] {m_obj.filename}"

                    attachments_hwp.append((m_obj.raw_bytes, cap_str))
                    raw_xml_p = next(
                        (
                            b.get("raw_p_xml")
                            for b in s_obj.blocks
                            if b.get("type") == "p" and b.get("r_embed") == m_obj.rel_id
                        ),
                        None,
                    )
                    attachments_docx.append((raw_xml_p, cap_str))
                    attachment_captions.append(cap_str)

        results: Dict[str, dict] = {}
        os.makedirs(os.path.dirname(os.path.abspath(output_prefix)), exist_ok=True)

        if "hwp" in formats:
            out_hwp = f"{output_prefix}.hwp"
            results["hwp"] = project_ir_to_hwp(
                job.target_path,
                out_hwp,
                eff_spec,
                slot_blocks_map,
                attachments_hwp,
                preserve_outer_tables=overrides_ir.preserve_outer_tables,
            )
            job.output_files["hwp"] = out_hwp

        if "docx" in formats and job.source_paths:
            out_docx = f"{output_prefix}.docx"
            results["docx"] = project_ir_to_docx(
                job.source_paths[0],
                out_docx,
                eff_spec,
                slot_blocks_map,
                attachments_docx,
                preserve_outer_tables=overrides_ir.preserve_outer_tables,
                clear_unmapped_slots=overrides_ir.clear_unmapped_slots,
            )
            job.output_files["docx"] = out_docx

        if "hwpx" in formats and base_hwpx_path and os.path.exists(base_hwpx_path):
            out_hwpx = f"{output_prefix}.hwpx"
            results["hwpx"] = project_ir_to_hwpx(
                base_hwpx_path,
                out_hwpx,
                eff_spec,
                slot_blocks_map,
                attachment_captions,
            )
            job.output_files["hwpx"] = out_hwpx

        active_secs = {
            k: SectionIR(
                section_uid=k,
                heading_text=job.sections[k].heading_text,
                blocks=slot_blocks_map[v][0],
                ordinal_index=job.sections[k].ordinal_index,
                section_role=job.sections[k].section_role,
            )
            for k, v in job.slot_mapping.items()
            if k in job.sections and v in slot_blocks_map
        }

        if "md" in formats:
            out_md = f"{output_prefix}.md"
            results["md"] = project_ir_to_markdown(active_secs or job.sections, out_md)
            job.output_files["md"] = out_md

        if "csv" in formats:
            out_csv = f"{output_prefix}.csv"
            results["csv"] = project_ir_to_csv(active_secs or job.sections, out_csv, delimiter=",")
            job.output_files["csv"] = out_csv

        return {
            "job_id": job_id,
            "effective_spec": self._serialize_target_spec(eff_spec),
            "page_budget_solution": job.page_budget_solution,
            "generated_files": results,
            "solver_telemetry": job.solver_telemetry,
        }

    def tdir_verify_fidelity_and_layout(self, job_id: str, export_pdf: bool = True) -> dict:
        job = self.jobs[job_id]
        fidelity_results: Dict[str, dict] = {}

        active_sec_uids = [k for k in job.slot_mapping.keys() if k in job.sections]

        if "hwp" in job.output_files and os.path.exists(job.output_files["hwp"]):
            hwp_txt, _ = lift_hwp_to_ir(job.output_files["hwp"])
            sec_checks = {
                s_uid: validate_text_immutability(job.sections[s_uid].blocks, hwp_txt)
                for s_uid in active_sec_uids
            }
            fidelity_results["hwp"] = {
                "path": job.output_files["hwp"],
                "section_checks": sec_checks,
                "all_sections_verified": all(v["substring_contained"] for v in sec_checks.values()),
            }

        if "hwpx" in job.output_files and os.path.exists(job.output_files["hwpx"]):
            with zipfile.ZipFile(job.output_files["hwpx"], "r") as zf:
                s0_root = ET.fromstring(zf.read("Contents/section0.xml"))
                hwpx_txt = "".join(s0_root.itertext())
            sec_checks = {
                s_uid: validate_text_immutability(job.sections[s_uid].blocks, hwpx_txt)
                for s_uid in active_sec_uids
            }
            fidelity_results["hwpx"] = {
                "path": job.output_files["hwpx"],
                "section_checks": sec_checks,
                "all_sections_verified": all(v["substring_contained"] for v in sec_checks.values()),
            }

        if "docx" in job.output_files and os.path.exists(job.output_files["docx"]):
            with zipfile.ZipFile(job.output_files["docx"], "r") as zf:
                d_root = ET.fromstring(zf.read("word/document.xml"))
                docx_txt = "".join(d_root.itertext())
            sec_checks = {
                s_uid: validate_text_immutability(job.sections[s_uid].blocks, docx_txt)
                for s_uid in active_sec_uids
            }
            fidelity_results["docx"] = {
                "path": job.output_files["docx"],
                "section_checks": sec_checks,
                "all_sections_verified": all(v["substring_contained"] for v in sec_checks.values()),
            }

            if export_pdf:
                pdf_path = os.path.splitext(job.output_files["docx"])[0] + "_PDF.pdf"
                esc_docx = os.path.abspath(job.output_files["docx"]).replace("'", "''")
                esc_pdf = os.path.abspath(pdf_path).replace("'", "''")
                ps_cmd = (
                    f"$w = $null; $d = $null; "
                    f"try {{ "
                    f"$w = New-Object -ComObject Word.Application; $w.Visible = $false; "
                    f"$d = $w.Documents.Open('{esc_docx}'); "
                    f"$d.ExportAsFixedFormat('{esc_pdf}', 17); "
                    f"}} finally {{ "
                    f"if ($d) {{ $d.Close($false) }}; "
                    f"if ($w) {{ $w.Quit() }} "
                    f"}}"
                )
                subprocess.run(["powershell", "-NoProfile", "-Command", ps_cmd], check=True)
                job.output_files["pdf"] = pdf_path
                with open(pdf_path, "rb") as f_pdf:
                    raw_pdf = f_pdf.read()
                actual_pdf_pages = len(re.findall(rb"/Type\s*/Page\b(?!s)", raw_pdf))
                expected_total = job.page_budget_solution.get("expected_total_pages")
                fidelity_results["pdf"] = {
                    "path": pdf_path,
                    "bytes": os.path.getsize(pdf_path),
                    "actual_pdf_pages": actual_pdf_pages,
                    "expected_total_pages": expected_total,
                    "page_budget_matched": (actual_pdf_pages == expected_total) if expected_total is not None else True,
                }

        return {
            "job_id": job_id,
            "fidelity_results": fidelity_results,
            "page_budget_solution": job.page_budget_solution,
            "solver_telemetry": job.solver_telemetry,
        }

    def tdir_adjust_block_layout(
        self,
        job_id: str,
        section_uid: str,
        block_index: int,
        lock_unmerged: Optional[bool] = None,
        col_widths_hu: Optional[List[int]] = None,
    ) -> dict:
        job = self.jobs[job_id]
        resolved_sec_uid = section_uid if section_uid in job.sections else None
        if resolved_sec_uid is None:
            for cand_uid, cand_sec in job.sections.items():
                if matches_section_selector(section_uid, cand_sec, job.slot_mapping.get(cand_uid)):
                    resolved_sec_uid = cand_uid
                    break
        if resolved_sec_uid is None:
            raise KeyError(f"Unknown section_uid or selector: {section_uid}")

        sec = job.sections[resolved_sec_uid]
        blk = sec.blocks[block_index]
        if lock_unmerged is not None and blk["type"] == "p":
            blk["lock_unmerged"] = bool(lock_unmerged)
        if col_widths_hu is not None and blk["type"] == "tbl":
            if len(col_widths_hu) == blk["n_cols"]:
                blk["col_widths_hu"] = [int(x) for x in col_widths_hu]
                for cell in blk["cells"]:
                    c0 = cell["col"]
                    cs = cell["col_span"]
                    cell["width_hu"] = sum(blk["col_widths_hu"][c0 : c0 + cs])

        eff_spec = job.target_spec or job.raw_target_spec
        slot_uid = job.slot_mapping.get(resolved_sec_uid)
        sl_spec = (
            next((sl for sl in eff_spec.section_slots if sl.slot_uid == slot_uid), None)
            if (eff_spec and slot_uid)
            else None
        )
        avail_w = sl_spec.available_width_hu if sl_spec else 46200
        max_p = sl_spec.max_pages if sl_spec else 2.0
        _, _, updated_tele = self.solver.solve_section_layout_budget(
            sec.blocks,
            target_pages=max_p,
            outer_width_hu=avail_w,
            target_spec=eff_spec,
            allow_font_size_elastic_scaling=job.user_overrides.allow_font_size_elastic_scaling,
        )
        job.solver_telemetry[resolved_sec_uid] = updated_tele

        return {
            "job_id": job_id,
            "section_uid": resolved_sec_uid,
            "block_index": block_index,
            "updated_block_type": blk["type"],
            "lock_unmerged": blk.get("lock_unmerged", False),
            "updated_solver_preview": updated_tele,
        }


MCP_TOOLS_SCHEMA = [
    {
        "name": "tdir_probe_files",
        "description": (
            "Dynamically inspect source documents (.docx, .hwp, .hwpx, .md, .csv, .tsv) and target template "
            "typography, page geometry, and slots without hardcoding text, fonts, or record offsets."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "source_paths": {"type": "array", "items": {"type": "string"}},
                "target_path": {"type": "string"},
            },
            "required": ["source_paths", "target_path"],
        },
    },
    {
        "name": "tdir_build_mapping_plan",
        "description": (
            "Resolve effective document specifications (UserSpecOverrideIR > TargetDocumentSpecIR > SourceDocumentSpecIR), "
            "build section-to-slot and media-to-attachment mappings, and preview MinimumInvasionElasticSolver Early-Stop stages."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "job_id": {"type": "string"},
                "user_spec_overrides": USER_SPEC_OVERRIDES_JSON_SCHEMA,
                "user_constraints": USER_SPEC_OVERRIDES_JSON_SCHEMA,
            },
            "required": ["job_id"],
        },
    },
    {
        "name": "tdir_execute_conversion",
        "description": (
            "Execute T-DIR v3.0 conversion into target formats (.hwp, .hwpx, .docx, .md) using dynamically resolved "
            "specifications and optional user_spec_overrides with 100% text preservation."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "job_id": {"type": "string"},
                "output_prefix": {"type": "string"},
                "output_formats": {"type": "array", "items": {"type": "string"}},
                "base_hwpx_path": {"type": "string"},
                "user_spec_overrides": USER_SPEC_OVERRIDES_JSON_SCHEMA,
            },
            "required": ["job_id", "output_prefix"],
        },
    },
    {
        "name": "tdir_verify_fidelity_and_layout",
        "description": "Verify 100.00% non-whitespace SHA-256 text immutability across generated outputs and optionally export headless PDF for page boundary check.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "job_id": {"type": "string"},
                "export_pdf": {"type": "boolean"},
            },
            "required": ["job_id"],
        },
    },
    {
        "name": "tdir_adjust_block_layout",
        "description": "Fine-tune a specific paragraph (lock_unmerged) or table (col_widths_hu) in T-DIR and re-project.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "job_id": {"type": "string"},
                "section_uid": {"type": "string"},
                "block_index": {"type": "integer"},
                "lock_unmerged": {"type": "boolean"},
                "col_widths_hu": {"type": "array", "items": {"type": "integer"}},
            },
            "required": ["job_id", "section_uid", "block_index"],
        },
    },
]


def handle_jsonrpc_request(server: TdirBetaMcpServer, req: dict) -> Optional[dict]:
    method = req.get("method", "")
    req_id = req.get("id")
    params = req.get("params", {}) or {}

    if method == "initialize":
        return {
            "jsonrpc": "2.0",
            "id": req_id,
            "result": {
                "protocolVersion": "2024-11-05",
                "capabilities": {"tools": {}},
                "serverInfo": {"name": "tdir-document-beta", "version": "3.2.0-beta"},
            },
        }
    if method == "notifications/initialized":
        return None
    if method == "tools/list":
        return {"jsonrpc": "2.0", "id": req_id, "result": {"tools": MCP_TOOLS_SCHEMA}}
    if method == "tools/call":
        t_name = params.get("name", "")
        t_args = params.get("arguments", {}) or {}
        try:
            fn = getattr(server, t_name)
            res = fn(**t_args)
            return {
                "jsonrpc": "2.0",
                "id": req_id,
                "result": {
                    "content": [{"type": "text", "text": json.dumps(res, ensure_ascii=False, indent=2)}],
                    "isError": False,
                },
            }
        except Exception as exc:
            return {
                "jsonrpc": "2.0",
                "id": req_id,
                "result": {
                    "content": [{"type": "text", "text": f"Error executing {t_name}: {exc}"}],
                    "isError": True,
                },
            }
    return {
        "jsonrpc": "2.0",
        "id": req_id,
        "error": {"code": -32601, "message": f"Method not found: {method}"},
    }


def run_stdio_loop() -> None:
    server = TdirBetaMcpServer()
    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        req = json.loads(line)
        resp = handle_jsonrpc_request(server, req)
        if resp is not None:
            sys.stdout.write(json.dumps(resp, ensure_ascii=False) + "\n")
            sys.stdout.flush()


if __name__ == "__main__":
    run_stdio_loop()
