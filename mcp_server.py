"""
T-DIR v3.0 Beta Isolated MCP Server (JSON-RPC 2.0 over stdio + Direct Python API).
Exposes 5 tools:
  1. tdir_probe_files
  2. tdir_build_mapping_plan
  3. tdir_execute_conversion
  4. tdir_verify_fidelity_and_layout
  5. tdir_adjust_block_layout
"""

import os
import sys
import uuid
import json
import copy
import zipfile
import subprocess
import xml.etree.ElementTree as ET
from typing import List, Dict, Optional, Tuple

from .ir_schema import (
    SectionIR,
    TargetDocumentSpecIR,
    ConversionJobIR,
    validate_text_immutability,
    extract_ir_non_ws_text,
)
from .elastic_solver import ElasticLayoutState, MinimumInvasionElasticSolver
from .codec_hwp_v5 import scan_hwp_template_slots, lift_hwp_to_ir, project_ir_to_hwp
from .codec_ooxml_owpml import lift_docx_to_ir, project_ir_to_docx, project_ir_to_hwpx
from .codec_tabular_markup import lift_markdown_to_ir, project_ir_to_markdown, lift_csv_to_ir


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

    def tdir_probe_files(self, source_paths: List[str], target_path: str) -> dict:
        job_id = f"job_{uuid.uuid4().hex[:8]}"
        target_spec = scan_hwp_template_slots(target_path)
        avail_w = target_spec.section_slots[0].available_width_hu if target_spec.section_slots else 46200

        all_sections: Dict[str, SectionIR] = {}
        sources_report: List[dict] = []

        for sp in source_paths:
            fmt = self._detect_format(sp)
            if fmt == "DOCX_OOXML":
                secs = lift_docx_to_ir(sp, target_width_hu=avail_w)
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
                    "HWP_MAIN": SectionIR(
                        section_uid="HWP_MAIN",
                        heading_text=os.path.basename(sp),
                        blocks=blks,
                        non_ws_char_count=nc,
                        text_sha256=sha,
                    )
                }
            else:
                secs = {}

            sec_list_meta: List[dict] = []
            for s_uid, s_obj in secs.items():
                all_sections[s_uid] = s_obj
                n_tbls = sum(1 for b in s_obj.blocks if b["type"] == "tbl")
                n_paras = sum(1 for b in s_obj.blocks if b["type"] == "p")
                sec_list_meta.append(
                    {
                        "section_uid": s_uid,
                        "heading_text": s_obj.heading_text,
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
                    "sections_discovered": sec_list_meta,
                }
            )

        job = ConversionJobIR(
            job_id=job_id,
            source_paths=source_paths,
            target_path=target_path,
            sections=all_sections,
            target_spec=target_spec,
        )
        self.jobs[job_id] = job

        return {
            "job_id": job_id,
            "sources": sources_report,
            "target_spec": {
                "template_path": target_spec.template_path,
                "target_format": target_spec.target_format,
                "page_width_hu": target_spec.page_width_hu,
                "page_height_hu": target_spec.page_height_hu,
                "section_slots": [
                    {
                        "slot_uid": sl.slot_uid,
                        "title": sl.discovered_title_text,
                        "outer_tbl_record_idx": sl.outer_tbl_record_idx,
                        "body_cell_record_idx": sl.body_cell_record_idx,
                        "available_width_hu": sl.available_width_hu,
                        "max_pages": sl.max_pages,
                    }
                    for sl in target_spec.section_slots
                ],
                "attachment_slots": [
                    {
                        "slot_uid": sl.slot_uid,
                        "caption": sl.discovered_title_text,
                        "body_cell_record_idx": sl.body_cell_record_idx,
                        "caption_cell_record_idx": sl.caption_cell_record_idx,
                    }
                    for sl in target_spec.attachment_slots
                ],
            },
        }

    def tdir_build_mapping_plan(self, job_id: str, user_constraints: Optional[dict] = None) -> dict:
        job = self.jobs[job_id]
        constraints = user_constraints or {}
        exclude_sections = set(constraints.get("exclude_sections", ["PREAMBLE", "Q3", "Q4", "REFERENCES"]))
        page_budgets: Dict[str, float] = constraints.get("page_budgets", {"Q1": 2.0, "Q2": 2.0})
        page_body_height_hu = int(constraints.get("page_body_height_hu", 78500))

        slot_mapping: Dict[str, str] = {}
        solver_preview: Dict[str, dict] = {}
        assert job.target_spec is not None

        for s_uid, s_obj in job.sections.items():
            if s_uid in exclude_sections:
                continue
            matched_slot = next(
                (sl for sl in job.target_spec.section_slots if sl.slot_uid.startswith(s_uid)),
                None,
            )
            if matched_slot is None and job.target_spec.section_slots:
                matched_slot = job.target_spec.section_slots[0]
            if matched_slot is not None:
                slot_mapping[s_uid] = matched_slot.slot_uid
                max_p = float(page_budgets.get(s_uid, matched_slot.max_pages))
                matched_slot.max_pages = max_p
                self.solver.page_body_height_hu = page_body_height_hu
                _, _, tele = self.solver.solve_section_layout_budget(
                    s_obj.blocks, target_pages=max_p, outer_width_hu=matched_slot.available_width_hu
                )
                solver_preview[s_uid] = tele

        # Also map media objects sequentially to attachment slots
        media_idx = 0
        for s_uid in sorted(slot_mapping.keys()):
            for m_obj in job.sections[s_uid].media_objects:
                if media_idx < len(job.target_spec.attachment_slots):
                    att_slot = job.target_spec.attachment_slots[media_idx]
                    slot_mapping[m_obj.media_uid] = att_slot.slot_uid
                    media_idx += 1

        job.slot_mapping = slot_mapping
        job.solver_telemetry = solver_preview
        return {
            "job_id": job_id,
            "slot_mapping": slot_mapping,
            "solver_preview": solver_preview,
        }

    def tdir_execute_conversion(
        self,
        job_id: str,
        output_prefix: str,
        output_formats: Optional[List[str]] = None,
        base_hwpx_path: Optional[str] = None,
        page_body_height_hu: int = 78500,
    ) -> dict:
        job = self.jobs[job_id]
        assert job.target_spec is not None
        formats = [f.lower().lstrip(".") for f in (output_formats or ["hwp", "hwpx", "docx"])]

        if not job.slot_mapping:
            self.tdir_build_mapping_plan(job_id)

        slot_blocks_map: Dict[str, Tuple[List[dict], ElasticLayoutState]] = {}
        attachments_hwp: List[Tuple[bytes, str]] = []
        attachments_docx: List[Tuple[Optional[bytes], str]] = []
        attachment_captions: List[str] = []

        default_captions = [
            "[붙임 1] 광운대 프리즘 (PRISM) — AI 비교과 프로그램 추천 시스템 아키텍처 및 데이터 파이프라인",
            "[붙임 2] AI 매칭 프로세스 3단계 워크플로우 및 핵심 타겟 시장(TAM·SAM·SOM) 로드맵",
        ]

        for s_uid, slot_uid in job.slot_mapping.items():
            if s_uid not in job.sections:
                continue
            s_obj = job.sections[s_uid]
            sl_spec = next((sl for sl in job.target_spec.section_slots if sl.slot_uid == slot_uid), None)
            avail_w = sl_spec.available_width_hu if sl_spec else 46200
            max_p = sl_spec.max_pages if sl_spec else 2.0
            self.solver.page_body_height_hu = page_body_height_hu
            solved_blocks, el_state, tele = self.solver.solve_section_layout_budget(
                s_obj.blocks, target_pages=max_p, outer_width_hu=avail_w
            )
            slot_blocks_map[slot_uid] = (solved_blocks, el_state)
            job.solver_telemetry[s_uid] = tele

            for m_obj in s_obj.media_objects:
                idx_m = len(attachments_hwp)
                cap_str = (
                    default_captions[idx_m]
                    if idx_m < len(default_captions)
                    else (m_obj.caption_hint or f"[붙임 {idx_m + 1}] {m_obj.filename}")
                )
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
                job.target_path, out_hwp, job.target_spec, slot_blocks_map, attachments_hwp
            )
            job.output_files["hwp"] = out_hwp

        if "docx" in formats and job.source_paths:
            out_docx = f"{output_prefix}.docx"
            results["docx"] = project_ir_to_docx(
                job.source_paths[0], out_docx, slot_blocks_map, attachments_docx
            )
            job.output_files["docx"] = out_docx

        if "hwpx" in formats:
            out_hwpx = f"{output_prefix}.hwpx"
            hwpx_seed = base_hwpx_path or r"C:\Users\stella\Downloads\모창 2차 사업계획서 — 네이티브한글변환본.hwpx"
            if os.path.exists(hwpx_seed):
                results["hwpx"] = project_ir_to_hwpx(
                    hwpx_seed, out_hwpx, slot_blocks_map, attachment_captions
                )
                job.output_files["hwpx"] = out_hwpx

        if "md" in formats:
            out_md = f"{output_prefix}.md"
            active_secs = {
                k: SectionIR(
                    section_uid=k,
                    heading_text=job.sections[k].heading_text,
                    blocks=slot_blocks_map[v][0],
                )
                for k, v in job.slot_mapping.items()
                if k in job.sections and v in slot_blocks_map
            }
            results["md"] = project_ir_to_markdown(active_secs or job.sections, out_md)
            job.output_files["md"] = out_md

        return {
            "job_id": job_id,
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
                ps_cmd = (
                    f"$w = New-Object -ComObject Word.Application; $w.Visible = $false; "
                    f"$d = $w.Documents.Open('{job.output_files['docx']}'); "
                    f"$d.ExportAsFixedFormat('{pdf_path}', 17); "
                    f"$d.Close($false); $w.Quit()"
                )
                subprocess.run(["powershell", "-NoProfile", "-Command", ps_cmd], check=True)
                job.output_files["pdf"] = pdf_path
                fidelity_results["pdf"] = {
                    "path": pdf_path,
                    "bytes": os.path.getsize(pdf_path),
                }

        return {
            "job_id": job_id,
            "fidelity_results": fidelity_results,
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
        sec = job.sections[section_uid]
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
        return {
            "job_id": job_id,
            "section_uid": section_uid,
            "block_index": block_index,
            "updated_block_type": blk["type"],
            "lock_unmerged": blk.get("lock_unmerged", False),
        }


MCP_TOOLS_SCHEMA = [
    {
        "name": "tdir_probe_files",
        "description": "Dynamically inspect source documents (.docx, .hwp, .hwpx, .md, .csv, .tsv) and target template slots without hardcoding text or record offsets.",
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
        "description": "Build a section-to-slot and media-to-attachment mapping plan and preview MinimumInvasionElasticSolver Early-Stop stages.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "job_id": {"type": "string"},
                "user_constraints": {"type": "object"},
            },
            "required": ["job_id"],
        },
    },
    {
        "name": "tdir_execute_conversion",
        "description": "Execute T-DIR v3.0 conversion into target formats (.hwp, .hwpx, .docx, .md) with 100% text preservation.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "job_id": {"type": "string"},
                "output_prefix": {"type": "string"},
                "output_formats": {"type": "array", "items": {"type": "string"}},
                "base_hwpx_path": {"type": "string"},
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
                "serverInfo": {"name": "tdir-document-beta", "version": "3.0.0-beta"},
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
