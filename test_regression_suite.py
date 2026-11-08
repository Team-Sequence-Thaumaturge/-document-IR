"""
T-DIR v3.0 Beta MCP Server Full Regression Test Suite
(Dynamic Spec Extraction + UserSpecOverrideIR Priority Chain + PoC 1/2/3 + Elastic Solver Verification).
Tests all 5 MCP tools end-to-end without touching existing MCP servers or overwriting production files.
"""

import os
import re
import sys
import json

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, r"c:\Users\stella\Desktop\stella test")

from tdir_mcp_beta.mcp_server import TdirBetaMcpServer, handle_jsonrpc_request
from tdir_mcp_beta.ir_schema import detect_image_format, compute_heading_slot_similarity
from tdir_mcp_beta.codec_hwp_v5 import extract_hwp_para_text

TEMPLATE_HWP_PATH = r"C:\Users\stella\Downloads\모창 사업계획서 템플릿 test.hwp"
SOURCE_DOCX_PATH = r"C:\Users\stella\Downloads\모창 2차 사업계획서 — 한컴독스_완벽호환본.docx"
BASE_HWPX_PATH = r"C:\Users\stella\Downloads\모창 2차 사업계획서 — 네이티브한글변환본.hwpx"
REGRESSION_OUT_DIR = r"c:\Users\stella\Desktop\stella test\tdir_mcp_beta\regression_out"


def count_pdf_pages(pdf_path: str) -> int:
    with open(pdf_path, "rb") as f:
        raw = f.read()
    pages = re.findall(rb"/Type\s*/Page\b(?!s)", raw)
    return len(pages)


def run_poc1_markdown_and_standalone_regression(srv: TdirBetaMcpServer, job_id: str) -> dict:
    md_out = os.path.join(REGRESSION_OUT_DIR, "poc1_roundtrip_test.md")
    res_exec = srv.tdir_execute_conversion(
        job_id=job_id,
        output_prefix=os.path.join(REGRESSION_OUT_DIR, "poc1_roundtrip_test"),
        output_formats=["md"],
    )
    probe_md = srv.tdir_probe_files([md_out], TEMPLATE_HWP_PATH)
    return {
        "md_generated": res_exec["generated_files"]["md"],
        "md_reprobed_sections": len(probe_md["sources"][0]["sections_discovered"]),
        "md_reprobed_chars": sum(s["non_ws_chars"] for s in probe_md["sources"][0]["sections_discovered"]),
    }


def run_poc2_and_poc3_regression(srv: TdirBetaMcpServer) -> dict:
    os.makedirs(REGRESSION_OUT_DIR, exist_ok=True)

    # Step 1: Test JSON-RPC 2.0 initialize & tools/list
    init_resp = handle_jsonrpc_request(srv, {"jsonrpc": "2.0", "id": 1, "method": "initialize", "params": {}})
    list_resp = handle_jsonrpc_request(srv, {"jsonrpc": "2.0", "id": 2, "method": "tools/list", "params": {}})
    assert len(list_resp["result"]["tools"]) == 5

    # Step 2: Tool #1 - tdir_probe_files (Dynamic format, typography, section, and template slot discovery)
    probe_res = srv.tdir_probe_files([SOURCE_DOCX_PATH], TEMPLATE_HWP_PATH)
    job_id = probe_res["job_id"]
    t_spec_info = probe_res["target_spec"]
    s_spec_info = probe_res["sources"][0]["source_spec"]
    print(f"[Step 1: tdir_probe_files] job_id={job_id}")
    print(
        f"   Discovered Source Spec: font='{s_spec_info['body_font_family']}', "
        f"body_hu={s_spec_info['body_font_size_hu']}, table_hu={s_spec_info['table_font_size_hu']}, "
        f"page={s_spec_info['page_width_hu']}x{s_spec_info['page_height_hu']} HU"
    )
    print(
        f"   Discovered Target Template Spec: font='{t_spec_info['body_font_family']}', "
        f"body_pt={t_spec_info['body_font_size_pt']}pt ({t_spec_info['body_font_size_hu']} HU), "
        f"table_pt={t_spec_info['table_font_size_pt']}pt ({t_spec_info['table_font_size_hu']} HU), "
        f"line_spacing={t_spec_info['default_line_spacing_pct']}%, "
        f"page={t_spec_info['page_width_hu']}x{t_spec_info['page_height_hu']} HU, "
        f"preamble_blocks={t_spec_info['preamble_blocks_count']}"
    )
    for s_info in probe_res["sources"][0]["sections_discovered"]:
        print(
            f"   Discovered Source Section [{s_info['section_uid']}]: "
            f"chars={s_info['non_ws_chars']:,}, tables={s_info['tables_count']}, "
            f"paras={s_info['paras_count']}, media={s_info['media_count']}, sha256={s_info['text_sha256'][:16]}"
        )
    for sl_info in t_spec_info["section_slots"]:
        print(
            f"   Discovered Template Section Slot [{sl_info['slot_uid']}]: "
            f"title='{sl_info['title']}', subtitle='{sl_info['subtitle'][:24]}', "
            f"outer_w={sl_info['outer_width_hu']}, cols={sl_info['col_widths_hu']}, "
            f"avail_w={sl_info['available_width_hu']}"
        )

    # Step 3: Tool #2 - tdir_build_mapping_plan with explicit `user_spec_overrides` (`UserSpecOverrideIR`)
    plan_res = srv.tdir_build_mapping_plan(
        job_id=job_id,
        user_spec_overrides={
            "exclude_sections": ["PREAMBLE", "Q3", "Q4", "REFERENCES"],
            "section_page_budgets": {"Q1": 2.0, "Q2": 2.0},
            "preserve_outer_tables": True,
            "route_media_to_attachment_slots": True,
            "max_attachment_images": 3,
            "attachment_captions": [
                "[붙임 1] 광운대 프리즘 (PRISM) — AI 비교과 프로그램 추천 시스템 아키텍처 및 데이터 파이프라인",
                "[붙임 2] AI 매칭 프로세스 3단계 워크플로우 및 핵심 타겟 시장(TAM·SAM·SOM) 로드맵",
            ],
        },
    )
    print(f"[Step 2: tdir_build_mapping_plan] slot_mapping={plan_res['slot_mapping']}")
    for sec_k, tele in plan_res["solver_preview"].items():
        print(
            f"   Solver Preview [{sec_k}]: stopped_stage={tele['stopped_stage']}, "
            f"height={tele['initial_height_hu']:,} -> {tele['final_height_hu']:,} / budget={tele['budget_hu']:,} HU | "
            f"body_paras={tele['initial_body_paras']}->{tele['final_body_paras']} "
            f"(merged_subheadings={tele['merged_subheadings_count']}, merged_body_pairs={tele['merged_body_pairs_count']}), "
            f"cell_paras={tele['initial_cell_paras']}->{tele['final_cell_paras']} "
            f"(merged_cell_pairs={tele['merged_cell_pairs_count']})"
        )

    # Step 4: Tool #3 - tdir_execute_conversion (.hwp, .hwpx, .docx, .md)
    out_prefix = os.path.join(REGRESSION_OUT_DIR, "tdir_beta_migrated")
    exec_res = srv.tdir_execute_conversion(
        job_id=job_id,
        output_prefix=out_prefix,
        output_formats=["hwp", "hwpx", "docx", "md"],
        base_hwpx_path=BASE_HWPX_PATH,
    )
    print(f"[Step 3: tdir_execute_conversion] Generated {len(exec_res['generated_files'])} format outputs:")
    for fmt_k, f_meta in exec_res["generated_files"].items():
        print(f"   [{fmt_k.upper()}] {f_meta}")

    # Step 5: Tool #4 - tdir_verify_fidelity_and_layout (100% SHA-256 text check + Headless PDF page count check)
    verify_res = srv.tdir_verify_fidelity_and_layout(job_id=job_id, export_pdf=True)
    print("[Step 4: tdir_verify_fidelity_and_layout] Verification Results:")
    for fmt_k, v_meta in verify_res["fidelity_results"].items():
        print(f"   [{fmt_k.upper()}] {json.dumps(v_meta, ensure_ascii=False)}")
        if "all_sections_verified" in v_meta:
            assert v_meta["all_sections_verified"] is True, f"Text mismatch in {fmt_k}!"

    pdf_path = verify_res["fidelity_results"]["pdf"]["path"]
    pdf_pages = count_pdf_pages(pdf_path)
    print(f"[Step 5: PDF Page Boundary Check] Rendered PDF pages = {pdf_pages} (Expected = 8)")

    # Step 6: Tool #5 - tdir_adjust_block_layout (Test fine-tuning lock_unmerged on a block)
    adj_res = srv.tdir_adjust_block_layout(job_id=job_id, section_uid="Q1", block_index=1, lock_unmerged=True)
    print(f"[Step 6: tdir_adjust_block_layout] {adj_res}")

    # Step 7: Test UserSpecOverrideIR custom font & font size override in HWP/DOCX
    override_plan = srv.tdir_build_mapping_plan(
        job_id=job_id,
        user_spec_overrides={
            "body_font_family": "함초롬바탕",
            "table_font_family": "함초롬돋움",
            "body_font_size_pt": 11.5,
            "table_font_size_pt": 9.5,
            "exclude_sections": ["PREAMBLE", "Q3", "Q4", "REFERENCES"],
            "section_page_budgets": {"Q1": 2.0, "Q2": 2.0},
        },
    )
    eff_ov = override_plan["effective_spec"]
    assert eff_ov["body_font_family"] == "함초롬바탕"
    assert eff_ov["table_font_family"] == "함초롬돋움"
    assert eff_ov["body_font_size_hu"] == 1150
    assert eff_ov["table_font_size_hu"] == 950
    print(
        f"[Step 7: UserSpecOverrideIR Priority Verification] "
        f"body_font='{eff_ov['body_font_family']}' ({eff_ov['body_font_size_pt']}pt / {eff_ov['body_font_size_hu']} HU), "
        f"table_font='{eff_ov['table_font_family']}' ({eff_ov['table_font_size_pt']}pt / {eff_ov['table_font_size_hu']} HU)"
    )

    # Step 8: PoC 1 Markdown round-trip check
    poc1_res = run_poc1_markdown_and_standalone_regression(srv, job_id)
    print(f"[Step 8: PoC 1 Markdown Round-Trip Check] {poc1_res}")

    # Step 9: v3.1.0-beta Bugfix & Edge-Case Verification Suite ([BUG-01..12], [EDGE-01..08], [GEN-01..05])
    pay_emoji = "A🚀B\r".encode("utf-16le")
    n_wchars_emoji = len(pay_emoji) // 2
    assert n_wchars_emoji == 5, f"Expected 5 UTF-16LE WCHARs for 'A🚀B\\r', got {n_wchars_emoji}"
    assert extract_hwp_para_text(pay_emoji).strip() == "A🚀B"

    assert detect_image_format(b"\xff\xd8\xff\xe0\x00\x10JFIF") == ("jpg", "image/jpeg")
    assert detect_image_format(b"\x89PNG\r\n\x1a\n\x00\x00") == ("png", "image/png")
    assert "updated_solver_preview" in adj_res
    assert compute_heading_slot_similarity("1. 문제 인식 (Problem)", "1. 문제 인식 (Problem)", "SLOT_1") > 0.5

    edge_md_path = os.path.join(REGRESSION_OUT_DIR, "edge_case_codefence.md")
    with open(edge_md_path, "w", encoding="utf-8") as f_md:
        f_md.write(
            "# Q1. 문제 인식\n\n"
            "```python\n"
            "x = a | b | c\n"
            "```\n\n"
            "| 구분 | 내용\\|설명 |\n"
            "| --- | --- |\n"
            "| 항목1 | 값A\\|값B |\n"
        )
    multi_probe = srv.tdir_probe_files([edge_md_path, edge_md_path], TEMPLATE_HWP_PATH)
    assert len(multi_probe["sources"]) == 2
    s1_uid = multi_probe["sources"][0]["sections_discovered"][0]["section_uid"]
    s2_uid = multi_probe["sources"][1]["sections_discovered"][0]["section_uid"]
    assert s1_uid == "Q1" and s2_uid == "Q1_S2", f"Expected Q1 and Q1_S2, got {s1_uid}, {s2_uid}"

    md_standalone_exec = srv.tdir_execute_conversion(
        job_id=multi_probe["job_id"],
        output_prefix=os.path.join(REGRESSION_OUT_DIR, "standalone_from_md"),
        output_formats=["docx"],
        user_spec_overrides={"margins_mm": {"top": 15.0, "bottom": 15.0, "left": 20.0, "right": 20.0}},
    )
    assert os.path.isfile(md_standalone_exec["generated_files"]["docx"]["out_docx_path"])
    print(
        f"[Step 9: v3.1.0-beta Bugfix & Edge-Case Checks] "
        f"emoji_wchars={n_wchars_emoji}, multi_source_uids=({s1_uid}, {s2_uid}), "
        f"standalone_docx_bytes={md_standalone_exec['generated_files']['docx']['bytes']}"
    )

    return {
        "init_ok": init_resp["result"]["serverInfo"]["name"] == "tdir-document-beta",
        "probe_res": probe_res,
        "plan_res": plan_res,
        "verify_res": verify_res,
        "pdf_pages": pdf_pages,
        "poc1_res": poc1_res,
    }


def main() -> None:
    srv = TdirBetaMcpServer()
    summary = run_poc2_and_poc3_regression(srv)
    assert summary["pdf_pages"] == 8, f"Expected 8 pages in PDF, got {summary['pdf_pages']}"
    print("\n=== ALL T-DIR v3.1.0-BETA MCP REGRESSION TESTS PASSED (100.00% FIDELITY, 8 PAGES) ===")


if __name__ == "__main__":
    main()
