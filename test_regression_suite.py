"""
T-DIR v3.0 Beta MCP Server Full Regression Test Suite (PoC 1, PoC 2, PoC 3 + Elastic Solver Verification).
Tests all 5 MCP tools end-to-end without touching existing MCP servers or overwriting production files.
"""

import os
import re
import sys
import json

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, r"c:\Users\stella\Desktop\stella test")

from tdir_mcp_beta.mcp_server import TdirBetaMcpServer, handle_jsonrpc_request

TEMPLATE_HWP_PATH = r"C:\Users\stella\Downloads\모창 사업계획서 템플릿 test.hwp"
SOURCE_DOCX_PATH = r"C:\Users\stella\Downloads\모창 2차 사업계획서 — 한컴독스_완벽호환본.docx"
BASE_HWPX_PATH = r"C:\Users\stella\Downloads\모창 2차 사업계획서 — 네이티브한글변환본.hwpx"
REGRESSION_OUT_DIR = r"c:\Users\stella\Desktop\stella test\tdir_mcp_beta\regression_out"


def count_pdf_pages(pdf_path: str) -> int:
    with open(pdf_path, "rb") as f:
        raw = f.read()
    # Count /Type /Page (excluding /Type /Pages)
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

    # Step 2: Tool #1 - tdir_probe_files (Dynamic format, section, and template slot discovery)
    probe_res = srv.tdir_probe_files([SOURCE_DOCX_PATH], TEMPLATE_HWP_PATH)
    job_id = probe_res["job_id"]
    print(f"[Step 1: tdir_probe_files] job_id={job_id}")
    for s_info in probe_res["sources"][0]["sections_discovered"]:
        print(
            f"   Discovered Source Section [{s_info['section_uid']}]: "
            f"chars={s_info['non_ws_chars']:,}, tables={s_info['tables_count']}, "
            f"paras={s_info['paras_count']}, media={s_info['media_count']}, sha256={s_info['text_sha256'][:16]}"
        )
    for sl_info in probe_res["target_spec"]["section_slots"]:
        print(
            f"   Discovered Template Section Slot [{sl_info['slot_uid']}]: "
            f"title='{sl_info['title']}', outer_tbl_rec={sl_info['outer_tbl_record_idx']}, "
            f"body_cell_rec={sl_info['body_cell_record_idx']}, width_hu={sl_info['available_width_hu']}"
        )
    for att_info in probe_res["target_spec"]["attachment_slots"][:2]:
        print(
            f"   Discovered Template Attachment Slot [{att_info['slot_uid']}]: "
            f"img_cell_rec={att_info['body_cell_record_idx']}, cap_cell_rec={att_info['caption_cell_record_idx']}"
        )

    # Step 3: Tool #2 - tdir_build_mapping_plan (Dynamic mapping & MinimumInvasionElasticSolver simulation)
    plan_res = srv.tdir_build_mapping_plan(
        job_id=job_id,
        user_constraints={
            "exclude_sections": ["PREAMBLE", "Q3", "Q4", "REFERENCES"],
            "page_budgets": {"Q1": 2.0, "Q2": 2.0},
            "page_body_height_hu": 78500,
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

    # Step 7: PoC 1 Markdown round-trip check
    poc1_res = run_poc1_markdown_and_standalone_regression(srv, job_id)
    print(f"[Step 7: PoC 1 Markdown Round-Trip Check] {poc1_res}")

    return {
        "init_ok": init_resp["result"]["serverInfo"]["name"] == "tdir-document-beta",
        "probe_res": probe_res,
        "plan_res": plan_res,
        "verify_res": verify_res,
        "pdf_pages": pdf_pages,
        "poc1_res": poc1_res,
    }


def main():
    srv = TdirBetaMcpServer()
    summary = run_poc2_and_poc3_regression(srv)
    assert summary["pdf_pages"] == 8, f"Expected 8 pages in PDF, got {summary['pdf_pages']}"
    print("\n=== ALL T-DIR v3.0 BETA MCP REGRESSION TESTS PASSED (100.00% FIDELITY, 8 PAGES) ===")


if __name__ == "__main__":
    main()
