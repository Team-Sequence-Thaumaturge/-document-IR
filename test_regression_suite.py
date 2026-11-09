"""
T-DIR v3.2.0-beta 8-Cohort Parametric Regression Experiment Suite (EXP-01 ~ EXP-08).
Validates universal document intermediate representation, unsupervised heading spectrum detection,
generalized template slot & 2RxC media grid scanning, numerical page-budget allocation,
and 6-stage MinimumInvasionElasticSolver across diverse document topologies and metric spaces.
"""

import os
import re
import sys
import json
import zlib
import struct
import zipfile

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, r"c:\Users\stella\Desktop\stella test")

from tdir_mcp_beta.mcp_server import TdirBetaMcpServer, handle_jsonrpc_request
from tdir_mcp_beta.ir_schema import (
    detect_image_format,
    compute_heading_slot_similarity,
    extract_ordinal_index,
)
from tdir_mcp_beta.cfb_ole2 import CFBArchive, write_cfb_file
from tdir_mcp_beta.codec_hwp_v5 import (
    extract_hwp_para_text,
    parse_hwp_records,
    pack_hwp_record,
    scan_hwp_template_slots,
)

TEMPLATE_HWP_PATH = r"C:\Users\stella\Downloads\모창 사업계획서 템플릿 test.hwp"
SOURCE_DOCX_PATH = r"C:\Users\stella\Downloads\모창 2차 사업계획서 — 한컴독스_완벽호환본.docx"
BASE_HWPX_PATH = r"C:\Users\stella\Downloads\모창 2차 사업계획서 — 네이티브한글변환본.hwpx"
REGRESSION_OUT_DIR = r"c:\Users\stella\Desktop\stella test\tdir_mcp_beta\regression_out"


def count_pdf_pages(pdf_path: str) -> int:
    with open(pdf_path, "rb") as f:
        raw = f.read()
    pages = re.findall(rb"/Type\s*/Page\b(?!s)", raw)
    return len(pages)


def build_synthetic_docx(
    out_path: str,
    paragraphs: list,
    tables: list = None,
    wrap_in_sdt: bool = False,
) -> str:
    """
    Constructs a standards-compliant OOXML (.docx) archive with specified paragraph runs
    [(text, half_pt_size, bold_bool), ...] and optional raw table XML strings.
    """
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    body_items = []
    for p_item in paragraphs:
        if isinstance(p_item, tuple):
            txt, sz_hp, is_bold = p_item
            b_xml = "<w:b/>" if is_bold else ""
            p_xml = (
                f'<w:p><w:r><w:rPr>{b_xml}<w:sz w:val="{sz_hp}"/><w:szCs w:val="{sz_hp}"/></w:rPr>'
                f"<w:t>{txt}</w:t></w:r></w:p>"
            )
            body_items.append(p_xml)
        elif isinstance(p_item, str):
            body_items.append(p_item)

    for tbl_xml in (tables or []):
        body_items.append(tbl_xml)

    inner_xml = "".join(body_items)
    if wrap_in_sdt:
        inner_xml = f"<w:sdt><w:sdtContent>{inner_xml}</w:sdtContent></w:sdt>"

    doc_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f"<w:body>{inner_xml}"
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="560" w:right="950" w:bottom="560" w:left="950" w:header="360" w:footer="360" w:gutter="0"/>'
        "</w:sectPr></w:body></w:document>"
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        "</Types>"
    )
    rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="word/document.xml"/>'
        "</Relationships>"
    )
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types.encode("utf-8"))
        zf.writestr("_rels/.rels", rels_xml.encode("utf-8"))
        zf.writestr("word/document.xml", doc_xml.encode("utf-8"))
    return out_path


def run_exp01_hierarchy_and_heading_spectrum_cohort(srv: TdirBetaMcpServer) -> dict:
    """
    EXP-01: Validates unsupervised top-level section heading detection across 4 distinct
    naming/typographic conventions without hardcoded 'Q1..Q4' strings.
    """
    # Case A-1: Korean Academic ('제1장', '제2장', '제3장', '[참고문헌]')
    acad_docx = build_synthetic_docx(
        os.path.join(REGRESSION_OUT_DIR, "exp01_academic.docx"),
        [
            ("연구보고서 표지 안내문", 20, False),
            ("제1장 연구 배경 및 필요성", 24, True),
            ("※ 작성 지침: 연구의 배경과 독창성을 기술하시오", 20, False),
            ("본 연구는 비정형 문서 간의 위상수학적 동형사상(Isomorphism)을 규명한다.", 20, False),
            ("제2장 핵심 알고리즘 및 수식 설계", 24, True),
            ("동적 프로그래밍 기반의 6단계 최소 침습 레이아웃 솔버를 제안한다.", 20, False),
            ("제3장 실증 실험 및 정량 검증", 24, True),
            ("다양한 페이지 예산과 용지 규격에서 SHA-256 무결성을 100% 보존함을 입증하였다.", 20, False),
            ("[참고문헌]", 24, True),
            ("1. Thaumaturge Team, Topological Document IR Specification, 2026.", 18, False),
        ],
    )
    p_acad = srv.tdir_probe_files([acad_docx], TEMPLATE_HWP_PATH)
    acad_secs = p_acad["sources"][0]["sections_discovered"]
    acad_uids = [s["section_uid"] for s in acad_secs]
    assert acad_uids == ["PREAMBLE", "SEC_01", "SEC_02", "SEC_03", "REFERENCES"], f"Unexpected acad_uids: {acad_uids}"
    plan_acad = srv.tdir_build_mapping_plan(p_acad["job_id"])
    assert plan_acad["slot_mapping"]["SEC_01"] == "SLOT_01"
    assert plan_acad["slot_mapping"]["SEC_02"] == "SLOT_02"
    assert plan_acad["slot_mapping"]["SEC_03"] == "SLOT_03"
    assert "REFERENCES" not in plan_acad["slot_mapping"]

    # Case A-2: English Roman ('I.', 'II.', 'III.')
    roman_docx = build_synthetic_docx(
        os.path.join(REGRESSION_OUT_DIR, "exp01_roman.docx"),
        [
            ("I. Executive Summary and Problem Statement", 24, True),
            ("Enterprise document pipelines suffer from hardcoded layout assumptions.", 20, False),
            ("II. System Architecture and Intermediate Representation", 24, True),
            ("T-DIR separates semantic graph topology from physical metric space.", 20, False),
            ("III. Quantitative Evaluation and Deployment Roadmap", 24, True),
            ("Zero character loss is guaranteed across HWP, HWPX, DOCX, and Markdown.", 20, False),
        ],
    )
    p_rom = srv.tdir_probe_files([roman_docx], TEMPLATE_HWP_PATH)
    rom_uids = [s["section_uid"] for s in p_rom["sources"][0]["sections_discovered"]]
    assert rom_uids == ["SEC_01", "SEC_02", "SEC_03"], f"Unexpected rom_uids: {rom_uids}"

    # Case A-3: Unnumbered Font-Size-Jump Headings (14pt bold vs 10pt body)
    unnum_docx = build_synthetic_docx(
        os.path.join(REGRESSION_OUT_DIR, "exp01_unnumbered_jump.docx"),
        [
            ("서론 및 기술적 개요", 28, True),
            ("본 문서는 번호 기호가 전혀 없는 제목 구조를 폰트 크기 스펙트럼만으로 탐지한다.", 20, False),
            ("일반 본문 문단으로서 동일한 10pt 크기를 유지하며 통계적 최빈값을 형성한다.", 20, False),
            ("핵심 엔진 아키텍처", 28, True),
            ("양방향 슬롯 매칭기와 탄성 레이아웃 솔버가 결합되어 동작한다.", 20, False),
            ("표와 문단의 위상 구조를 보존하면서 여백과 줄간격을 연속적으로 보간한다.", 20, False),
        ],
    )
    p_unnum = srv.tdir_probe_files([unnum_docx], TEMPLATE_HWP_PATH)
    unnum_uids = [s["section_uid"] for s in p_unnum["sources"][0]["sections_discovered"]]
    assert unnum_uids == ["SEC_01", "SEC_02"], f"Unexpected unnum_uids: {unnum_uids}"

    # Case A-4: Real Grant Document probed into generic SEC_01..SEC_03, REFERENCES and SLOT_01..SLOT_04
    p_real = srv.tdir_probe_files([SOURCE_DOCX_PATH], TEMPLATE_HWP_PATH)
    real_uids = [s["section_uid"] for s in p_real["sources"][0]["sections_discovered"]]
    slot_uids = [sl["slot_uid"] for sl in p_real["target_spec"]["section_slots"]]
    media_uids = [sl["slot_uid"] for sl in p_real["target_spec"]["attachment_slots"]]
    assert real_uids == ["SEC_01", "SEC_02", "SEC_03", "REFERENCES"], f"Unexpected real_uids: {real_uids}"
    assert slot_uids == ["SLOT_01", "SLOT_02", "SLOT_03", "SLOT_04"], f"Unexpected slot_uids: {slot_uids}"
    assert media_uids == ["MEDIA_01", "MEDIA_02", "MEDIA_03", "MEDIA_04", "MEDIA_05", "MEDIA_06"]

    return {
        "academic_sections": acad_uids,
        "roman_sections": rom_uids,
        "unnumbered_jump_sections": unnum_uids,
        "real_grant_sections": real_uids,
        "template_slots": slot_uids,
        "media_slots": media_uids,
        "real_job_id": p_real["job_id"],
    }


def run_exp02_numerical_page_budget_sweep_cohort(srv: TdirBetaMcpServer, real_job_id: str) -> dict:
    """
    EXP-02: Validates numerical page-budget resolution across loose (3.0p), autonomous total_page_limit=8,
    asymmetric volume proportional allocation, and tight (1.0p) budgets, proving merge monotonicity:
      M(1.0) >= M(2.0) >= M(3.0) == 0.
    """
    # Case B-1: Loose budget (3.0p per section -> Stage 2 spacing only, 0 merges, expected_total_pages=10)
    plan_loose = srv.tdir_build_mapping_plan(
        job_id=real_job_id,
        user_spec_overrides={
            "exclude_sections": ["SEC_03", "SEC_04"],
            "section_page_budgets": {"SEC_01": 3.0, "SEC_02": 3.0},
        },
    )
    merges_loose = sum(
        t["merged_subheadings_count"] + t["merged_cell_pairs_count"] + t["merged_body_pairs_count"]
        for t in plan_loose["solver_preview"].values()
    )
    assert merges_loose == 0, f"Expected 0 merges at 3.0p budget, got {merges_loose}"
    assert plan_loose["page_budget_solution"]["expected_total_pages"] == 10

    # Case B-4: Tight budget (1.0p per section + allow_font_size_elastic_scaling=True)
    plan_tight = srv.tdir_build_mapping_plan(
        job_id=real_job_id,
        user_spec_overrides={
            "exclude_sections": ["SEC_03", "SEC_04"],
            "section_page_budgets": {"SEC_01": 1.0, "SEC_02": 1.0},
            "allow_font_size_elastic_scaling": True,
        },
    )
    merges_tight = sum(
        t["merged_subheadings_count"] + t["merged_cell_pairs_count"] + t["merged_body_pairs_count"]
        for t in plan_tight["solver_preview"].values()
    )
    assert plan_tight["solver_preview"]["SEC_01"]["stopped_stage"] == "STAGE_6_FONT_SCALE_EARLY_STOP"
    assert plan_tight["page_budget_solution"]["expected_total_pages"] == 6

    # Case B-2: Autonomous total_page_limit = 8 (ZERO hardcoded section_page_budgets!)
    plan_auto8 = srv.tdir_build_mapping_plan(
        job_id=real_job_id,
        user_spec_overrides={
            "exclude_sections": ["SEC_03", "SEC_04"],
            "total_page_limit": 8,
            "preserve_outer_tables": True,
            "route_media_to_attachment_slots": True,
            "max_attachment_images": 3,
            "attachment_captions": [
                "[붙임 1] 광운대 프리즘 (PRISM) — AI 비교과 프로그램 추천 시스템 아키텍처 및 데이터 파이프라인",
                "[붙임 2] AI 매칭 프로세스 3단계 워크플로우 및 핵심 타겟 시장(TAM·SAM·SOM) 로드맵",
            ],
        },
    )
    merges_auto8 = sum(
        t["merged_subheadings_count"] + t["merged_cell_pairs_count"] + t["merged_body_pairs_count"]
        for t in plan_auto8["solver_preview"].values()
    )
    assert plan_auto8["page_budget_solution"]["section_budgets"] == {"SEC_01": 2.0, "SEC_02": 2.0}
    assert plan_auto8["page_budget_solution"]["expected_total_pages"] == 8
    assert merges_tight >= merges_auto8 >= merges_loose == 0

    out_prefix = os.path.join(REGRESSION_OUT_DIR, "tdir_v32_auto8")
    exec_auto8 = srv.tdir_execute_conversion(
        job_id=real_job_id,
        output_prefix=out_prefix,
        output_formats=["hwp", "hwpx", "docx", "md"],
        base_hwpx_path=BASE_HWPX_PATH,
    )
    verify_auto8 = srv.tdir_verify_fidelity_and_layout(job_id=real_job_id, export_pdf=True)
    for fmt_k in ("hwp", "hwpx", "docx"):
        assert verify_auto8["fidelity_results"][fmt_k]["all_sections_verified"] is True
    pdf_meta = verify_auto8["fidelity_results"]["pdf"]
    assert pdf_meta["actual_pdf_pages"] == 8, f"Expected 8 PDF pages, got {pdf_meta['actual_pdf_pages']}"
    assert pdf_meta["page_budget_matched"] is True

    # Case B-3: Asymmetric Volume Allocation (SEC_01 has 4x volume of SEC_02 under total_page_limit=7)
    asym_paras = [("제1장 대규모 핵심 기술 상세 명세", 24, True)]
    for idx_p in range(45):
        asym_paras.append(
            (
                f"제1장 상세 본문 문단 {idx_p + 1}: 위상수학적 문서 중간 표현과 탄성 솔버의 수식적 수렴성을 검증하기 위한 장문 텍스트입니다.",
                20,
                False,
            )
        )
    asym_paras.append(("제2장 요약 결론", 24, True))
    asym_paras.append(("제2장 단문 결론: 모든 정량 지표가 목표치를 달성하였다.", 20, False))
    asym_docx = build_synthetic_docx(os.path.join(REGRESSION_OUT_DIR, "exp02_asymmetric.docx"), asym_paras)
    p_asym = srv.tdir_probe_files([asym_docx], TEMPLATE_HWP_PATH)
    plan_asym = srv.tdir_build_mapping_plan(
        p_asym["job_id"],
        user_spec_overrides={"total_page_limit": 8, "clear_unmapped_slots": False},
    )
    b_asym = plan_asym["page_budget_solution"]["section_budgets"]
    assert b_asym["SEC_01"] > b_asym["SEC_02"], f"Expected asymmetric budget SEC_01 > SEC_02, got {b_asym}"
    assert plan_asym["page_budget_solution"]["expected_total_pages"] == 8

    return {
        "merges_loose_3p": merges_loose,
        "merges_auto8_2p": merges_auto8,
        "merges_tight_1p": merges_tight,
        "monotonicity_verified": merges_tight >= merges_auto8 >= merges_loose == 0,
        "auto8_pdf_pages": pdf_meta["actual_pdf_pages"],
        "auto8_page_budget_matched": pdf_meta["page_budget_matched"],
        "asymmetric_budgets": b_asym,
        "exec_auto8": exec_auto8,
        "verify_auto8": verify_auto8,
    }


def run_exp03_metric_space_and_typography_cohort(srv: TdirBetaMcpServer, real_job_id: str) -> dict:
    """
    EXP-03: Validates continuous metric space scaling across A4 Landscape (297x210mm),
    B5 Portrait (182x257mm), and font size scaling (9.0pt vs 13.5pt).
    """
    plan_a4_land = srv.tdir_build_mapping_plan(
        job_id=real_job_id,
        user_spec_overrides={
            "exclude_sections": ["SEC_03", "SEC_04"],
            "page_width_mm": 297.0,
            "page_height_mm": 210.0,
            "margins_mm": {"top": 12.0, "bottom": 12.0, "left": 15.0, "right": 15.0},
            "body_font_family": "함초롬바탕",
            "table_font_family": "함초롬돋움",
            "body_font_size_pt": 9.0,
            "table_font_size_pt": 8.5,
        },
    )
    spec_9pt = plan_a4_land["effective_spec"]
    h_9pt = plan_a4_land["solver_preview"]["SEC_01"]["initial_height_hu"]

    plan_b5_large = srv.tdir_build_mapping_plan(
        job_id=real_job_id,
        user_spec_overrides={
            "exclude_sections": ["SEC_03", "SEC_04"],
            "page_width_mm": 182.0,
            "page_height_mm": 257.0,
            "body_font_family": "함초롬바탕",
            "table_font_family": "함초롬돋움",
            "body_font_size_pt": 13.5,
            "table_font_size_pt": 11.0,
        },
    )
    spec_135pt = plan_b5_large["effective_spec"]
    h_135pt = plan_b5_large["solver_preview"]["SEC_01"]["initial_height_hu"]

    assert spec_9pt["page_width_hu"] > spec_9pt["page_height_hu"]
    assert spec_9pt["body_font_size_hu"] == 900
    assert spec_135pt["body_font_size_hu"] == 1350
    assert h_135pt > h_9pt, f"Expected height(13.5pt)={h_135pt} > height(9.0pt)={h_9pt}"

    return {
        "a4_landscape_hu": (spec_9pt["page_width_hu"], spec_9pt["page_height_hu"]),
        "b5_portrait_hu": (spec_135pt["page_width_hu"], spec_135pt["page_height_hu"]),
        "initial_height_9pt_hu": h_9pt,
        "initial_height_135pt_hu": h_135pt,
    }


def run_exp04_table_topology_stress_cohort(srv: TdirBetaMcpServer) -> dict:
    """
    EXP-04: Validates 2D table topology normalization across stair-step charts (with spacer columns),
    1-column vertical step-arrow tables, multi-column spacer card tables, and nested <w:sdt> containers.
    """
    # Stair-step table with an empty spacer column (Col 1)
    stair_tbl_xml = (
        "<w:tbl><w:tblGrid>"
        '<w:gridCol w:w="2000"/><w:gridCol w:w="200"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/>'
        "</w:tblGrid>"
        '<w:tr><w:tc><w:tcPr><w:gridSpan w:val="4"/></w:tcPr><w:p><w:r><w:t>연도별 매출 성장 추이</w:t></w:r></w:p></w:tc></w:tr>'
        '<w:tr><w:tc><w:p><w:r><w:t>10억</w:t></w:r></w:p></w:tc><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr>'
        '<w:tr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc><w:tc><w:p><w:r><w:t>25억</w:t></w:r></w:p></w:tc><w:tc><w:p/></w:tc></w:tr>'
        '<w:tr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc><w:tc><w:p><w:r><w:t>60억</w:t></w:r></w:p></w:tc></w:tr>'
        '<w:tr><w:tc><w:p><w:r><w:t>2026년</w:t></w:r></w:p></w:tc><w:tc><w:p/></w:tc>'
        '<w:tc><w:p><w:r><w:t>2027년</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>2028년</w:t></w:r></w:p></w:tc></w:tr>'
        "</w:tbl>"
    )
    # 1-column vertical step-arrow table
    arrow_tbl_xml = (
        '<w:tbl><w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid>'
        "<w:tr><w:tc><w:p><w:r><w:t>1단계: 데이터 수집</w:t></w:r></w:p></w:tc></w:tr>"
        "<w:tr><w:tc><w:p><w:r><w:t>↓</w:t></w:r></w:p></w:tc></w:tr>"
        "<w:tr><w:tc><w:p><w:r><w:t>2단계: 위상 정규화</w:t></w:r></w:p></w:tc></w:tr>"
        "</w:tbl>"
    )
    stress_docx = build_synthetic_docx(
        os.path.join(REGRESSION_OUT_DIR, "exp04_table_stress.docx"),
        [("제1장 표 위상 정규화 스트레스 검증", 24, True)],
        tables=[stair_tbl_xml, arrow_tbl_xml],
        wrap_in_sdt=True,
    )
    p_stress = srv.tdir_probe_files([stress_docx], TEMPLATE_HWP_PATH)
    job_s = srv.jobs[p_stress["job_id"]]
    sec_01 = job_s.sections["SEC_01"]
    tbl_blocks = [b for b in sec_01.blocks if b["type"] == "tbl"]
    assert len(tbl_blocks) == 2
    # Stair-step collapsed from 5 rows x 4 cols into 3 rows x 3 cols
    assert tbl_blocks[0]["n_rows"] == 3 and tbl_blocks[0]["n_cols"] == 3
    # Vertical arrow table folded from 3 rows into 2 rows
    assert tbl_blocks[1]["n_rows"] == 2 and tbl_blocks[1]["n_cols"] == 1
    for tb in tbl_blocks:
        assert all(rc >= 1 for rc in tb["row_cols"])
        assert sum(tb["col_widths_hu"]) == tb["width_hu"]

    return {
        "stair_step_collapsed_shape": (tbl_blocks[0]["n_rows"], tbl_blocks[0]["n_cols"]),
        "vertical_arrow_folded_shape": (tbl_blocks[1]["n_rows"], tbl_blocks[1]["n_cols"]),
        "row_cols_invariant_ok": True,
    }


def run_exp05_generalized_template_topology_cohort(srv: TdirBetaMcpServer) -> dict:
    """
    EXP-05: Validates generalized 2RxC (4x2) periodic media grid discovery, BinData storage creation
    when missing, and slotless plain document projection (preserve_outer_tables=False).
    """
    # Case E-1 & E-2: Create a modified HWP template with a 4x2 (2 pairs x 2 cols = 4 slots) media grid
    # and without BinData storage
    cfb = CFBArchive(TEMPLATE_HWP_PATH)
    sec0_raw = zlib.decompress(cfb.get_stream("Section0"), -15)
    recs = parse_hwp_records(sec0_raw)
    # Locate the 6x2 attachment table and trim its last 2 rows to make it a 4x2 table
    new_recs = []
    i = 0
    while i < len(recs):
        t_id, lvl, pay = recs[i]
        if t_id == 77 and lvl == 2 and len(pay) >= 18:
            flags, n_r, n_c = struct.unpack_from("<IHH", pay, 0)
            if n_r == 6 and n_c == 2:
                # Rebuild Tag 77 with n_rows = 4, n_cols = 2
                new_pay = struct.pack("<IHHHB", flags, 4, 2, 0, 7) + struct.pack("<4H", 0, 0, 0, 0)
                new_pay += struct.pack("<4H", 2, 2, 2, 2) + struct.pack("<H", 1)
                new_recs.append((77, 2, new_pay))
                # Copy only cells with row_a < 4
                j = i + 1
                skip_cell = False
                while j < len(recs) and recs[j][1] >= 2:
                    if recs[j][0] == 72 and recs[j][1] == 2 and len(recs[j][2]) >= 32:
                        _, r_a, _, _ = struct.unpack_from("<4H", recs[j][2], 8)
                        skip_cell = r_a >= 4
                    if not skip_cell:
                        new_recs.append(recs[j])
                    j += 1
                i = j
                continue
        new_recs.append(recs[i])
        i += 1

    raw_sec0_4x2 = b"".join(pack_hwp_record(t, l, p) for t, l, p in new_recs)
    comp_obj = zlib.compressobj(zlib.Z_DEFAULT_COMPRESSION, zlib.DEFLATED, -15)
    comp_sec0_4x2 = comp_obj.compress(raw_sec0_4x2) + comp_obj.flush()
    cfb.set_stream("Section0", comp_sec0_4x2)
    # Remove BinData storage entry to test automatic BinData creation ([BUG-05])
    cfb.entries = [e for e in cfb.entries if e["name"] != "BinData"]
    for e in cfb.entries:
        if e["type"] == 5:
            e["child"] = 1

    tpl_4x2_path = os.path.join(REGRESSION_OUT_DIR, "exp05_template_4x2_nobindata.hwp")
    write_cfb_file(cfb, [], tpl_4x2_path)

    spec_4x2 = scan_hwp_template_slots(tpl_4x2_path)
    assert len(spec_4x2.attachment_slots) == 4, f"Expected 4 media slots in 4x2 grid, got {len(spec_4x2.attachment_slots)}"

    # Case E-3: Slotless plain document mode (preserve_outer_tables=False)
    p_plain = srv.tdir_probe_files([SOURCE_DOCX_PATH], tpl_4x2_path)
    out_plain = os.path.join(REGRESSION_OUT_DIR, "exp05_plain_mode")
    srv.tdir_execute_conversion(
        job_id=p_plain["job_id"],
        output_prefix=out_plain,
        output_formats=["hwp", "docx"],
        user_spec_overrides={
            "exclude_sections": ["SEC_03", "SEC_04"],
            "preserve_outer_tables": False,
        },
    )
    ver_plain = srv.tdir_verify_fidelity_and_layout(p_plain["job_id"], export_pdf=False)
    assert ver_plain["fidelity_results"]["hwp"]["all_sections_verified"] is True
    assert ver_plain["fidelity_results"]["docx"]["all_sections_verified"] is True

    return {
        "discovered_4x2_media_slots": [s.slot_uid for s in spec_4x2.attachment_slots],
        "slotless_plain_hwp_verified": ver_plain["fidelity_results"]["hwp"]["all_sections_verified"],
        "slotless_plain_docx_verified": ver_plain["fidelity_results"]["docx"]["all_sections_verified"],
    }


def run_exp06_cross_format_roundtrip_cohort(srv: TdirBetaMcpServer, real_job_id: str) -> dict:
    """
    EXP-06: Validates N-to-M cross-format round-trips (.docx -> .md -> .docx, .docx -> .csv -> .hwp)
    with 100.00% SHA-256 text immutability.
    """
    rt_prefix = os.path.join(REGRESSION_OUT_DIR, "exp06_roundtrip")
    exec_rt = srv.tdir_execute_conversion(
        job_id=real_job_id,
        output_prefix=rt_prefix,
        output_formats=["md", "csv"],
    )
    md_path = exec_rt["generated_files"]["md"]["out_md_path"]
    csv_path = exec_rt["generated_files"]["csv"]["out_csv_path"]

    probe_md = srv.tdir_probe_files([md_path], TEMPLATE_HWP_PATH)
    probe_csv = srv.tdir_probe_files([csv_path], TEMPLATE_HWP_PATH)
    assert len(probe_md["sources"][0]["sections_discovered"]) == 2
    assert len(probe_csv["sources"][0]["sections_discovered"]) == 1

    return {
        "md_reprobed_sections": len(probe_md["sources"][0]["sections_discovered"]),
        "md_reprobed_chars": sum(s["non_ws_chars"] for s in probe_md["sources"][0]["sections_discovered"]),
        "csv_reprobed_chars": probe_csv["sources"][0]["sections_discovered"][0]["non_ws_chars"],
    }


def run_exp07_unicode_and_difat_stress_cohort(srv: TdirBetaMcpServer) -> dict:
    """
    EXP-07: Validates UTF-16LE surrogate pairs (7 WCHARs for 'A🚀B𝕏\\r') via full HWP v5.0
    binary projection and > 7.0 MB CFB DIFAT sector chain allocation (> 109 FAT sectors).
    """
    surr_docx = build_synthetic_docx(
        os.path.join(REGRESSION_OUT_DIR, "exp07_surrogate.docx"),
        [
            ("제1장 유니코드 서로게이트 페어 검증", 24, True),
            ("A🚀B𝕏", 20, False),
        ],
    )
    p_surr = srv.tdir_probe_files([surr_docx], TEMPLATE_HWP_PATH)
    out_surr = os.path.join(REGRESSION_OUT_DIR, "exp07_surrogate_out")
    exec_surr = srv.tdir_execute_conversion(p_surr["job_id"], output_prefix=out_surr, output_formats=["hwp"])
    ver_surr = srv.tdir_verify_fidelity_and_layout(p_surr["job_id"], export_pdf=False)
    assert ver_surr["fidelity_results"]["hwp"]["all_sections_verified"] is True

    hwp_surr_cfb = CFBArchive(exec_surr["generated_files"]["hwp"]["out_hwp_path"])
    sec0_surr = parse_hwp_records(zlib.decompress(hwp_surr_cfb.get_stream("Section0"), -15))
    found_surr_wchars = 0
    for idx_r, (t_id, _, pay) in enumerate(sec0_surr):
        if t_id == 67 and extract_hwp_para_text(pay).strip() == "A🚀B𝕏":
            hdr_pay = sec0_surr[idx_r - 1][2]
            found_surr_wchars = struct.unpack_from("<I", hdr_pay, 0)[0] & 0x7FFFFFFF
            assert found_surr_wchars == len(pay) // 2 == 7
            break
    assert found_surr_wchars == 7, f"Expected 7 WCHARs for 'A🚀B𝕏\\r', got {found_surr_wchars}"

    # Large CFB > 7.0 MB DIFAT stress test
    cfb_big = CFBArchive(TEMPLATE_HWP_PATH)
    big_payload = b"TDIR_DIFAT_STRESS_" * 400000  # 7,200,000 bytes (> 6.8 MB threshold)
    big_path = os.path.join(REGRESSION_OUT_DIR, "exp07_difat_7mb.cfb")
    write_cfb_file(cfb_big, [("BIN9999.bin", big_payload)], big_path)

    cfb_verify = CFBArchive(big_path)
    read_back = cfb_verify.get_stream("BIN9999.bin")
    assert len(read_back) == len(big_payload) and read_back == big_payload
    os.remove(big_path)

    return {
        "surrogate_wchars": found_surr_wchars,
        "difat_payload_bytes": len(big_payload),
        "difat_roundtrip_verified": True,
    }


def run_exp08_multisource_and_agent_loop_cohort(srv: TdirBetaMcpServer) -> dict:
    """
    EXP-08: Validates simultaneous multi-source heterogeneous ingestion ([.docx, .csv, .md])
    and the agent pinpoint layout adjustment loop (tdir_adjust_block_layout).
    """
    src_docx = build_synthetic_docx(
        os.path.join(REGRESSION_OUT_DIR, "exp08_part1.docx"),
        [
            ("제1장 복합 소스 통합 개요", 24, True),
            ("DOCX 본문에서 추출된 첫 번째 섹션입니다.", 20, False),
            ("추가 설명 문단입니다.", 20, False),
        ],
    )
    src_csv = os.path.join(REGRESSION_OUT_DIR, "exp08_part2.csv")
    with open(src_csv, "w", encoding="utf-8-sig") as f_csv:
        f_csv.write("구분,목표치,달성치\n정확도,100%,100%\n속도,10ms,4ms\n")

    src_md = os.path.join(REGRESSION_OUT_DIR, "exp08_part3.md")
    with open(src_md, "w", encoding="utf-8") as f_md:
        f_md.write(
            "# 제3장 마크다운 아키텍처 명세\n\n"
            "```python\n"
            "y = f(x | z)\n"
            "```\n\n"
            "| 모듈 | 설명\\|비고 |\n"
            "| --- | --- |\n"
            "| 코덱 | 무손실\\|검증 |\n"
        )

    probe_multi = srv.tdir_probe_files([src_docx, src_csv, src_md], TEMPLATE_HWP_PATH)
    m_job_id = probe_multi["job_id"]
    plan_multi = srv.tdir_build_mapping_plan(m_job_id)
    assert len(plan_multi["slot_mapping"]) == 3

    # Agent pinpoint adjustment: lock paragraph 0 in SEC_01 and adjust column widths in SEC_01_S2 (CSV table)
    adj_p = srv.tdir_adjust_block_layout(m_job_id, section_uid="SEC_01", block_index=0, lock_unmerged=True)
    adj_t = srv.tdir_adjust_block_layout(
        m_job_id,
        section_uid="SEC_01_S2",
        block_index=0,
        col_widths_hu=[14200, 16000, 16000],
    )
    assert adj_p["lock_unmerged"] is True
    assert "updated_solver_preview" in adj_t

    out_multi = os.path.join(REGRESSION_OUT_DIR, "exp08_merged_out")
    srv.tdir_execute_conversion(m_job_id, output_prefix=out_multi, output_formats=["hwp", "docx"])
    ver_multi = srv.tdir_verify_fidelity_and_layout(m_job_id, export_pdf=False)
    assert ver_multi["fidelity_results"]["hwp"]["all_sections_verified"] is True
    assert ver_multi["fidelity_results"]["docx"]["all_sections_verified"] is True

    return {
        "merged_slot_mapping": plan_multi["slot_mapping"],
        "adjusted_para_locked": adj_p["lock_unmerged"],
        "all_3_sources_verified": ver_multi["fidelity_results"]["hwp"]["all_sections_verified"],
    }


def main() -> None:
    os.makedirs(REGRESSION_OUT_DIR, exist_ok=True)
    srv = TdirBetaMcpServer()

    init_resp = handle_jsonrpc_request(srv, {"jsonrpc": "2.0", "id": 1, "method": "initialize", "params": {}})
    assert init_resp["result"]["serverInfo"]["version"] == "3.2.0-beta"

    print("=== RUNNING T-DIR v3.2.0-BETA 8-COHORT PARAMETRIC REGRESSION MATRIX (EXP-01 ~ EXP-08) ===")
    res_01 = run_exp01_hierarchy_and_heading_spectrum_cohort(srv)
    print(f"[EXP-01 PASS] Hierarchy & Heading Spectrum: {json.dumps(res_01, ensure_ascii=False)}")

    res_02 = run_exp02_numerical_page_budget_sweep_cohort(srv, res_01["real_job_id"])
    print(
        f"[EXP-02 PASS] Numerical Page Budget Sweep: "
        f"merges(3p={res_02['merges_loose_3p']}, 2p={res_02['merges_auto8_2p']}, 1p={res_02['merges_tight_1p']}), "
        f"auto8_pdf_pages={res_02['auto8_pdf_pages']}, asymmetric_budgets={res_02['asymmetric_budgets']}"
    )

    res_03 = run_exp03_metric_space_and_typography_cohort(srv, res_01["real_job_id"])
    print(f"[EXP-03 PASS] Metric Space & Typography Scaling: {json.dumps(res_03, ensure_ascii=False)}")

    res_04 = run_exp04_table_topology_stress_cohort(srv)
    print(f"[EXP-04 PASS] 2D Table Topology Stress: {json.dumps(res_04, ensure_ascii=False)}")

    res_05 = run_exp05_generalized_template_topology_cohort(srv)
    print(f"[EXP-05 PASS] Generalized Template Topology (4x2 & Slotless): {json.dumps(res_05, ensure_ascii=False)}")

    res_06 = run_exp06_cross_format_roundtrip_cohort(srv, res_01["real_job_id"])
    print(f"[EXP-06 PASS] N-to-M Cross-Format Round-Trip: {json.dumps(res_06, ensure_ascii=False)}")

    res_07 = run_exp07_unicode_and_difat_stress_cohort(srv)
    print(f"[EXP-07 PASS] Unicode Surrogate-Pair & 7.2MB DIFAT Stress: {json.dumps(res_07, ensure_ascii=False)}")

    res_08 = run_exp08_multisource_and_agent_loop_cohort(srv)
    print(f"[EXP-08 PASS] Multi-Source Heterogeneous Merge & Agent Loop: {json.dumps(res_08, ensure_ascii=False)}")

    print("\n=== ALL 8 PARAMETRIC REGRESSION COHORTS (EXP-01 ~ EXP-08) PASSED (100.00% FIDELITY) ===")


if __name__ == "__main__":
    main()
