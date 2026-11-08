"""
T-DIR v3.0 Universal Document Intermediate Representation Schema & Text Immutability Verifier.
Enforces 100.00% non-whitespace character preservation across all format conversions.
"""

import hashlib
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Tuple


@dataclass
class TextRunIR:
    text: str
    bold: bool = False
    italic: bool = False
    underline: bool = False
    color_hex: str = "000000"


@dataclass
class ParagraphIR:
    para_uid: str
    jc: str = "both"
    left_tw: int = 0
    right_tw: int = 0
    indent_tw: int = 0
    before_tw: int = 0
    after_tw: int = 10
    runs: List[TextRunIR] = field(default_factory=list)
    r_embed: Optional[str] = None
    structural_role: str = "BODY_PROSE"
    lock_unmerged: bool = False


@dataclass
class TableCellIR:
    row: int
    col: int
    row_span: int = 1
    col_span: int = 1
    width_hu: int = 10000
    height_hu: int = 1000
    mar: Tuple[int, int, int, int] = (140, 140, 35, 35)
    borders: Dict[str, Tuple[int, int, Tuple[int, int, int]]] = field(default_factory=dict)
    fill_hex: Optional[str] = None
    paras: List[ParagraphIR] = field(default_factory=list)
    override_bf_id: Optional[int] = None


@dataclass
class TableGridIR:
    table_uid: str
    n_rows: int
    n_cols: int
    col_widths_hu: List[int]
    row_cols: List[int]
    cells: List[TableCellIR]
    width_hu: int = 46200
    height_hu: int = 2000
    structural_role: str = "DATA_TABLE"


@dataclass
class MediaObjectIR:
    media_uid: str
    rel_id: str
    filename: str
    mime_type: str
    raw_bytes: bytes
    caption_hint: str = ""
    width_hu: int = 22500
    height_hu: int = 12560


@dataclass
class SectionIR:
    section_uid: str
    heading_text: str
    blocks: List[dict] = field(default_factory=list)
    media_objects: List[MediaObjectIR] = field(default_factory=list)
    non_ws_char_count: int = 0
    text_sha256: str = ""


@dataclass
class SlotSpecIR:
    slot_uid: str
    slot_role: str
    discovered_title_text: str
    outer_tbl_record_idx: int
    body_cell_record_idx: int
    caption_cell_record_idx: int = -1
    available_width_hu: int = 46200
    available_height_hu: int = 66921
    max_pages: float = 2.0


@dataclass
class TargetDocumentSpecIR:
    template_path: str
    target_format: str
    page_width_hu: int = 59528
    page_height_hu: int = 84186
    margin_top_hu: int = 4251
    margin_bottom_hu: int = 4252
    margin_left_hu: int = 5669
    margin_right_hu: int = 5669
    body_font_family: str = "맑은 고딕"
    body_font_size_hu: int = 1200
    table_font_size_hu: int = 1000
    default_line_spacing_pct: int = 160
    min_line_spacing_pct: int = 112
    section_slots: List[SlotSpecIR] = field(default_factory=list)
    attachment_slots: List[SlotSpecIR] = field(default_factory=list)


@dataclass
class ConversionJobIR:
    job_id: str
    source_paths: List[str]
    target_path: str
    sections: Dict[str, SectionIR] = field(default_factory=dict)
    target_spec: Optional[TargetDocumentSpecIR] = None
    slot_mapping: Dict[str, str] = field(default_factory=dict)
    output_files: Dict[str, str] = field(default_factory=dict)
    solver_telemetry: Dict[str, dict] = field(default_factory=dict)


def compute_non_ws_sha256(raw_text: str) -> Tuple[str, int, str]:
    normalized = "".join(ch for ch in raw_text if not ch.isspace() and ord(ch) > 32)
    digest = hashlib.sha256(normalized.encode("utf-8")).hexdigest()
    return normalized, len(normalized), digest


def extract_ir_non_ws_text(blocks: List[dict]) -> Tuple[str, int, str]:
    pieces: List[str] = []
    for blk in blocks:
        btype = blk.get("type", "")
        if btype == "p":
            for r in blk.get("runs", []):
                pieces.append(r.get("text", ""))
        elif btype == "tbl":
            for cell in blk.get("cells", []):
                for p in cell.get("paras", []):
                    for r in p.get("runs", []):
                        pieces.append(r.get("text", ""))
    return compute_non_ws_sha256("".join(pieces))


def validate_text_immutability(source_blocks: List[dict], output_text: str) -> dict:
    src_norm, src_len, src_hash = extract_ir_non_ws_text(source_blocks)
    out_norm, out_len, out_hash = compute_non_ws_sha256(output_text)
    is_substring_match = (src_norm in out_norm)
    is_exact_match = (src_hash == out_hash)
    return {
        "source_non_ws_chars": src_len,
        "output_non_ws_chars": out_len,
        "source_sha256": src_hash,
        "output_sha256": out_hash,
        "exact_hash_match": is_exact_match,
        "substring_contained": is_substring_match,
        "lost_chars": max(0, src_len - out_len) if not is_substring_match else 0,
    }
