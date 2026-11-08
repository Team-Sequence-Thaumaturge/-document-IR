"""
T-DIR v3.1 Universal Document Intermediate Representation Schema, Dynamic Spec Resolver,
Mixed-Script Width Estimator, Bipartite Slot Matcher, and Cryptographic Verifier.
Zero hardcoded font names or pixel/HWPUNIT constants: all typography and geometry are
resolved via the priority chain:
  1. UserSpecOverrideIR (Explicit user/MCP specification)
  2. TargetDocumentSpecIR (Dynamically extracted from target template file)
  3. SourceDocumentSpecIR (Dynamically extracted from source document file)
"""

import re
import copy
import hashlib
import unicodedata
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Tuple


def pt_to_hu(pt_val: float) -> int:
    return max(400, min(7200, int(round(pt_val * 100.0))))


def hu_to_pt(hu_val: int) -> float:
    return round(hu_val / 100.0, 2)


def hu_to_half_pt(hu_val: int) -> int:
    return max(4, min(144, int(round(hu_val / 50.0))))


def hu_to_twips(hu_val: int) -> int:
    return int(round(hu_val / 5.0))


def twips_to_hu(tw_val: int) -> int:
    return int(round(tw_val * 5.0))


def mm_to_hu(mm_val: float) -> int:
    return int(round(mm_val * 7200.0 / 25.4))


def hu_to_emu(hu_val: int) -> int:
    return int(round(hu_val * 127.0))


def detect_image_format(raw_bytes: bytes) -> Tuple[str, str]:
    """
    Detects physical image format extension and MIME type from binary magic bytes ([BUG-04]).
    """
    if raw_bytes.startswith(b"\x89PNG\r\n\x1a\n"):
        return "png", "image/png"
    if raw_bytes.startswith(b"\xff\xd8\xff"):
        return "jpg", "image/jpeg"
    if raw_bytes.startswith((b"GIF87a", b"GIF89a")):
        return "gif", "image/gif"
    if raw_bytes.startswith(b"BM"):
        return "bmp", "image/bmp"
    return "png", "image/png"


def compute_weighted_text_units(txt: str) -> float:
    """
    Computes East-Asian-Width weighted character units ([EDGE-04]):
      - Fullwidth / Wide CJK ('F', 'W'): 1.00 unit
      - Whitespace: 0.35 unit
      - Halfwidth / Narrow ASCII alphanumeric & symbols: 0.56 unit
    """
    total = 0.0
    for ch in txt:
        if ch.isspace():
            total += 0.35
        elif unicodedata.east_asian_width(ch) in ("F", "W"):
            total += 1.00
        else:
            total += 0.56
    return total


def extract_ordinal_index(text: str) -> Optional[int]:
    """
    Extracts leading ordinal index from section headings or slot titles
    (e.g. 'Q1.', '1.', '제2장', '[3]', 'Ⅰ.', 'Ⅱ.').
    """
    s = text.strip()
    m_q = re.match(r"^[Qq]\s*(\d+)\b", s)
    if m_q:
        return int(m_q.group(1))
    m_num = re.match(r"^[\[<\(]?\s*(\d+)\s*[\.\]\>\)]", s)
    if m_num:
        return int(m_num.group(1))
    m_ch = re.match(r"^제\s*(\d+)\s*[장절항문]", s)
    if m_ch:
        return int(m_ch.group(1))
    roman_map = {"Ⅰ": 1, "Ⅱ": 2, "Ⅲ": 3, "Ⅳ": 4, "Ⅴ": 5, "I.": 1, "II.": 2, "III.": 3, "IV.": 4}
    for rk, rv in roman_map.items():
        if s.startswith(rk):
            return rv
    return None


def compute_heading_slot_similarity(heading_txt: str, slot_title_txt: str, slot_uid: str = "") -> float:
    """
    Domain-agnostic similarity score between a source section heading and a target slot title ([GEN-02]).
    Combines ordinal index alignment with character 2-gram Jaccard similarity.
    """
    ord_h = extract_ordinal_index(heading_txt)
    ord_s = extract_ordinal_index(slot_title_txt) or extract_ordinal_index(slot_uid)
    if ord_h is not None and ord_s is not None:
        if ord_h != ord_s:
            return 0.0
        base_score = 0.70
    else:
        base_score = 0.0

    norm_h = re.sub(r"[\s\W_]+", "", heading_txt.lower())
    norm_s = re.sub(r"[\s\W_]+", "", slot_title_txt.lower())
    if not norm_h or not norm_s:
        return base_score

    grams_h = {norm_h[i : i + 2] for i in range(max(1, len(norm_h) - 1))}
    grams_s = {norm_s[i : i + 2] for i in range(max(1, len(norm_s) - 1))}
    union_len = len(grams_h | grams_s)
    jaccard = (len(grams_h & grams_s) / float(union_len)) if union_len > 0 else 0.0
    return min(1.0, base_score + 0.30 * jaccard)


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
    width_hu: int = 0
    height_hu: int = 0
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
    width_hu: int = 0
    height_hu: int = 0
    structural_role: str = "DATA_TABLE"


@dataclass
class MediaObjectIR:
    media_uid: str
    rel_id: str
    filename: str
    mime_type: str
    raw_bytes: bytes
    caption_hint: str = ""
    width_hu: int = 0
    height_hu: int = 0


@dataclass
class SectionIR:
    section_uid: str
    heading_text: str
    subtitle_text: str = ""
    blocks: List[dict] = field(default_factory=list)
    media_objects: List[MediaObjectIR] = field(default_factory=list)
    non_ws_char_count: int = 0
    text_sha256: str = ""


@dataclass
class SlotSpecIR:
    slot_uid: str
    slot_role: str
    discovered_title_text: str
    discovered_subtitle_text: str = ""
    outer_tbl_record_idx: int = -1
    body_cell_record_idx: int = -1
    caption_cell_record_idx: int = -1
    outer_width_hu: int = 0
    outer_height_hu: int = 0
    col_widths_hu: List[int] = field(default_factory=list)
    header_height_hu: int = 0
    cell_margins_hu: Tuple[int, int, int, int] = (350, 350, 100, 100)
    available_width_hu: int = 0
    available_height_hu: int = 0
    max_pages: float = 2.0


@dataclass
class SourceDocumentSpecIR:
    source_path: str
    source_format: str
    page_width_hu: Optional[int] = None
    page_height_hu: Optional[int] = None
    margin_top_hu: Optional[int] = None
    margin_bottom_hu: Optional[int] = None
    margin_left_hu: Optional[int] = None
    margin_right_hu: Optional[int] = None
    body_font_family: Optional[str] = None
    table_font_family: Optional[str] = None
    body_font_size_hu: Optional[int] = None
    table_font_size_hu: Optional[int] = None
    default_line_spacing_pct: Optional[int] = None


@dataclass
class TargetDocumentSpecIR:
    template_path: str
    target_format: str
    page_width_hu: int = 0
    page_height_hu: int = 0
    margin_top_hu: int = 0
    margin_bottom_hu: int = 0
    margin_left_hu: int = 0
    margin_right_hu: int = 0
    margin_header_hu: int = 0
    margin_footer_hu: int = 0
    margins_explicitly_overridden: bool = False
    body_font_family: str = ""
    table_font_family: str = ""
    body_font_size_hu: int = 0
    table_font_size_hu: int = 0
    default_line_spacing_pct: int = 0
    min_line_spacing_pct: int = 122
    table_default_line_spacing_pct: int = 125
    table_min_line_spacing_pct: int = 112
    preamble_blocks: List[dict] = field(default_factory=list)
    attachment_header_paras: List[dict] = field(default_factory=list)
    attachment_placeholder_img_text: str = ""
    attachment_placeholder_cap_text: str = ""
    section_slots: List[SlotSpecIR] = field(default_factory=list)
    attachment_slots: List[SlotSpecIR] = field(default_factory=list)


@dataclass
class UserSpecOverrideIR:
    body_font_family: Optional[str] = None
    table_font_family: Optional[str] = None
    body_font_size_pt: Optional[float] = None
    table_font_size_pt: Optional[float] = None
    body_line_spacing_min_pct: Optional[int] = None
    body_line_spacing_max_pct: Optional[int] = None
    table_line_spacing_min_pct: Optional[int] = None
    table_line_spacing_max_pct: Optional[int] = None
    page_width_mm: Optional[float] = None
    page_height_mm: Optional[float] = None
    margins_mm: Optional[Dict[str, float]] = None
    total_page_limit: Optional[int] = None
    section_page_budgets: Dict[str, float] = field(default_factory=dict)
    page_body_height_hu: Optional[int] = None
    include_sections: Optional[List[str]] = None
    exclude_sections: List[str] = field(default_factory=list)
    preserve_outer_tables: bool = True
    route_media_to_attachment_slots: bool = True
    max_attachment_images: Optional[int] = None
    attachment_captions: List[str] = field(default_factory=list)
    allow_font_size_elastic_scaling: bool = False
    clear_unmapped_slots: bool = False
    style_inheritance_mode: str = "TARGET_STRICT"

    @classmethod
    def from_dict(cls, d: Optional[dict]) -> "UserSpecOverrideIR":
        if not d:
            return cls()
        sec_budgets = dict(d.get("section_page_budgets") or d.get("page_budgets") or {})
        return cls(
            body_font_family=d.get("body_font_family") or d.get("font_family"),
            table_font_family=d.get("table_font_family") or d.get("font_family"),
            body_font_size_pt=float(d["body_font_size_pt"]) if d.get("body_font_size_pt") is not None else None,
            table_font_size_pt=float(d["table_font_size_pt"]) if d.get("table_font_size_pt") is not None else None,
            body_line_spacing_min_pct=(
                int(d["body_line_spacing_min_pct"]) if d.get("body_line_spacing_min_pct") is not None else None
            ),
            body_line_spacing_max_pct=(
                int(d["body_line_spacing_max_pct"]) if d.get("body_line_spacing_max_pct") is not None else None
            ),
            table_line_spacing_min_pct=(
                int(d["table_line_spacing_min_pct"]) if d.get("table_line_spacing_min_pct") is not None else None
            ),
            table_line_spacing_max_pct=(
                int(d["table_line_spacing_max_pct"]) if d.get("table_line_spacing_max_pct") is not None else None
            ),
            page_width_mm=float(d["page_width_mm"]) if d.get("page_width_mm") is not None else None,
            page_height_mm=float(d["page_height_mm"]) if d.get("page_height_mm") is not None else None,
            margins_mm=d.get("margins_mm"),
            total_page_limit=int(d["total_page_limit"]) if d.get("total_page_limit") is not None else None,
            section_page_budgets={k: float(v) for k, v in sec_budgets.items()},
            page_body_height_hu=int(d["page_body_height_hu"]) if d.get("page_body_height_hu") is not None else None,
            include_sections=list(d["include_sections"]) if d.get("include_sections") is not None else None,
            exclude_sections=list(d.get("exclude_sections") or []),
            preserve_outer_tables=bool(d.get("preserve_outer_tables", True)),
            route_media_to_attachment_slots=bool(d.get("route_media_to_attachment_slots", True)),
            max_attachment_images=(
                int(d["max_attachment_images"]) if d.get("max_attachment_images") is not None else None
            ),
            attachment_captions=list(d.get("attachment_captions") or []),
            allow_font_size_elastic_scaling=bool(d.get("allow_font_size_elastic_scaling", False)),
            clear_unmapped_slots=bool(d.get("clear_unmapped_slots", False)),
            style_inheritance_mode=str(d.get("style_inheritance_mode", "TARGET_STRICT")),
        )


def resolve_effective_document_spec(
    source_specs: List[SourceDocumentSpecIR],
    target_spec: TargetDocumentSpecIR,
    user_overrides: UserSpecOverrideIR,
) -> TargetDocumentSpecIR:
    """
    Resolves the final effective TargetDocumentSpecIR using the strict priority chain:
      Priority 1: UserSpecOverrideIR (Explicit user/MCP specification)
      Priority 2: TargetDocumentSpecIR (Dynamically extracted from target template file)
      Priority 3: SourceDocumentSpecIR (Dynamically extracted from source document file)
      Priority 4: Neutral geometric fallback (only if neither user, target, nor source specifies it)
    """
    eff = copy.deepcopy(target_spec)
    src = source_specs[0] if source_specs else None

    # 1. Resolve Font Families
    if user_overrides.body_font_family:
        eff.body_font_family = user_overrides.body_font_family
    elif not eff.body_font_family and src and src.body_font_family:
        eff.body_font_family = src.body_font_family
    elif not eff.body_font_family:
        eff.body_font_family = "sans-serif"

    if user_overrides.table_font_family:
        eff.table_font_family = user_overrides.table_font_family
    elif not eff.table_font_family:
        eff.table_font_family = eff.body_font_family

    # 2. Resolve Body & Table Font Sizes (HU)
    if user_overrides.body_font_size_pt is not None:
        eff.body_font_size_hu = pt_to_hu(user_overrides.body_font_size_pt)
    elif eff.body_font_size_hu <= 0 and src and src.body_font_size_hu:
        eff.body_font_size_hu = src.body_font_size_hu
    elif eff.body_font_size_hu <= 0:
        eff.body_font_size_hu = 1000

    if user_overrides.table_font_size_pt is not None:
        eff.table_font_size_hu = pt_to_hu(user_overrides.table_font_size_pt)
    elif eff.table_font_size_hu <= 0 and src and src.table_font_size_hu:
        eff.table_font_size_hu = src.table_font_size_hu
    elif eff.table_font_size_hu <= 0:
        eff.table_font_size_hu = max(800, int(round(eff.body_font_size_hu * 0.85)))

    # 3. Resolve Line Spacing Bounds (%)
    if user_overrides.body_line_spacing_max_pct is not None:
        eff.default_line_spacing_pct = max(100, min(300, user_overrides.body_line_spacing_max_pct))
    elif eff.default_line_spacing_pct <= 0 and src and src.default_line_spacing_pct:
        eff.default_line_spacing_pct = src.default_line_spacing_pct
    elif eff.default_line_spacing_pct <= 0:
        eff.default_line_spacing_pct = 160

    if user_overrides.body_line_spacing_min_pct is not None:
        eff.min_line_spacing_pct = max(100, min(eff.default_line_spacing_pct, user_overrides.body_line_spacing_min_pct))

    if user_overrides.table_line_spacing_max_pct is not None:
        eff.table_default_line_spacing_pct = max(100, min(300, user_overrides.table_line_spacing_max_pct))

    if user_overrides.table_line_spacing_min_pct is not None:
        eff.table_min_line_spacing_pct = max(
            100, min(eff.table_default_line_spacing_pct, user_overrides.table_line_spacing_min_pct)
        )

    # 4. Resolve Page Dimensions & Margins (HU)
    if user_overrides.page_width_mm is not None:
        eff.page_width_hu = mm_to_hu(user_overrides.page_width_mm)
    elif eff.page_width_hu <= 0 and src and src.page_width_hu:
        eff.page_width_hu = src.page_width_hu
    elif eff.page_width_hu <= 0:
        eff.page_width_hu = mm_to_hu(210.0)

    if user_overrides.page_height_mm is not None:
        eff.page_height_hu = mm_to_hu(user_overrides.page_height_mm)
    elif eff.page_height_hu <= 0 and src and src.page_height_hu:
        eff.page_height_hu = src.page_height_hu
    elif eff.page_height_hu <= 0:
        eff.page_height_hu = mm_to_hu(297.0)

    if user_overrides.margins_mm:
        eff.margins_explicitly_overridden = True
        if "top" in user_overrides.margins_mm:
            eff.margin_top_hu = mm_to_hu(user_overrides.margins_mm["top"])
        if "bottom" in user_overrides.margins_mm:
            eff.margin_bottom_hu = mm_to_hu(user_overrides.margins_mm["bottom"])
        if "left" in user_overrides.margins_mm:
            eff.margin_left_hu = mm_to_hu(user_overrides.margins_mm["left"])
        if "right" in user_overrides.margins_mm:
            eff.margin_right_hu = mm_to_hu(user_overrides.margins_mm["right"])

    # 5. Apply Section Page Budgets to Discovered Slots
    for sl in eff.section_slots:
        for sec_k, p_val in user_overrides.section_page_budgets.items():
            if sl.slot_uid == sec_k or sl.slot_uid.startswith(f"{sec_k}_"):
                sl.max_pages = float(p_val)

    return eff


@dataclass
class ConversionJobIR:
    job_id: str
    source_paths: List[str]
    target_path: str
    sections: Dict[str, SectionIR] = field(default_factory=dict)
    source_specs: List[SourceDocumentSpecIR] = field(default_factory=list)
    raw_target_spec: Optional[TargetDocumentSpecIR] = None
    target_spec: Optional[TargetDocumentSpecIR] = None
    user_overrides: UserSpecOverrideIR = field(default_factory=UserSpecOverrideIR)
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
    is_substring_match = src_norm in out_norm
    is_exact_match = src_hash == out_hash
    return {
        "source_non_ws_chars": src_len,
        "output_non_ws_chars": out_len,
        "source_sha256": src_hash,
        "output_sha256": out_hash,
        "exact_hash_match": is_exact_match,
        "substring_contained": is_substring_match,
        "lost_chars": max(0, src_len - out_len) if not is_substring_match else 0,
    }
