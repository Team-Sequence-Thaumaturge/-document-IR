"""
T-DIR v3.2 Universal Document Intermediate Representation Schema, Dynamic Spec Resolver,
Unsupervised Heading Spectrum Detector, Universal Selector Matcher, Mixed-Script Width Estimator,
Bipartite Slot Matcher, and Cryptographic Verifier.
Zero hardcoded font names, pixel/HWPUNIT constants, or document-specific section/slot IDs:
  1. UserSpecOverrideIR (Explicit user/MCP specification)
  2. TargetDocumentSpecIR (Dynamically extracted from target template file)
  3. SourceDocumentSpecIR (Dynamically extracted from source document file)
"""

import os
import re
import copy
import hashlib
import zipfile
import unicodedata
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Tuple, Union

M_OMML_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
ET.register_namespace("m", M_OMML_NS)

_WINDOWS_RESERVED_DEVICES = {
    "CON",
    "PRN",
    "AUX",
    "NUL",
    "COM1",
    "COM2",
    "COM3",
    "COM4",
    "COM5",
    "COM6",
    "COM7",
    "COM8",
    "COM9",
    "LPT1",
    "LPT2",
    "LPT3",
    "LPT4",
    "LPT5",
    "LPT6",
    "LPT7",
    "LPT8",
    "LPT9",
}


def strip_illegal_xml_control_chars(text: str) -> str:
    """
    Strips XML 1.0 illegal control characters (0x00..0x08, 0x0B, 0x0C, 0x0E..0x1F, 0xFFFE, 0xFFFF)
    while preserving 100% of valid whitespace (\t, \n, \r) and printable Unicode characters ([SEC-05]).
    """
    if not text:
        return ""
    return "".join(
        ch
        for ch in text
        if (ord(ch) >= 0x20 and ord(ch) not in (0xFFFE, 0xFFFF)) or ch in ("\t", "\n", "\r")
    )


def validate_safe_filesystem_path(path_str: str, must_exist: bool = False) -> str:
    """
    Validates a local filesystem path against null-byte injection and Windows reserved device names ([SEC-04]).
    """
    if not path_str or "\x00" in path_str:
        raise ValueError("Invalid filesystem path: empty or contains null byte.")
    norm_p = os.path.abspath(path_str)
    base_stem = os.path.splitext(os.path.basename(norm_p))[0].upper()
    if base_stem in _WINDOWS_RESERVED_DEVICES:
        raise ValueError(f"Rejected reserved device filename: {base_stem}")
    if must_exist and not os.path.exists(norm_p):
        raise FileNotFoundError(f"Required file does not exist: {norm_p}")
    return norm_p


def sanitize_archive_rel_path(rel_target: str) -> str:
    """
    Sanitizes an internal ZIP archive relationship path and rejects path traversal ('..') or drive prefixes ([SEC-02]).
    """
    cleaned = rel_target.replace("\\", "/").lstrip("/")
    parts = [p for p in cleaned.split("/") if p and p != "."]
    if any(p == ".." or ":" in p for p in parts):
        raise ValueError(f"Blocked archive path traversal in relationship target: {rel_target}")
    return "/".join(parts)


def validate_safe_zip_archive(
    zf_or_path: Union[zipfile.ZipFile, str],
    max_entries: int = 4096,
    max_single_bytes: int = 67108864,
    max_total_bytes: int = 268435456,
) -> None:
    """
    Validates a ZIP archive (.docx, .hwpx, .xlsx) against Zip-Slip path traversal and ZIP decompression bombs ([SEC-02]).
    """
    if isinstance(zf_or_path, str):
        with zipfile.ZipFile(zf_or_path, "r") as zf_opened:
            validate_safe_zip_archive(
                zf_opened,
                max_entries=max_entries,
                max_single_bytes=max_single_bytes,
                max_total_bytes=max_total_bytes,
            )
        return
    infos = zf_or_path.infolist()
    if len(infos) > max_entries:
        raise ValueError(f"ZIP archive entry count ({len(infos)}) exceeds safety limit ({max_entries}).")
    total_uncompressed = 0
    for info in infos:
        fname = info.filename.replace("\\", "/")
        parts = [p for p in fname.split("/") if p]
        if fname.startswith("/") or any(p == ".." or ":" in p for p in parts):
            raise ValueError(f"Blocked Zip-Slip path traversal entry: {info.filename}")
        if info.file_size > max_single_bytes:
            raise ValueError(f"ZIP entry '{info.filename}' uncompressed size ({info.file_size}) exceeds limit.")
        total_uncompressed += info.file_size
        if total_uncompressed > max_total_bytes:
            raise ValueError(f"ZIP archive total uncompressed size ({total_uncompressed}) exceeds limit.")


def safe_parse_xml_bytes(xml_bytes: bytes, max_xml_bytes: int = 67108864) -> ET.Element:
    """
    Parses XML bytes after enforcing size limits and rejecting DOCTYPE / ENTITY declarations
    to prevent XML External Entity (XXE) and Billion Laughs entity expansion attacks ([SEC-01]).
    """
    if len(xml_bytes) > max_xml_bytes:
        raise ValueError(f"XML payload size ({len(xml_bytes)} B) exceeds safety limit ({max_xml_bytes} B).")
    head_sample = xml_bytes[:16384].upper()
    if b"<!DOCTYPE" in head_sample or b"<!ENTITY" in head_sample:
        raise ValueError("Rejected XML containing DOCTYPE or ENTITY declarations (XXE / Entity Bomb guard).")
    return ET.fromstring(xml_bytes)


def sanitize_spreadsheet_cell_formula(text: str, neutralize_dde: bool = True) -> str:
    """
    Guards against DDE / command execution payloads in spreadsheet cells while preserving normal numeric/text cells ([SEC-05]).
    """
    cleaned = strip_illegal_xml_control_chars(text)
    if not neutralize_dde or not cleaned:
        return cleaned
    stripped = cleaned.lstrip()
    if stripped.startswith(("=", "+", "-", "@")) and ("|" in stripped or "CMD" in stripped.upper() or "MSEXCEL" in stripped.upper()):
        return "'" + cleaned
    return cleaned


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
    Extracts leading ordinal index from section headings, slot titles, or selector strings
    (e.g. 'SEC_01', 'SLOT_02', 'Q1.', '1.', '제2장', '[3]', 'Ⅰ.', 'II.', 'Chapter 3').
    Avoids matching sub-ordinals like '1-1.' or '1.2'.
    """
    s = text.strip()
    if not s or re.match(r"^\d+[\-\.]\d+", s):
        return None
    m_uid = re.match(r"^(?:SEC|SLOT|ATTACH|MEDIA)[_\-\s]*0*(\d+)\b", s, re.IGNORECASE)
    if m_uid:
        return int(m_uid.group(1))
    m_q = re.match(r"^[Qq]\s*0*(\d+)\b", s)
    if m_q:
        return int(m_q.group(1))
    m_ch_en = re.match(r"^(?:Chapter|Section|Part)\s+0*(\d+)\b", s, re.IGNORECASE)
    if m_ch_en:
        return int(m_ch_en.group(1))
    m_num = re.match(r"^[\[<\(]?\s*0*(\d+)\s*[\.\]\>\)]", s)
    if m_num:
        return int(m_num.group(1))
    m_ch = re.match(r"^제\s*0*(\d+)\s*[장절항문편부]", s)
    if m_ch:
        return int(m_ch.group(1))
    if s.isdigit():
        return int(s)
    roman_unicode = {
        "Ⅰ": 1,
        "Ⅱ": 2,
        "Ⅲ": 3,
        "Ⅳ": 4,
        "Ⅴ": 5,
        "Ⅵ": 6,
        "Ⅶ": 7,
        "Ⅷ": 8,
        "Ⅸ": 9,
        "Ⅹ": 10,
    }
    if s[0] in roman_unicode:
        return roman_unicode[s[0]]
    m_rom = re.match(r"^(X|IX|VIII|VII|VI|V|IV|III|II|I)[\.\)\s]", s)
    if m_rom:
        rom_ascii = {
            "I": 1,
            "II": 2,
            "III": 3,
            "IV": 4,
            "V": 5,
            "VI": 6,
            "VII": 7,
            "VIII": 8,
            "IX": 9,
            "X": 10,
        }
        return rom_ascii.get(m_rom.group(1))
    return None


def is_top_level_section_heading(
    ptxt: str,
    p_max_sz_hu: int,
    p_is_bold: bool,
    p_outline_lvl: Optional[int],
    mode_body_sz_hu: int,
    is_in_table: bool = False,
) -> Tuple[bool, str]:
    """
    Unsupervised typographic and topological top-level section heading detector ([NUM-01]).
    Detects top-level sections across arbitrary document domains without hardcoded title strings.
    Returns (is_heading, section_role) where section_role in {'BODY', 'REFERENCE'}.
    """
    s = ptxt.strip()
    if is_in_table or not s or len(s) > 95:
        return False, "BODY"

    if re.match(
        r"^\[?\s*(?:출처|참고\s*문헌|References?|Bibliography)(?:\s*[\(\[（].*[\)\]）]|\s*\]?)$",
        s,
        re.IGNORECASE,
    ):
        return True, "REFERENCE"

    bullet_prefixes = (
        "■",
        "□",
        "●",
        "○",
        "•",
        "◦",
        "◆",
        "◇",
        "-",
        "–",
        "—",
        "※",
        "*",
        "▸",
        "▹",
        "▶",
        "▷",
        "(",
        "①",
        "②",
        "③",
        "④",
        "⑤",
        "가.",
        "나.",
        "다.",
        "라.",
    )
    if s.startswith(bullet_prefixes) or re.match(r"^\d+[\-\.]\d+", s):
        return False, "BODY"

    if p_outline_lvl is not None:
        if p_outline_lvl == 0:
            return True, "BODY"
        return False, "BODY"

    # Explicit structural chapter/section prefixes (e.g. '제1장', 'Ⅰ.', 'I.', 'Chapter 1', 'Q1.')
    if (
        (
            re.match(r"^[Qq]\s*\d+[\.\s]", s)
            or re.match(r"^제\s*\d+\s*[장절편부]", s)
            or re.match(r"^(?:Chapter|Section|Part)\s+\d+\b", s, re.IGNORECASE)
            or re.match(r"^[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ][\.\)\s]", s)
            or re.match(r"^(?:X|IX|VIII|VII|VI|V|IV|III|II|I)\.\s+\S+", s)
        )
        and len(s) <= 85
        and not s.endswith(("다.", "입니다.", "함.", "음."))
    ):
        return True, "BODY"

    # Typographic size jump relative to the document's mode body font size (S_mode)
    eff_mode_hu = mode_body_sz_hu if mode_body_sz_hu > 0 else 1000
    if (
        p_max_sz_hu >= eff_mode_hu + 150
        and (p_is_bold or p_max_sz_hu >= eff_mode_hu + 200)
        and len(s) <= 75
        and not s.endswith((".", "다", "함", "음", ",", ";"))
    ):
        return True, "BODY"

    return False, "BODY"


def is_subtitle_or_instruction_line(ptxt: str, p_max_sz_hu: int, mode_body_sz_hu: int) -> bool:
    """
    Domain-agnostic detector for a single subtitle or instruction line immediately following
    a top-level section heading (replacing hardcoded strings like '아이디어의 개발').
    """
    s = ptxt.strip()
    if not s or len(s) > 110:
        return False
    if s.startswith(("■", "□", "●", "○", "•", "-", "1.", "2.", "3.", "Q", "제1", "Ⅰ", "I.")):
        return False
    if s.startswith(("※", "*", "▸", "▹")):
        return True
    eff_mode_hu = mode_body_sz_hu if mode_body_sz_hu > 0 else 1000
    if (
        len(s) <= 85
        and "(" in s
        and ")" in s
        and p_max_sz_hu <= eff_mode_hu + 50
        and not s.endswith((".", "다", "함", "음"))
    ):
        return True
    return False


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
    equation_ast: Optional[dict] = None


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
    ordinal_index: Optional[int] = None
    section_role: str = "BODY"


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
    ordinal_index: Optional[int] = None
    grid_cols: int = 2


def matches_section_selector(
    selector: str,
    sec: SectionIR,
    matched_slot_uid: Optional[str] = None,
) -> bool:
    """
    Universal selector matcher supporting generic UIDs ('SEC_01', 'SLOT_01'),
    structural roles ('PREAMBLE', 'REFERENCES'), ordinal aliases ('Q1', '1', '제1장', 'I.'),
    and heading substrings ([NUM-02]).
    """
    sel = selector.strip()
    if not sel:
        return False
    sel_up = sel.upper()
    uid_up = sec.section_uid.upper()
    base_uid_up = uid_up.split("_S")[0]
    if sel_up in (uid_up, base_uid_up):
        return True
    if matched_slot_uid and sel_up == matched_slot_uid.upper():
        return True
    if sel_up == "PREAMBLE" and sec.section_role == "PREAMBLE":
        return True
    if sel_up in ("REFERENCES", "REFERENCE") and sec.section_role == "REFERENCE":
        return True
    # Only match bare ordinal aliases against primary (non-secondary _S2) body sections
    if "_S" not in uid_up and sec.section_role == "BODY":
        sel_ord = extract_ordinal_index(sel)
        sec_ord = sec.ordinal_index or extract_ordinal_index(sec.heading_text) or extract_ordinal_index(sec.section_uid)
        if sel_ord is not None and sec_ord is not None and sel_ord == sec_ord:
            return True
    if len(sel) >= 2 and sel.lower() in sec.heading_text.lower():
        return True
    return False


def resolve_budget_for_section_or_slot(
    budgets_map: Dict[str, float],
    sec: Optional[SectionIR],
    slot: Optional[SlotSpecIR],
) -> Optional[float]:
    """
    Resolves a numerical page budget for a section or slot from `section_page_budgets`
    whether keyed by 'SEC_01', 'SLOT_01', 'Q1', '1', or title substring ([NUM-04]).
    """
    if not budgets_map:
        return None
    for k, val in budgets_map.items():
        if sec is not None and matches_section_selector(k, sec, slot.slot_uid if slot else None):
            return float(val)
        if slot is not None:
            k_up = k.strip().upper()
            sl_up = slot.slot_uid.upper()
            if k_up == sl_up or sl_up.startswith(f"{k_up}_"):
                return float(val)
            k_ord = extract_ordinal_index(k)
            sl_ord = slot.ordinal_index or extract_ordinal_index(slot.discovered_title_text) or extract_ordinal_index(slot.slot_uid)
            if k_ord is not None and sl_ord is not None and k_ord == sl_ord:
                return float(val)
    return None


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

    # 5. Apply Section Page Budgets to Discovered Slots via Universal Selector Resolution
    for sl in eff.section_slots:
        resolved_b = resolve_budget_for_section_or_slot(user_overrides.section_page_budgets, None, sl)
        if resolved_b is not None:
            sl.max_pages = float(resolved_b)

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
    page_budget_solution: Dict[str, object] = field(default_factory=dict)


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


# ==============================================================================
# 5. 3-Way Isomorphic Equation AST Compiler ([P2-02]: OMML <-> HWP EQN <-> LaTeX)
# ==============================================================================

NS_OMML = "http://schemas.openxmlformats.org/officeDocument/2006/math"

_LATEX_TO_UNICODE_SYM: Dict[str, str] = {
    r"\alpha": "α",
    r"\beta": "β",
    r"\gamma": "γ",
    r"\delta": "δ",
    r"\epsilon": "ϵ",
    r"\theta": "θ",
    r"\lambda": "λ",
    r"\mu": "μ",
    r"\pi": "π",
    r"\rho": "ρ",
    r"\sigma": "σ",
    r"\tau": "τ",
    r"\phi": "ϕ",
    r"\omega": "ω",
    r"\Delta": "Δ",
    r"\Sigma": "Σ",
    r"\Omega": "Ω",
    r"\pm": "±",
    r"\times": "×",
    r"\div": "÷",
    r"\cdot": "·",
    r"\le": "≤",
    r"\ge": "≥",
    r"\ne": "≠",
    r"\approx": "≈",
    r"\to": "→",
    r"\infty": "∞",
    r"\sum": "∑",
    r"\prod": "∏",
    r"\int": "∫",
    r"\partial": "∂",
}

_UNICODE_TO_LATEX_SYM: Dict[str, str] = {v: k for k, v in _LATEX_TO_UNICODE_SYM.items()}

_HWPEQN_TO_UNICODE_SYM: Dict[str, str] = {
    "alpha": "α",
    "beta": "β",
    "gamma": "γ",
    "delta": "δ",
    "epsilon": "ϵ",
    "theta": "θ",
    "lambda": "λ",
    "mu": "μ",
    "pi": "π",
    "rho": "ρ",
    "sigma": "σ",
    "tau": "τ",
    "phi": "ϕ",
    "omega": "ω",
    "DELTA": "Δ",
    "SIGMA": "Σ",
    "OMEGA": "Ω",
    "+-": "±",
    "times": "×",
    "div": "÷",
    "cdot": "·",
    "<=": "≤",
    ">=": "≥",
    "!=": "≠",
    "approx": "≈",
    "->": "→",
    "inf": "∞",
    "sum": "∑",
    "prod": "∏",
    "int": "∫",
    "partial": "∂",
}

_UNICODE_TO_HWPEQN_SYM: Dict[str, str] = {v: k for k, v in _HWPEQN_TO_UNICODE_SYM.items()}


def normalize_equation_ast(node: dict) -> dict:
    if not isinstance(node, dict):
        return {"kind": "sym", "val": str(node)}
    kind = node.get("kind", "sym")
    if kind == "sym":
        return {"kind": "sym", "val": str(node.get("val", "")).strip()}
    if kind == "seq":
        flat: List[dict] = []
        for item in node.get("items", []):
            norm_item = normalize_equation_ast(item)
            if norm_item.get("kind") == "seq":
                flat.extend(norm_item.get("items", []))
            elif norm_item.get("kind") == "sym" and not norm_item.get("val"):
                continue
            else:
                flat.append(norm_item)
        if len(flat) == 1:
            return flat[0]
        return {"kind": "seq", "items": flat}
    if kind == "frac":
        return {
            "kind": "frac",
            "num": normalize_equation_ast(node.get("num", {"kind": "sym", "val": "1"})),
            "den": normalize_equation_ast(node.get("den", {"kind": "sym", "val": "1"})),
        }
    if kind == "sqrt":
        return {
            "kind": "sqrt",
            "body": normalize_equation_ast(node.get("body", {"kind": "sym", "val": "x"})),
        }
    if kind == "sup":
        return {
            "kind": "sup",
            "base": normalize_equation_ast(node.get("base", {"kind": "sym", "val": "x"})),
            "exp": normalize_equation_ast(node.get("exp", {"kind": "sym", "val": "2"})),
        }
    if kind == "sub":
        return {
            "kind": "sub",
            "base": normalize_equation_ast(node.get("base", {"kind": "sym", "val": "x"})),
            "sub": normalize_equation_ast(node.get("sub", {"kind": "sym", "val": "i"})),
        }
    if kind == "subsup":
        return {
            "kind": "subsup",
            "base": normalize_equation_ast(node.get("base", {"kind": "sym", "val": "x"})),
            "sub": normalize_equation_ast(node.get("sub", {"kind": "sym", "val": "i"})),
            "exp": normalize_equation_ast(node.get("exp", {"kind": "sym", "val": "2"})),
        }
    return {"kind": "sym", "val": str(node.get("val", ""))}


def _extract_braced_group(text: str, pos: int) -> Tuple[str, int]:
    while pos < len(text) and text[pos].isspace():
        pos += 1
    if pos >= len(text):
        return "", pos
    if text[pos] != "{":
        return text[pos], pos + 1
    depth = 0
    start = pos + 1
    i = pos
    while i < len(text):
        ch = text[i]
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                return text[start:i], i + 1
        i += 1
    return text[start:], len(text)


def parse_latex_to_equation_ast(latex_str: str) -> dict:
    s = latex_str.strip()
    if s.startswith("$$") and s.endswith("$$") and len(s) >= 4:
        s = s[2:-2].strip()
    elif s.startswith("$") and s.endswith("$") and len(s) >= 2:
        s = s[1:-1].strip()

    items: List[dict] = []
    i = 0
    n = len(s)
    while i < n:
        if s[i].isspace():
            i += 1
            continue
        if s.startswith(r"\frac", i):
            i += 5
            num_str, i = _extract_braced_group(s, i)
            den_str, i = _extract_braced_group(s, i)
            items.append({
                "kind": "frac",
                "num": parse_latex_to_equation_ast(num_str),
                "den": parse_latex_to_equation_ast(den_str),
            })
            continue
        if s.startswith(r"\sqrt", i):
            i += 5
            body_str, i = _extract_braced_group(s, i)
            items.append({
                "kind": "sqrt",
                "body": parse_latex_to_equation_ast(body_str),
            })
            continue
        if s[i] == "{":
            grp_str, i = _extract_braced_group(s, i)
            items.append(parse_latex_to_equation_ast(grp_str))
            continue
        if s[i] in ("_", "^") and items:
            base_node = items.pop()
            sub_node: Optional[dict] = None
            exp_node: Optional[dict] = None
            for _ in range(2):
                while i < n and s[i].isspace():
                    i += 1
                if i < n and s[i] == "_":
                    i += 1
                    sub_str, i = _extract_braced_group(s, i)
                    sub_node = parse_latex_to_equation_ast(sub_str)
                elif i < n and s[i] == "^":
                    i += 1
                    exp_str, i = _extract_braced_group(s, i)
                    exp_node = parse_latex_to_equation_ast(exp_str)
                else:
                    break
            if sub_node is not None and exp_node is not None:
                items.append({"kind": "subsup", "base": base_node, "sub": sub_node, "exp": exp_node})
            elif sub_node is not None:
                items.append({"kind": "sub", "base": base_node, "sub": sub_node})
            elif exp_node is not None:
                items.append({"kind": "sup", "base": base_node, "exp": exp_node})
            continue
        if s[i] == "\\":
            j = i + 1
            while j < n and s[j].isalpha():
                j += 1
            cmd = s[i:j]
            sym_val = _LATEX_TO_UNICODE_SYM.get(cmd, cmd.lstrip("\\"))
            items.append({"kind": "sym", "val": sym_val})
            i = j
            continue
        if s[i].isdigit() or s[i].isalpha():
            j = i + 1
            while j < n and (s[j].isdigit() or (s[i].isdigit() and s[j] == ".")):
                j += 1
            items.append({"kind": "sym", "val": s[i:j]})
            i = j
            continue
        items.append({"kind": "sym", "val": s[i]})
        i += 1

    return normalize_equation_ast({"kind": "seq", "items": items})


def _tokenize_hwpeqn(eqn_str: str) -> List[Union[str, Tuple[str, str]]]:
    tokens: List[Union[str, Tuple[str, str]]] = []
    i = 0
    n = len(eqn_str)
    while i < n:
        if eqn_str[i].isspace() or eqn_str[i] == "~":
            i += 1
            continue
        if eqn_str[i] == "{":
            grp_str, i = _extract_braced_group(eqn_str, i)
            tokens.append(("GROUP", grp_str))
            continue
        for two_op in ("+-", "<=", ">=", "!=", "->"):
            if eqn_str.startswith(two_op, i):
                tokens.append(two_op)
                i += len(two_op)
                break
        else:
            if eqn_str[i].isalpha():
                j = i + 1
                while j < n and eqn_str[j].isalpha():
                    j += 1
                tokens.append(eqn_str[i:j])
                i = j
            elif eqn_str[i].isdigit():
                j = i + 1
                while j < n and (eqn_str[j].isdigit() or eqn_str[j] == "."):
                    j += 1
                tokens.append(eqn_str[i:j])
                i = j
            else:
                tokens.append(eqn_str[i])
                i += 1
    return tokens


def parse_hwpeqn_to_equation_ast(eqn_str: str) -> dict:
    raw_tokens = _tokenize_hwpeqn(eqn_str.strip())
    items: List[dict] = []
    idx = 0
    total = len(raw_tokens)

    def _tok_to_ast(tok: Union[str, Tuple[str, str]]) -> dict:
        if isinstance(tok, tuple) and tok[0] == "GROUP":
            return parse_hwpeqn_to_equation_ast(tok[1])
        s_tok = str(tok)
        return {"kind": "sym", "val": _HWPEQN_TO_UNICODE_SYM.get(s_tok, s_tok)}

    while idx < total:
        tok = raw_tokens[idx]
        if tok == "sqrt" and idx + 1 < total:
            body_ast = _tok_to_ast(raw_tokens[idx + 1])
            items.append({"kind": "sqrt", "body": body_ast})
            idx += 2
            continue
        if tok in ("_", "^") and items and idx + 1 < total:
            base_node = items.pop()
            sub_node: Optional[dict] = None
            exp_node: Optional[dict] = None
            for _ in range(2):
                if idx < total and raw_tokens[idx] == "_" and idx + 1 < total:
                    sub_node = _tok_to_ast(raw_tokens[idx + 1])
                    idx += 2
                elif idx < total and raw_tokens[idx] == "^" and idx + 1 < total:
                    exp_node = _tok_to_ast(raw_tokens[idx + 1])
                    idx += 2
                else:
                    break
            if sub_node is not None and exp_node is not None:
                items.append({"kind": "subsup", "base": base_node, "sub": sub_node, "exp": exp_node})
            elif sub_node is not None:
                items.append({"kind": "sub", "base": base_node, "sub": sub_node})
            elif exp_node is not None:
                items.append({"kind": "sup", "base": base_node, "exp": exp_node})
            continue
        if tok == "over" and items and idx + 1 < total:
            num_node = items.pop()
            den_node = _tok_to_ast(raw_tokens[idx + 1])
            items.append({"kind": "frac", "num": num_node, "den": den_node})
            idx += 2
            continue
        items.append(_tok_to_ast(tok))
        idx += 1

    return normalize_equation_ast({"kind": "seq", "items": items})


def parse_omml_to_equation_ast(omml_el: ET.Element) -> dict:
    def _local_tag(el: ET.Element) -> str:
        return el.tag.split("}")[-1] if "}" in el.tag else el.tag

    def _parse_children(container: ET.Element) -> dict:
        child_items: List[dict] = []
        for ch in container:
            child_items.append(_parse_el(ch))
        return normalize_equation_ast({"kind": "seq", "items": child_items})

    def _parse_el(el: ET.Element) -> dict:
        lt = _local_tag(el)
        if lt in ("oMathPara", "oMath", "e", "num", "den", "sub", "sup"):
            return _parse_children(el)
        if lt == "r":
            txt = "".join(t.text or "" for t in el.iter() if _local_tag(t) == "t")
            return parse_latex_to_equation_ast(txt) if len(txt) > 1 and not txt.isalnum() else {"kind": "sym", "val": txt}
        if lt == "f":
            num_el = next((c for c in el if _local_tag(c) == "num"), None)
            den_el = next((c for c in el if _local_tag(c) == "den"), None)
            return normalize_equation_ast({
                "kind": "frac",
                "num": _parse_el(num_el) if num_el is not None else {"kind": "sym", "val": "1"},
                "den": _parse_el(den_el) if den_el is not None else {"kind": "sym", "val": "1"},
            })
        if lt == "rad":
            e_el = next((c for c in el if _local_tag(c) == "e"), None)
            return normalize_equation_ast({
                "kind": "sqrt",
                "body": _parse_el(e_el) if e_el is not None else {"kind": "sym", "val": "x"},
            })
        if lt == "sSup":
            e_el = next((c for c in el if _local_tag(c) == "e"), None)
            sup_el = next((c for c in el if _local_tag(c) == "sup"), None)
            return normalize_equation_ast({
                "kind": "sup",
                "base": _parse_el(e_el) if e_el is not None else {"kind": "sym", "val": "x"},
                "exp": _parse_el(sup_el) if sup_el is not None else {"kind": "sym", "val": "2"},
            })
        if lt == "sSub":
            e_el = next((c for c in el if _local_tag(c) == "e"), None)
            sub_el = next((c for c in el if _local_tag(c) == "sub"), None)
            return normalize_equation_ast({
                "kind": "sub",
                "base": _parse_el(e_el) if e_el is not None else {"kind": "sym", "val": "x"},
                "sub": _parse_el(sub_el) if sub_el is not None else {"kind": "sym", "val": "i"},
            })
        if lt == "sSubSup":
            e_el = next((c for c in el if _local_tag(c) == "e"), None)
            sub_el = next((c for c in el if _local_tag(c) == "sub"), None)
            sup_el = next((c for c in el if _local_tag(c) == "sup"), None)
            return normalize_equation_ast({
                "kind": "subsup",
                "base": _parse_el(e_el) if e_el is not None else {"kind": "sym", "val": "x"},
                "sub": _parse_el(sub_el) if sub_el is not None else {"kind": "sym", "val": "i"},
                "exp": _parse_el(sup_el) if sup_el is not None else {"kind": "sym", "val": "2"},
            })
        return _parse_children(el)

    return _parse_el(omml_el)


def compile_equation_ast_to_latex(ast_node: dict) -> str:
    node = normalize_equation_ast(ast_node)
    kind = node.get("kind", "sym")
    if kind == "sym":
        val = str(node.get("val", ""))
        return _UNICODE_TO_LATEX_SYM.get(val, val)
    if kind == "seq":
        return " ".join(compile_equation_ast_to_latex(it) for it in node.get("items", []))
    if kind == "frac":
        return f"\\frac{{{compile_equation_ast_to_latex(node['num'])}}}{{{compile_equation_ast_to_latex(node['den'])}}}"
    if kind == "sqrt":
        return f"\\sqrt{{{compile_equation_ast_to_latex(node['body'])}}}"
    if kind == "sup":
        return f"{{{compile_equation_ast_to_latex(node['base'])}}}^{{{compile_equation_ast_to_latex(node['exp'])}}}"
    if kind == "sub":
        return f"{{{compile_equation_ast_to_latex(node['base'])}}}_{{{compile_equation_ast_to_latex(node['sub'])}}}"
    if kind == "subsup":
        return (
            f"{{{compile_equation_ast_to_latex(node['base'])}}}"
            f"_{{{compile_equation_ast_to_latex(node['sub'])}}}"
            f"^{{{compile_equation_ast_to_latex(node['exp'])}}}"
        )
    return str(node.get("val", ""))


def compile_equation_ast_to_hwpeqn(ast_node: dict) -> str:
    node = normalize_equation_ast(ast_node)
    kind = node.get("kind", "sym")
    if kind == "sym":
        val = str(node.get("val", ""))
        return _UNICODE_TO_HWPEQN_SYM.get(val, val)
    if kind == "seq":
        return " ".join(compile_equation_ast_to_hwpeqn(it) for it in node.get("items", []))
    if kind == "frac":
        return f"{{{compile_equation_ast_to_hwpeqn(node['num'])}}} over {{{compile_equation_ast_to_hwpeqn(node['den'])}}}"
    if kind == "sqrt":
        return f"sqrt {{{compile_equation_ast_to_hwpeqn(node['body'])}}}"
    if kind == "sup":
        return f"{{{compile_equation_ast_to_hwpeqn(node['base'])}}} ^ {{{compile_equation_ast_to_hwpeqn(node['exp'])}}}"
    if kind == "sub":
        return f"{{{compile_equation_ast_to_hwpeqn(node['base'])}}} _ {{{compile_equation_ast_to_hwpeqn(node['sub'])}}}"
    if kind == "subsup":
        return (
            f"{{{compile_equation_ast_to_hwpeqn(node['base'])}}} "
            f"_ {{{compile_equation_ast_to_hwpeqn(node['sub'])}}} "
            f"^ {{{compile_equation_ast_to_hwpeqn(node['exp'])}}}"
        )
    return str(node.get("val", ""))


def compile_equation_ast_to_omml_element(ast_node: dict) -> ET.Element:
    def _m(tag: str) -> str:
        return f"{{{NS_OMML}}}{tag}"

    def _emit_into(parent: ET.Element, n: dict) -> None:
        node = normalize_equation_ast(n)
        kind = node.get("kind", "sym")
        if kind == "sym":
            val = strip_illegal_xml_control_chars(str(node.get("val", "")))
            if val:
                r_el = ET.SubElement(parent, _m("r"))
                t_el = ET.SubElement(r_el, _m("t"))
                t_el.text = val
            return
        if kind == "seq":
            for it in node.get("items", []):
                _emit_into(parent, it)
            return
        if kind == "frac":
            f_el = ET.SubElement(parent, _m("f"))
            num_el = ET.SubElement(f_el, _m("num"))
            _emit_into(num_el, node["num"])
            den_el = ET.SubElement(f_el, _m("den"))
            _emit_into(den_el, node["den"])
            return
        if kind == "sqrt":
            rad_el = ET.SubElement(parent, _m("rad"))
            rad_pr = ET.SubElement(rad_el, _m("radPr"))
            ET.SubElement(rad_pr, _m("degHide"), {_m("val"): "1"})
            ET.SubElement(rad_el, _m("deg"))
            e_el = ET.SubElement(rad_el, _m("e"))
            _emit_into(e_el, node["body"])
            return
        if kind == "sup":
            ssup_el = ET.SubElement(parent, _m("sSup"))
            e_el = ET.SubElement(ssup_el, _m("e"))
            _emit_into(e_el, node["base"])
            sup_el = ET.SubElement(ssup_el, _m("sup"))
            _emit_into(sup_el, node["exp"])
            return
        if kind == "sub":
            ssub_el = ET.SubElement(parent, _m("sSub"))
            e_el = ET.SubElement(ssub_el, _m("e"))
            _emit_into(e_el, node["base"])
            sub_el = ET.SubElement(ssub_el, _m("sub"))
            _emit_into(sub_el, node["sub"])
            return
        if kind == "subsup":
            ssubsup_el = ET.SubElement(parent, _m("sSubSup"))
            e_el = ET.SubElement(ssubsup_el, _m("e"))
            _emit_into(e_el, node["base"])
            sub_el = ET.SubElement(ssubsup_el, _m("sub"))
            _emit_into(sub_el, node["sub"])
            sup_el = ET.SubElement(ssubsup_el, _m("sup"))
            _emit_into(sup_el, node["exp"])
            return

    omath_el = ET.Element(_m("oMath"))
    _emit_into(omath_el, ast_node)
    return omath_el


def compile_equation_ast_to_linear_text(ast_node: dict) -> str:
    omml_el = compile_equation_ast_to_omml_element(ast_node)
    return "".join(omml_el.itertext())

