"""
HWP v5.0 Binary Record Parser, Dynamic Slot Scanner, Lifter, and Projector.
Preserves template 2x2 outer tables (treatAsChar=False, repeatHeader=False) and routes
attachment images into bottom attachment slots dynamically.
"""

import os
import zlib
import struct
from typing import List, Dict, Tuple, Optional

from .cfb_ole2 import CFBArchive, write_cfb_file
from .ir_schema import SlotSpecIR, TargetDocumentSpecIR
from .elastic_solver import ElasticLayoutState


def pack_hwp_record(tag_id: int, level: int, payload: bytes) -> bytes:
    sz = len(payload)
    if sz < 0xFFF:
        hdr = (tag_id & 0x3FF) | ((level & 0x3FF) << 10) | ((sz & 0xFFF) << 20)
        return struct.pack("<I", hdr) + payload
    hdr = (tag_id & 0x3FF) | ((level & 0x3FF) << 10) | (0xFFF << 20)
    return struct.pack("<II", hdr, sz) + payload


def parse_hwp_records(data: bytes) -> List[ list ]:
    pos = 0
    records: List[list] = []
    while pos + 4 <= len(data):
        hdr = struct.unpack_from("<I", data, pos)[0]
        pos += 4
        tag_id = hdr & 0x3FF
        level = (hdr >> 10) & 0x3FF
        size = (hdr >> 20) & 0xFFF
        if size == 0xFFF and pos + 4 <= len(data):
            size = struct.unpack_from("<I", data, pos)[0]
            pos += 4
        payload = data[pos : pos + size]
        records.append([tag_id, level, payload])
        pos += size
    return records


def extract_hwp_para_text(payload: bytes) -> str:
    out: List[str] = []
    i = 0
    while i + 1 < len(payload):
        ch = struct.unpack_from("<H", payload, i)[0]
        if ch in (1, 2, 3, 11, 12, 14, 15, 16, 17, 18, 21, 22, 23):
            i += 16
        elif ch in (4, 5, 6, 7, 8, 9, 19, 20):
            i += 16
        elif ch < 32:
            i += 2
        else:
            out.append(chr(ch))
            i += 2
    return "".join(out)


def parse_hex_rgb(hex_str: str) -> Tuple[int, int, int]:
    if not hex_str or hex_str == "auto":
        return (0, 0, 0)
    h = hex_str.lstrip("#")
    if len(h) == 6:
        return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    return (0, 0, 0)


def make_char_shape_payload(height_hu: int, bold: bool, italic: bool, rgb: Tuple[int, int, int]) -> bytes:
    clamped_h = max(200, min(7200, height_hu))
    buf = bytearray(74)
    buf[14:21] = b"\x64" * 7
    buf[28:35] = b"\x64" * 7
    struct.pack_into("<i", buf, 42, clamped_h)
    attr = (1 if italic else 0) | (2 if bold else 0)
    struct.pack_into("<I", buf, 46, attr)
    buf[50] = 10
    buf[51] = 10
    r, g, b = rgb
    struct.pack_into("<BBBB", buf, 52, r, g, b, 0)
    struct.pack_into("<I", buf, 56, 0x00000000)
    struct.pack_into("<I", buf, 60, 0xFFFFFFFF)
    struct.pack_into("<I", buf, 64, 0x00C0C0C0)
    struct.pack_into("<H", buf, 68, 2)
    struct.pack_into("<I", buf, 70, 0x00000000)
    return bytes(buf)


def make_para_shape_payload(
    align_code: int,
    left_hu: int,
    right_hu: int,
    indent_hu: int,
    before_hu: int,
    after_hu: int,
    line_spacing_pct: int,
) -> bytes:
    clamped_ls = max(110, min(200, line_spacing_pct))
    buf = bytearray(58)
    attr1 = 0x00000180 | ((align_code & 7) << 2)
    struct.pack_into("<I", buf, 0, attr1)
    struct.pack_into("<i", buf, 4, left_hu)
    struct.pack_into("<i", buf, 8, right_hu)
    struct.pack_into("<i", buf, 12, indent_hu)
    struct.pack_into("<i", buf, 16, before_hu)
    struct.pack_into("<i", buf, 20, after_hu)
    struct.pack_into("<i", buf, 24, clamped_ls)
    struct.pack_into("<H", buf, 28, 0)
    struct.pack_into("<H", buf, 30, 0)
    struct.pack_into("<H", buf, 32, 2)
    struct.pack_into("<I", buf, 46, 0)
    struct.pack_into("<I", buf, 50, clamped_ls)
    struct.pack_into("<I", buf, 54, 0)
    return bytes(buf)


def make_border_fill_payload(borders: dict, fill_hex: Optional[str]) -> bytes:
    buf = bytearray()
    buf.extend(struct.pack("<H", 0))
    for side in ("left", "right", "top", "bottom"):
        btype, bwidth, (r, g, b) = borders.get(side, (1, 1, (0xCB, 0xD5, 0xE1)))
        buf.extend(struct.pack("<BBBBBB", btype, bwidth, r, g, b, 0))
    buf.extend(struct.pack("<BBBBBB", 1, 0, 0, 0, 0, 0))
    if fill_hex and fill_hex != "auto":
        fr, fg, fb = parse_hex_rgb(fill_hex)
        buf.extend(struct.pack("<I", 1))
        buf.extend(struct.pack("<BBBB", fr, fg, fb, 0))
        buf.extend(struct.pack("<I", 0x00000000))
        buf.extend(struct.pack("<i", -1))
        buf.extend(b"\x00" * 5)
    else:
        buf.extend(struct.pack("<I", 0))
        buf.extend(struct.pack("<I", 0))
    return bytes(buf)


def make_image_border_fill_payload(bin_data_id: int) -> bytes:
    return (
        bytes.fromhex("00000000000000000000000000000000000000000000000000000100000000000200000005000000")
        + struct.pack("<H", bin_data_id)
        + b"\x00\x00\x00\x00\x00"
    )


def scan_hwp_template_slots(template_path: str) -> TargetDocumentSpecIR:
    """
    Dynamically scans Section0 of an HWP v5.0 template to discover 2x2 section tables (Q1..Q4)
    and 6x2 attachment tables ('붙임') without any hardcoded record indices.
    """
    cfb = CFBArchive(template_path)
    sec0_raw = zlib.decompress(cfb.get_stream("Section0"), -15)
    sec_recs = parse_hwp_records(sec0_raw)

    spec = TargetDocumentSpecIR(template_path=template_path, target_format="HWP_V5_BINARY")

    # Scan PAGE_DEF (Tag 73) if present
    for tag_id, _, payload in sec_recs:
        if tag_id == 73 and len(payload) >= 24:
            w_hu, h_hu, m_l, m_r, m_t, m_b = struct.unpack_from("<6I", payload, 0)
            spec.page_width_hu = w_hu
            spec.page_height_hu = h_hu
            spec.margin_left_hu = m_l
            spec.margin_right_hu = m_r
            spec.margin_top_hu = m_t
            spec.margin_bottom_hu = m_b
            break

    n_recs = len(sec_recs)
    i = 0
    while i < n_recs:
        tag_id, level, payload = sec_recs[i]
        if tag_id == 71 and level == 1 and len(payload) >= 4 and payload[:4] == b" lbt":
            ctrl_idx = i
            if i + 1 < n_recs and sec_recs[i + 1][0] == 77:
                tbl_payload = sec_recs[i + 1][2]
                _, n_rows, n_cols = struct.unpack_from("<IHH", tbl_payload, 0)
                # Find all direct cell LIST_HEADERs (tag=72, level=2) belonging to this L1 table
                cells_info: List[Tuple[int, int, int, int, int, int, int, str]] = []
                j = i + 2
                while j < n_recs and sec_recs[j][1] >= 2:
                    if sec_recs[j][0] == 72 and sec_recs[j][1] == 2 and len(sec_recs[j][2]) >= 24:
                        col_a, row_a, c_span, r_span = struct.unpack_from("<4H", sec_recs[j][2], 8)
                        w_cell, h_cell = struct.unpack_from("<2I", sec_recs[j][2], 16)
                        cell_txt = ""
                        if j + 2 < n_recs and sec_recs[j + 2][0] == 67:
                            cell_txt = extract_hwp_para_text(sec_recs[j + 2][2]).strip()
                        cells_info.append((j, row_a, col_a, r_span, c_span, w_cell, h_cell, cell_txt))
                    j += 1

                # Case A: 2x2 Section Outer Table (Row 0 has title, Row 1 has col_span=2 large body cell)
                if n_rows == 2 and n_cols == 2 and len(cells_info) >= 3:
                    title_cell = next((c for c in cells_info if c[1] == 0 and c[2] == 0), None)
                    body_cell = next((c for c in cells_info if c[1] == 1 and c[4] == 2), None)
                    if title_cell and body_cell:
                        t_str = title_cell[7]
                        slot_uid = f"SLOT_{len(spec.section_slots) + 1}"
                        if "Q1" in t_str or "문제" in t_str:
                            slot_uid = "Q1_BODY"
                        elif "Q2" in t_str or "실현" in t_str:
                            slot_uid = "Q2_BODY"
                        elif "Q3" in t_str or "성장" in t_str:
                            slot_uid = "Q3_BODY"
                        elif "Q4" in t_str or "향후" in t_str:
                            slot_uid = "Q4_BODY"
                        spec.section_slots.append(
                            SlotSpecIR(
                                slot_uid=slot_uid,
                                slot_role="SECTION_BODY_SLOT",
                                discovered_title_text=t_str,
                                outer_tbl_record_idx=ctrl_idx,
                                body_cell_record_idx=body_cell[0],
                                available_width_hu=max(40000, body_cell[5] - 1400),
                                available_height_hu=body_cell[6],
                                max_pages=2.0,
                            )
                        )
                # Case B: 6x2 Photo Attachment Table ('붙임')
                elif n_rows == 6 and n_cols == 2 and len(cells_info) >= 4:
                    for pair_r in range(3):
                        for col_c in range(2):
                            img_c = next((c for c in cells_info if c[1] == pair_r * 2 and c[2] == col_c), None)
                            cap_c = next((c for c in cells_info if c[1] == pair_r * 2 + 1 and c[2] == col_c), None)
                            if img_c and cap_c:
                                att_idx = len(spec.attachment_slots) + 1
                                spec.attachment_slots.append(
                                    SlotSpecIR(
                                        slot_uid=f"ATTACH_{att_idx}",
                                        slot_role="ATTACHMENT_MEDIA_SLOT",
                                        discovered_title_text=cap_c[7],
                                        outer_tbl_record_idx=ctrl_idx,
                                        body_cell_record_idx=img_c[0],
                                        caption_cell_record_idx=cap_c[0],
                                        available_width_hu=img_c[5],
                                        available_height_hu=img_c[6],
                                        max_pages=1.0,
                                    )
                                )
                i = j
                continue
        i += 1

    return spec


def lift_hwp_to_ir(hwp_path: str) -> Tuple[str, List[dict]]:
    cfb = CFBArchive(hwp_path)
    sec0_raw = zlib.decompress(cfb.get_stream("Section0"), -15)
    sec_recs = parse_hwp_records(sec0_raw)
    all_texts: List[str] = []
    blocks: List[dict] = []
    for tag_id, _, payload in sec_recs:
        if tag_id == 67:
            txt = extract_hwp_para_text(payload)
            all_texts.append(txt)
            if txt.strip():
                blocks.append(
                    {
                        "type": "p",
                        "jc": "both",
                        "left_tw": 0,
                        "right_tw": 0,
                        "indent_tw": 0,
                        "before_tw": 0,
                        "after_tw": 10,
                        "runs": [{"text": txt.strip(), "bold": False, "italic": False, "color": "000000"}],
                        "r_embed": None,
                    }
                )
    return "".join(all_texts), blocks


def project_ir_to_hwp(
    template_path: str,
    out_hwp_path: str,
    target_spec: TargetDocumentSpecIR,
    slot_blocks_map: Dict[str, Tuple[List[dict], ElasticLayoutState]],
    attachments: List[Tuple[bytes, str]],
) -> dict:
    cfb = CFBArchive(template_path)
    docinfo_raw = zlib.decompress(cfb.get_stream("DocInfo"), -15)
    sec0_raw = zlib.decompress(cfb.get_stream("Section0"), -15)

    doc_recs = parse_hwp_records(docinfo_raw)
    sec_recs = parse_hwp_records(sec0_raw)

    id_map_idx = next(i for i, r in enumerate(doc_recs) if r[0] == 17)
    id_counts = list(struct.unpack("<18I", doc_recs[id_map_idx][2]))

    new_bin_records: List[list] = []
    new_bindata_streams: List[Tuple[str, bytes]] = []
    bf_cache: Dict[bytes, int] = {}
    new_bf_records: List[list] = []
    base_bf_count = id_counts[8]

    def raw_deflate(b: bytes) -> bytes:
        co = zlib.compressobj(9, zlib.DEFLATED, -15)
        return co.compress(b) + co.flush()

    def get_or_add_bf(borders: dict, fill_hex: Optional[str]) -> int:
        payload = make_border_fill_payload(borders, fill_hex)
        if payload not in bf_cache:
            bf_id = base_bf_count + len(new_bf_records) + 1
            bf_cache[payload] = bf_id
            new_bf_records.append([20, 1, payload])
        return bf_cache[payload]

    attachment_bf_ids: List[int] = []
    base_bin_id = id_counts[0] + 1
    for idx_a, (img_bytes, _) in enumerate(attachments):
        bin_id = base_bin_id + idx_a
        new_bin_records.append(
            [18, 1, struct.pack("<HH", 1, bin_id) + struct.pack("<H", 3) + "png".encode("utf-16le")]
        )
        new_bindata_streams.append((f"BIN{bin_id:04X}.png", raw_deflate(img_bytes)))
        bf_img_payload = make_image_border_fill_payload(bin_id)
        bf_img_id = base_bf_count + len(new_bf_records) + 1
        new_bf_records.append([20, 1, bf_img_payload])
        attachment_bf_ids.append(bf_img_id)

    cs_cache: Dict[bytes, int] = {}
    new_cs_records: List[list] = []
    base_cs_count = id_counts[9]

    def get_or_add_cs(height_hu: int, bold: bool, italic: bool, color_hex: str) -> int:
        rgb = parse_hex_rgb(color_hex)
        payload = make_char_shape_payload(height_hu, bold, italic, rgb)
        if payload not in cs_cache:
            cs_id = base_cs_count + len(new_cs_records)
            cs_cache[payload] = cs_id
            new_cs_records.append([21, 1, payload])
        return cs_cache[payload]

    ps_cache: Dict[bytes, int] = {}
    new_ps_records: List[list] = []
    base_ps_count = id_counts[13]

    def get_or_add_ps(
        jc: str, left_tw: int, right_tw: int, indent_tw: int, before_tw: int, after_tw: int, line_spacing_pct: int
    ) -> int:
        align_map = {"both": 0, "distribute": 0, "left": 1, "right": 2, "center": 3}
        align_code = align_map.get(jc, 0)
        payload = make_para_shape_payload(
            align_code,
            left_tw * 5,
            right_tw * 5,
            indent_tw * 5,
            before_tw * 5,
            after_tw * 5,
            line_spacing_pct,
        )
        if payload not in ps_cache:
            ps_id = base_ps_count + len(new_ps_records)
            ps_cache[payload] = ps_id
            new_ps_records.append([25, 1, payload])
        return ps_cache[payload]

    def build_para_records(
        p_block: dict, para_level: int, height_hu: int, line_spacing_pct: int, is_last_in_list: bool
    ) -> List[list]:
        ps_id = get_or_add_ps(
            p_block["jc"],
            p_block["left_tw"],
            p_block["right_tw"],
            p_block["indent_tw"],
            p_block["before_tw"],
            p_block["after_tw"],
            line_spacing_pct,
        )
        runs = p_block["runs"]
        if not runs:
            cs_id = get_or_add_cs(height_hu, False, False, "000000")
            nchars_val = 1 | (0x80000000 if is_last_in_list else 0)
            phdr = struct.pack("<IIHBBHHHIH", nchars_val, 0, ps_id, 0, 0, 1, 0, 0, 0, 0)
            return [
                [66, para_level, phdr],
                [67, para_level + 1, b"\r\x00"],
                [68, para_level + 1, struct.pack("<II", 0, cs_id)],
            ]

        text_buf = bytearray()
        cs_pairs: List[Tuple[int, int]] = []
        char_pos = 0
        last_cs_id: Optional[int] = None
        for r in runs:
            txt = r["text"]
            if not txt:
                continue
            cs_id = get_or_add_cs(height_hu, r["bold"], r["italic"], r["color"])
            if cs_id != last_cs_id:
                cs_pairs.append((char_pos, cs_id))
                last_cs_id = cs_id
            text_buf.extend(txt.encode("utf-16le"))
            char_pos += len(txt)
        text_buf.extend(b"\r\x00")
        char_count = char_pos + 1
        if not cs_pairs:
            cs_pairs.append((0, get_or_add_cs(height_hu, False, False, "000000")))

        nchars_val = char_count | (0x80000000 if is_last_in_list else 0)
        phdr = struct.pack("<IIHBBHHHIH", nchars_val, 0, ps_id, 0, 0, len(cs_pairs), 0, 0, 0, 0)
        pcs = b"".join(struct.pack("<II", pos_i, cid) for pos_i, cid in cs_pairs)
        return [
            [66, para_level, phdr],
            [67, para_level + 1, bytes(text_buf)],
            [68, para_level + 1, pcs],
        ]

    def build_inner_table_records(
        tbl_block: dict, para_level: int, tbl_ls_pct: int, is_last_in_list: bool, inst_id: int
    ) -> List[list]:
        ps_id = get_or_add_ps("center", 0, 0, 0, 0, 6, tbl_ls_pct)
        cs_id = get_or_add_cs(target_spec.table_font_size_hu, False, False, "000000")
        nchars_val = 9 | (0x80000000 if is_last_in_list else 0)
        phdr = struct.pack("<IIHBBHHHIH", nchars_val, 0x800, ps_id, 0, 0, 1, 0, 0, 0, 0)
        ptxt = struct.pack("<H", 11) + b" tbl" + b"\x00" * 8 + struct.pack("<HH", 11, 13)
        pcs = struct.pack("<II", 0, cs_id)

        ctrl_payload = bytearray(
            bytes.fromhex(
                "206c627411232a18000000000000000006ba0000840e0100000000000000000000000000d3faac43000000000000"
            )
        )
        struct.pack_into("<II", ctrl_payload, 16, tbl_block["width_hu"], tbl_block["height_hu"])
        struct.pack_into("<HHHH", ctrl_payload, 28, 0, 0, 15, 15)
        struct.pack_into("<I", ctrl_payload, 36, inst_id)

        n_rows = tbl_block["n_rows"]
        n_cols = tbl_block["n_cols"]
        tbl_payload = bytearray()
        tbl_payload.extend(struct.pack("<IHH", 0x04000002, n_rows, n_cols))
        tbl_payload.extend(struct.pack("<HHHHH", 0, 60, 60, 25, 25))
        for rc in tbl_block["row_cols"]:
            tbl_payload.extend(struct.pack("<H", rc))
        tbl_payload.extend(struct.pack("<H", 4))
        tbl_payload.extend(struct.pack("<H", 0))

        out_recs: List[list] = [
            [66, para_level, phdr],
            [67, para_level + 1, ptxt],
            [68, para_level + 1, pcs],
            [71, para_level + 1, bytes(ctrl_payload)],
            [77, para_level + 2, bytes(tbl_payload)],
        ]

        for cell in tbl_block["cells"]:
            bf_id = cell.get("override_bf_id") or get_or_add_bf(cell["borders"], cell["fill_hex"])
            n_cell_paras = len(cell["paras"])
            mar_l, mar_r, mar_t, mar_b = cell["mar"]
            lhdr = bytearray(47)
            struct.pack_into("<H", lhdr, 0, n_cell_paras)
            struct.pack_into("<I", lhdr, 4, 0x00000020)
            struct.pack_into("<HHHH", lhdr, 8, cell["col"], cell["row"], cell["col_span"], cell["row_span"])
            struct.pack_into("<II", lhdr, 16, cell["width_hu"], cell["height_hu"])
            struct.pack_into("<HHHH", lhdr, 24, mar_l, mar_r, mar_t, mar_b)
            struct.pack_into("<H", lhdr, 32, bf_id)
            struct.pack_into("<I", lhdr, 34, cell["width_hu"])
            out_recs.append([72, para_level + 2, bytes(lhdr)])

            for p_i, c_para in enumerate(cell["paras"]):
                is_last_p = p_i == n_cell_paras - 1
                out_recs.extend(
                    build_para_records(
                        c_para, para_level + 2, target_spec.table_font_size_hu, tbl_ls_pct, is_last_p
                    )
                )

        return out_recs

    # Agenda #1: Enforce treatAsChar=False and repeatHeader=False on all discovered outer section tables
    for s_slot in target_spec.section_slots:
        t_idx = s_slot.outer_tbl_record_idx
        p = bytearray(sec_recs[t_idx][2])
        attr = struct.unpack_from("<I", p, 4)[0]
        attr &= ~0x1
        struct.pack_into("<I", p, 4, attr)
        sec_recs[t_idx][2] = bytes(p)

        if t_idx + 1 < len(sec_recs) and sec_recs[t_idx + 1][0] == 77:
            p_tbl = bytearray(sec_recs[t_idx + 1][2])
            struct.pack_into("<I", p_tbl, 0, 0x04000002)
            sec_recs[t_idx + 1][2] = bytes(p_tbl)

        # Page break before outer paragraph (t_idx - 3)
        if t_idx >= 3 and sec_recs[t_idx - 3][0] == 66 and sec_recs[t_idx - 3][1] == 0:
            p_l0 = bytearray(sec_recs[t_idx - 3][2])
            p_l0[11] = 4
            sec_recs[t_idx - 3][2] = bytes(p_l0)

    # Agenda #2: Route attachments to discovered attachment slots
    for idx_a, (_, cap_str) in enumerate(attachments):
        if idx_a >= len(target_spec.attachment_slots):
            break
        att_slot = target_spec.attachment_slots[idx_a]
        img_rec_idx = att_slot.body_cell_record_idx
        cap_rec_idx = att_slot.caption_cell_record_idx
        bf_img_id = attachment_bf_ids[idx_a]

        p_img = bytearray(sec_recs[img_rec_idx][2])
        struct.pack_into("<H", p_img, 32, bf_img_id)
        sec_recs[img_rec_idx][2] = bytes(p_img)
        if img_rec_idx + 2 < len(sec_recs) and sec_recs[img_rec_idx + 2][0] == 67:
            sec_recs[img_rec_idx + 2][2] = b"\r\x00"
            p_hdr = bytearray(sec_recs[img_rec_idx + 1][2])
            struct.pack_into("<I", p_hdr, 0, 0x80000001)
            sec_recs[img_rec_idx + 1][2] = bytes(p_hdr)

        if cap_rec_idx >= 0 and cap_rec_idx + 2 < len(sec_recs) and sec_recs[cap_rec_idx + 2][0] == 67:
            sec_recs[cap_rec_idx + 2][2] = cap_str.encode("utf-16le") + b"\r\x00"
            p_chdr = bytearray(sec_recs[cap_rec_idx + 1][2])
            struct.pack_into("<I", p_chdr, 0, 0x80000000 | (len(cap_str) + 1))
            sec_recs[cap_rec_idx + 1][2] = bytes(p_chdr)

    # Build replacement map for body cell records (sorted in reverse index order so splices don't shift earlier indices)
    splices: List[Tuple[int, int, List[list]]] = []
    inst_seed = 0x50001000
    for s_slot in target_spec.section_slots:
        if s_slot.slot_uid not in slot_blocks_map:
            continue
        blocks, el_state = slot_blocks_map[s_slot.slot_uid]
        if not blocks:
            continue
        b_idx = s_slot.body_cell_record_idx
        p_cell = bytearray(sec_recs[b_idx][2])
        struct.pack_into("<H", p_cell, 0, len(blocks))
        struct.pack_into("<HHHH", p_cell, 24, 350, 350, 100, 100)
        sec_recs[b_idx][2] = bytes(p_cell)

        cell_recs: List[list] = []
        n_b = len(blocks)
        for idx_blk, blk in enumerate(blocks):
            is_last = idx_blk == n_b - 1
            if blk["type"] == "p":
                cell_recs.extend(
                    build_para_records(
                        blk, 2, target_spec.body_font_size_hu, el_state.body_line_spacing_pct, is_last
                    )
                )
            elif blk["type"] == "tbl":
                cell_recs.extend(
                    build_inner_table_records(blk, 2, el_state.table_line_spacing_pct, is_last, inst_seed)
                )
                inst_seed += 1

        # Determine existing paragraph records inside b_idx (typically b_idx+1 .. b_idx+2)
        end_idx = b_idx + 1
        while end_idx < len(sec_recs) and sec_recs[end_idx][1] >= 2 and sec_recs[end_idx][0] != 72:
            end_idx += 1
        splices.append((b_idx + 1, end_idx, cell_recs))

    for start_i, end_i, replacement_recs in sorted(splices, key=lambda x: x[0], reverse=True):
        sec_recs = sec_recs[:start_i] + replacement_recs + sec_recs[end_i:]

    id_counts[0] += len(new_bin_records)
    id_counts[8] += len(new_bf_records)
    id_counts[9] += len(new_cs_records)
    id_counts[13] += len(new_ps_records)
    doc_recs[id_map_idx][2] = struct.pack("<18I", *id_counts)

    def insert_after_last_tag(rec_list: List[list], target_tag: int, new_items: List[list]) -> List[list]:
        if not new_items:
            return rec_list
        last_i = max(i for i, r in enumerate(rec_list) if r[0] == target_tag)
        return rec_list[: last_i + 1] + new_items + rec_list[last_i + 1 :]

    doc_recs = insert_after_last_tag(doc_recs, 18, new_bin_records)
    doc_recs = insert_after_last_tag(doc_recs, 20, new_bf_records)
    doc_recs = insert_after_last_tag(doc_recs, 21, new_cs_records)
    doc_recs = insert_after_last_tag(doc_recs, 25, new_ps_records)

    new_docinfo_bytes = b"".join(pack_hwp_record(t, l, p) for t, l, p in doc_recs)
    new_sec0_bytes = b"".join(pack_hwp_record(t, l, p) for t, l, p in sec_recs)

    cfb.set_stream("DocInfo", raw_deflate(new_docinfo_bytes))
    cfb.set_stream("Section0", raw_deflate(new_sec0_bytes))

    cfb_info = write_cfb_file(cfb, new_bindata_streams, out_hwp_path)
    return {
        "out_hwp_path": out_hwp_path,
        "bytes": cfb_info["total_bytes"],
        "added_bin": len(new_bin_records),
        "added_bf": len(new_bf_records),
        "added_cs": len(new_cs_records),
        "added_ps": len(new_ps_records),
    }
