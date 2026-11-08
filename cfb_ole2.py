"""
Pure-Python MS-CFB (OLE2 Compound File Binary) Reader & Writer for HWP v5.0 Binary Containers.
Implements Red-Black Directory BST construction, MiniFAT/MiniStream (< 4,096 B),
512B Sector Chains, DIFAT (> 6.8 MB) chains ([BUG-05]), and automatic BinData Storage
provisioning ([BUG-06]).
"""

import os
import math
import copy
import struct
from dataclasses import dataclass
from typing import List, Tuple, Set


@dataclass
class CFBDirectoryEntry:
    name: str
    etype: int
    color: int
    left: int
    right: int
    child: int
    raw_entry: bytes
    data: bytes
    start_sec: int = 0
    size: int = 0


def cfb_name_key(name: str) -> Tuple[int, str]:
    return (len(name), name.upper())


def build_balanced_bst(indices: List[int], entries: List[dict]) -> int:
    if not indices:
        return 0xFFFFFFFF
    sorted_indices = sorted(indices, key=lambda idx: cfb_name_key(entries[idx]["name"]))
    mid = len(sorted_indices) // 2
    root_idx = sorted_indices[mid]
    entries[root_idx]["color"] = 1
    entries[root_idx]["left"] = build_balanced_bst(sorted_indices[:mid], entries)
    entries[root_idx]["right"] = build_balanced_bst(sorted_indices[mid + 1 :], entries)
    return root_idx


def collect_bst_nodes(root_idx: int, entries: List[dict]) -> List[int]:
    visited: Set[int] = set()
    stack: List[int] = [root_idx]
    out: List[int] = []
    while stack:
        curr = stack.pop()
        if curr == 0xFFFFFFFF or curr < 0 or curr >= len(entries) or curr in visited:
            continue
        visited.add(curr)
        out.append(curr)
        stack.append(entries[curr]["left"])
        stack.append(entries[curr]["right"])
    return out


def alloc_sector_chain(blob: bytes, sector_size: int, sectors: bytearray, fat: List[int]) -> Tuple[int, int]:
    if len(blob) == 0:
        return 0xFFFFFFFE, 0
    n_sec = math.ceil(len(blob) / float(sector_size))
    first_s = len(fat)
    padded = blob + b"\x00" * (n_sec * sector_size - len(blob))
    sectors.extend(padded)
    for s in range(first_s, first_s + n_sec - 1):
        fat.append(s + 1)
    fat.append(0xFFFFFFFE)
    return first_s, n_sec


class CFBArchive:
    def __init__(self, path: str):
        with open(path, "rb") as f:
            self.raw = f.read()
        self.sector_shift = struct.unpack_from("<H", self.raw, 30)[0]
        self.sector_size = 1 << self.sector_shift
        first_dir_sec = struct.unpack_from("<I", self.raw, 48)[0]
        first_minifat_sec = struct.unpack_from("<I", self.raw, 60)[0]
        num_minifat_sec = struct.unpack_from("<I", self.raw, 64)[0]
        first_difat_sec = struct.unpack_from("<I", self.raw, 68)[0]
        num_difat_sec = struct.unpack_from("<I", self.raw, 72)[0]

        difat = [struct.unpack_from("<I", self.raw, 76 + i * 4)[0] for i in range(109)]
        difat = [x for x in difat if x < 0xFFFFFFFE]

        # Traverse DIFAT sector chain for large CFB files (> 6.8 MB) ([BUG-05])
        entries_per_sec = self.sector_size // 4
        curr_difat = first_difat_sec
        difat_visited = 0
        while curr_difat < 0xFFFFFFFE and difat_visited < num_difat_sec:
            off_d = (curr_difat + 1) * self.sector_size
            d_vals = struct.unpack_from(f"<{entries_per_sec}I", self.raw, off_d)
            for fat_sec_id in d_vals[:-1]:
                if fat_sec_id < 0xFFFFFFFE:
                    difat.append(fat_sec_id)
            curr_difat = d_vals[-1]
            difat_visited += 1

        self.fat: List[int] = []
        for sec in difat:
            off = (sec + 1) * self.sector_size
            self.fat.extend(struct.unpack_from(f"<{entries_per_sec}I", self.raw, off))

        dir_bytes = self._get_chain(first_dir_sec)
        root_start = struct.unpack_from("<I", dir_bytes, 116)[0]
        root_size = struct.unpack_from("<Q", dir_bytes, 120)[0]
        ministream = self._get_chain(root_start)[:root_size] if root_size else b""
        minifat: List[int] = []
        if num_minifat_sec > 0 and first_minifat_sec < 0xFFFFFFFE:
            mf_bytes = self._get_chain(first_minifat_sec)
            minifat = list(struct.unpack_from(f"<{len(mf_bytes) // 4}I", mf_bytes, 0))

        self.entries: List[dict] = []
        for i in range(0, len(dir_bytes), 128):
            entry = bytearray(dir_bytes[i : i + 128])
            nlen = struct.unpack_from("<H", entry, 64)[0]
            if nlen == 0:
                continue
            name = bytes(entry[: max(0, nlen - 2)]).decode("utf-16le", errors="replace")
            etype = entry[66]
            color = entry[67]
            left_id, right_id, child_id = struct.unpack_from("<III", entry, 68)
            st_sec = struct.unpack_from("<I", entry, 116)[0]
            sz = struct.unpack_from("<Q", entry, 120)[0]
            stream_data = b""
            if etype == 2:
                if sz < 4096:
                    out = bytearray()
                    s = st_sec
                    while s < 0xFFFFFFFE and s < len(minifat):
                        out.extend(ministream[s * 64 : (s + 1) * 64])
                        s = minifat[s]
                    stream_data = bytes(out[:sz])
                else:
                    stream_data = self._get_chain(st_sec)[:sz]
            self.entries.append(
                {
                    "name": name,
                    "type": etype,
                    "color": color,
                    "left": left_id,
                    "right": right_id,
                    "child": child_id,
                    "raw_entry": bytes(entry),
                    "data": stream_data,
                }
            )

    def _get_chain(self, start: int) -> bytes:
        out = bytearray()
        s = start
        visited: Set[int] = set()
        while s < 0xFFFFFFFE and s < len(self.fat) and s not in visited:
            visited.add(s)
            off = (s + 1) * self.sector_size
            out.extend(self.raw[off : off + self.sector_size])
            s = self.fat[s]
        return bytes(out)

    def get_stream(self, name: str) -> bytes:
        for e in self.entries:
            if e["name"] == name:
                return e["data"]
        raise KeyError(f"CFB stream not found: {name}")

    def set_stream(self, name: str, data: bytes) -> None:
        for e in self.entries:
            if e["name"] == name:
                e["data"] = data
                return
        raise KeyError(f"CFB stream not found: {name}")


def write_cfb_file(cfb: CFBArchive, new_bindata_streams: List[Tuple[str, bytes]], out_path: str) -> dict:
    entries = [copy.deepcopy(e) for e in cfb.entries]

    # Ensure BinData Storage entry exists if binary streams are added ([BUG-06])
    bindata_dir_idx = next((idx for idx, e in enumerate(entries) if e["name"] == "BinData" and e["type"] == 1), -1)
    if new_bindata_streams and bindata_dir_idx < 0:
        bindata_dir_idx = len(entries)
        raw_bd = bytearray(128)
        bd_enc = "BinData".encode("utf-16le") + b"\x00\x00"
        raw_bd[: len(bd_enc)] = bd_enc
        struct.pack_into("<H", raw_bd, 64, len(bd_enc))
        raw_bd[66] = 1
        raw_bd[67] = 1
        struct.pack_into("<III", raw_bd, 68, 0xFFFFFFFF, 0xFFFFFFFF, 0xFFFFFFFF)
        entries.append(
            {
                "name": "BinData",
                "type": 1,
                "color": 1,
                "left": 0xFFFFFFFF,
                "right": 0xFFFFFFFF,
                "child": 0xFFFFFFFF,
                "raw_entry": bytes(raw_bd),
                "data": b"",
            }
        )
        root_top_nodes = collect_bst_nodes(entries[0]["child"], entries)
        if bindata_dir_idx not in root_top_nodes:
            root_top_nodes.append(bindata_dir_idx)
        entries[0]["child"] = build_balanced_bst(root_top_nodes, entries)

    bindata_children = (
        collect_bst_nodes(entries[bindata_dir_idx]["child"], entries)
        if bindata_dir_idx >= 0
        else [idx for idx, e in enumerate(entries) if e["name"].upper().startswith("BIN")]
    )

    for bname, bdata in new_bindata_streams:
        existing_idx = next((idx for idx, e in enumerate(entries) if e["name"].upper() == bname.upper()), -1)
        if existing_idx >= 0:
            entries[existing_idx]["data"] = bdata
            if existing_idx not in bindata_children:
                bindata_children.append(existing_idx)
        else:
            new_idx = len(entries)
            raw_e = bytearray(128)
            enc_name = bname.encode("utf-16le") + b"\x00\x00"
            raw_e[: len(enc_name)] = enc_name
            struct.pack_into("<H", raw_e, 64, len(enc_name))
            raw_e[66] = 2
            raw_e[67] = 1
            struct.pack_into("<III", raw_e, 68, 0xFFFFFFFF, 0xFFFFFFFF, 0xFFFFFFFF)
            entries.append(
                {
                    "name": bname,
                    "type": 2,
                    "color": 1,
                    "left": 0xFFFFFFFF,
                    "right": 0xFFFFFFFF,
                    "child": 0xFFFFFFFF,
                    "raw_entry": bytes(raw_e),
                    "data": bdata,
                }
            )
            bindata_children.append(new_idx)

    if bindata_children and 0 <= bindata_dir_idx < len(entries):
        entries[bindata_dir_idx]["child"] = build_balanced_bst(bindata_children, entries)

    sector_size = 512
    mini_sector_size = 64
    ministream = bytearray()
    minifat: List[int] = []

    for e in entries:
        if e["type"] == 2 and len(e["data"]) < 4096:
            sz = len(e["data"])
            if sz == 0:
                e["start_sec"] = 0xFFFFFFFE
                e["size"] = 0
            else:
                n_mini = math.ceil(sz / float(mini_sector_size))
                first_m = len(minifat)
                e["start_sec"] = first_m
                e["size"] = sz
                padded = e["data"] + b"\x00" * (n_mini * mini_sector_size - sz)
                ministream.extend(padded)
                for m in range(first_m, first_m + n_mini - 1):
                    minifat.append(m + 1)
                minifat.append(0xFFFFFFFE)

    sectors = bytearray()
    fat: List[int] = []

    root_start_sec, root_n_sec = alloc_sector_chain(bytes(ministream), sector_size, sectors, fat)
    entries[0]["start_sec"] = root_start_sec
    entries[0]["size"] = len(ministream)

    for e in entries:
        if e["type"] == 2 and len(e["data"]) >= 4096:
            st_sec, _ = alloc_sector_chain(e["data"], sector_size, sectors, fat)
            e["start_sec"] = st_sec
            e["size"] = len(e["data"])
        elif e["type"] == 1:
            e["start_sec"] = 0
            e["size"] = 0

    if minifat:
        mf_bytes = struct.pack(f"<{len(minifat)}I", *minifat)
        first_minifat_sec, num_minifat_sec = alloc_sector_chain(mf_bytes, sector_size, sectors, fat)
    else:
        first_minifat_sec, num_minifat_sec = 0xFFFFFFFE, 0

    dir_blob = bytearray()
    for e in entries:
        raw_e = bytearray(e["raw_entry"])
        raw_e[66] = e["type"]
        raw_e[67] = e["color"]
        struct.pack_into("<III", raw_e, 68, e["left"], e["right"], e["child"])
        struct.pack_into("<I", raw_e, 116, e["start_sec"])
        struct.pack_into("<Q", raw_e, 120, e["size"])
        dir_blob.extend(raw_e)
    while len(dir_blob) % sector_size != 0:
        dir_blob.extend(b"\x00" * 128)
    first_dir_sec, num_dir_sec = alloc_sector_chain(bytes(dir_blob), sector_size, sectors, fat)

    # Exact fixed-point FAT & DIFAT sector allocation ([BUG-05])
    total_non_fat_secs = len(fat)
    num_fat_secs = max(1, math.ceil(total_non_fat_secs / 127.0))
    num_difat_secs = math.ceil(max(0, num_fat_secs - 109) / 127.0) if num_fat_secs > 109 else 0
    while math.ceil((total_non_fat_secs + num_fat_secs + num_difat_secs) / 128.0) > num_fat_secs:
        num_fat_secs += 1
        num_difat_secs = math.ceil(max(0, num_fat_secs - 109) / 127.0) if num_fat_secs > 109 else 0

    fat_sec_indices: List[int] = []
    for _ in range(num_fat_secs):
        s_idx = len(fat)
        fat_sec_indices.append(s_idx)
        fat.append(0xFFFFFFFD)

    difat_sec_indices: List[int] = []
    for _ in range(num_difat_secs):
        d_idx = len(fat)
        difat_sec_indices.append(d_idx)
        fat.append(0xFFFFFFFC)

    while len(fat) % 128 != 0:
        fat.append(0xFFFFFFFF)

    for k in range(num_fat_secs):
        chunk = fat[k * 128 : (k + 1) * 128]
        sectors.extend(struct.pack("<128I", *chunk))

    if num_difat_secs > 0:
        rem_fat = fat_sec_indices[109:]
        for d_i in range(num_difat_secs):
            d_chunk = rem_fat[d_i * 127 : (d_i + 1) * 127]
            while len(d_chunk) < 127:
                d_chunk.append(0xFFFFFFFF)
            next_difat = difat_sec_indices[d_i + 1] if d_i + 1 < num_difat_secs else 0xFFFFFFFE
            d_chunk.append(next_difat)
            sectors.extend(struct.pack("<128I", *d_chunk))
        first_difat_sec = difat_sec_indices[0]
    else:
        first_difat_sec = 0xFFFFFFFE

    hdr = bytearray(cfb.raw[:512])
    struct.pack_into("<I", hdr, 40, 0)
    struct.pack_into("<I", hdr, 44, num_fat_secs)
    struct.pack_into("<I", hdr, 48, first_dir_sec)
    struct.pack_into("<I", hdr, 56, 4096)
    struct.pack_into("<I", hdr, 60, first_minifat_sec)
    struct.pack_into("<I", hdr, 64, num_minifat_sec)
    struct.pack_into("<I", hdr, 68, first_difat_sec)
    struct.pack_into("<I", hdr, 72, num_difat_secs)
    for i in range(109):
        val = fat_sec_indices[i] if i < len(fat_sec_indices) else 0xFFFFFFFF
        struct.pack_into("<I", hdr, 76 + i * 4, val)

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    with open(out_path, "wb") as f:
        f.write(hdr)
        f.write(sectors)

    return {
        "out_path": out_path,
        "total_bytes": len(hdr) + len(sectors),
        "num_entries": len(entries),
        "num_fat_secs": num_fat_secs,
        "num_difat_secs": num_difat_secs,
        "num_dir_secs": num_dir_sec,
        "root_n_sec": root_n_sec,
    }
