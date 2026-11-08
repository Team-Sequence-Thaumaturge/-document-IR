"""
Pure-Python MS-CFB (OLE2 Compound File Binary) Reader & Writer for HWP v5.0 Binary Containers.
Implements Red-Black Directory BST construction, MiniFAT/MiniStream (< 4,096 B), and 512B Sector Chains.
"""

import os
import math
import copy
import struct
from dataclasses import dataclass
from typing import List, Tuple


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
        difat = [struct.unpack_from("<I", self.raw, 76 + i * 4)[0] for i in range(109)]
        difat = [x for x in difat if x < 0xFFFFFFFE]
        self.fat: List[int] = []
        for sec in difat:
            off = (sec + 1) * self.sector_size
            self.fat.extend(struct.unpack_from(f"<{self.sector_size // 4}I", self.raw, off))

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
        while s < 0xFFFFFFFE and s < len(self.fat):
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
    bindata_children = [idx for idx, e in enumerate(entries) if e["name"].upper().startswith("BIN")]
    for bname, bdata in new_bindata_streams:
        existing_idx = next((idx for idx, e in enumerate(entries) if e["name"].upper() == bname.upper()), -1)
        if existing_idx >= 0:
            entries[existing_idx]["data"] = bdata
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

    bindata_dir_idx = next((idx for idx, e in enumerate(entries) if e["name"] == "BinData"), 1)
    if bindata_children and bindata_dir_idx < len(entries):
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

    total_non_fat_secs = len(fat)
    num_fat_secs = math.ceil((total_non_fat_secs + 16) / 128.0)
    fat_sec_indices: List[int] = []
    for _ in range(num_fat_secs):
        s_idx = len(fat)
        fat_sec_indices.append(s_idx)
        fat.append(0xFFFFFFFD)
    while len(fat) % 128 != 0:
        fat.append(0xFFFFFFFF)
    for k in range(num_fat_secs):
        chunk = fat[k * 128 : (k + 1) * 128]
        sectors.extend(struct.pack("<128I", *chunk))

    hdr = bytearray(cfb.raw[:512])
    struct.pack_into("<I", hdr, 40, 0)
    struct.pack_into("<I", hdr, 44, num_fat_secs)
    struct.pack_into("<I", hdr, 48, first_dir_sec)
    struct.pack_into("<I", hdr, 56, 4096)
    struct.pack_into("<I", hdr, 60, first_minifat_sec)
    struct.pack_into("<I", hdr, 64, num_minifat_sec)
    struct.pack_into("<I", hdr, 68, 0xFFFFFFFE)
    struct.pack_into("<I", hdr, 72, 0)
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
        "num_dir_secs": num_dir_sec,
        "root_n_sec": root_n_sec,
    }
