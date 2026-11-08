# T-DIR (`document-IR`) — Topological Document Intermediate Representation MCP Server

**[last patch : 2026-11-08 17:52:00 UTC]**

🌐 **[English (./README.md)](./README.md)** | 🇰🇷 **[한국어 (./README_KO.md)](./README_KO.md)**

> [!WARNING]
> **Work in Progress (Experimental Beta — `v3.0.0-beta`)**
> This repository is currently under active development and architectural verification. **It is NOT yet feature-complete or production-hardened.**
> - Several format lifters and projectors (such as direct `<hp:pic>` binary stream embedding inside standalone `.hwpx` packages, full `.xlsx` spreadsheet cell-style lifting, and multi-engine equation AST conversion) are still incomplete.
> - Internal data structures (`SectionIR`, `SlotSpecIR`, `SourceDocumentSpecIR`, `TargetDocumentSpecIR`, `UserSpecOverrideIR`, `ElasticLayoutState`) and MCP tool parameters may undergo breaking changes as beta regression testing progresses.

---

## 1. Overview & Motivation

**T-DIR (Topological Document Intermediate Representation)** is a Model Context Protocol (MCP) server and document transformation engine designed to bridge structural and layout incompatibilities across Korean government/enterprise document formats (**HWP v5.0 Binary `.hwp`**, **OWPML `.hwpx`**) and global document formats (**OOXML `.docx`**, **Markdown `.md`**, **CSV/TSV `.csv`/`.tsv`**).

### Why Traditional 1:1 Document Converters Fail
1. **Box-Model & Coordinate Unit Discrepancies**: Direct translation between `HWPUNIT` ($1\text{ pt} = 100\text{ HU}$), `Twips` ($1\text{ pt} = 20\text{ dxa}$), `EMU` ($1\text{ HU} = 127\text{ EMU}$), and CSS pixels fails when nested tables have asymmetric cell margins, merged spans (`gridSpan` / `vMerge`), or empty decorative spacer columns.
2. **Line-Grid Snap Inflation (`snapToGrid`)**: Fonts with large vertical bounding boxes trigger line-grid snapping in MS Word (`w:lineRule="auto"`), inflating body text height per line and causing content to overflow fixed-page form templates.
3. **Rigid Form Template Boundaries**: Official Korean business plan and grant templates (`.hwp`) wrap each section inside fixed `2x2` outer layout tables and place media attachments in bottom `6x2` grid slots (`붙임`). Naive copy-pasting breaks outer table borders, duplicates headers across pages, or truncates overflowing text.
4. **LLM Context & Token Overhead**: Feeding tens of thousands of lines of raw OOXML (`word/document`) or OWPML (`Contents/section0`) XML into an LLM wastes context window tokens and risks accidental text hallucination or omission.

---

## 2. Core Architecture & Zero-Hardcoding Specification Resolution

T-DIR separates **content semantics**, **dynamic specification resolution**, **2D grid topology**, and **physical metric projection** into strictly isolated layers with **zero hardcoded font names, pixel/HWPUNIT constants, or document-specific strings**:

```mermaid
flowchart LR
    A["Source & Target Files<br/>(.docx, .hwp, .hwpx, .md, .csv)"] --> B["Layer 1: Dynamic Spec & Slot Scanner<br/>(DocInfo / Section0 / OOXML Styles)"]
    U["User MCP Spec Overrides<br/>(UserSpecOverrideIR)"] --> C["Priority Spec Resolver<br/>(User > Target > Source)"]
    B --> C
    C --> D["Layer 2: 2D Grid Normalizer &<br/>5-Stage Elastic Solver (Early-Stop)"]
    D --> E["Layer 3: Deterministic Projectors<br/>(.hwp CFB, .hwpx, .docx, .md)"]
    E --> F["Cryptographic Verifier<br/>(Non-Whitespace SHA-256 + PDF Pages)"]
```

### 2.1. 3-Tier Specification Priority Resolution
Instead of hardcoding fonts, font sizes, or page dimensions, T-DIR extracts physical specifications from both the target template and source document at runtime and resolves them against explicit user MCP overrides ([`./ir_schema.py`](./ir_schema.py)):

$$\text{EffectiveSpec} = \text{Priority 1 }(\text{UserSpecOverrideIR}) \succ \text{Priority 2 }(\text{TargetDocumentSpecIR}) \succ \text{Priority 3 }(\text{SourceDocumentSpecIR})$$

1. **Priority 1 — `UserSpecOverrideIR` (Explicit User MCP Specification)**: Explicit formatting, typography (body font family, table font family, body font size in points, table font size in points), line spacing percentage bounds, page geometry (page width, height, and margins in millimeters), per-section page budgets, outer table preservation policy, and attachment captions passed by the user or agent via MCP tool calls.
2. **Priority 2 — `TargetDocumentSpecIR` ([`./codec_hwp_v5.py`](./codec_hwp_v5.py))**: Dynamically extracted from the target template's binary streams:
   - `DocInfo` stream: `Tag 19` (UTF-16LE font families), `Tag 21` (character shape font sizes in `HU` and RGB colors), and `Tag 25` (paragraph shape line spacing percentages).
   - `Section0` stream: `Tag 73` (page definition width, height, and margins in `HU`), preamble/cover blocks, `2x2` section outer tables (outer width, column widths, discovered title text, discovered subtitle text, and available inner width in `HU`), and `6x2` attachment slot geometries (`ATTACH 1..6`).
3. **Priority 3 — `SourceDocumentSpecIR` ([`./codec_ooxml_owpml.py`](./codec_ooxml_owpml.py))**: Dynamically extracted from the source document (`word/styles` and `word/document` `sectPr` in `.docx`, or `DocInfo`/`Section0` in `.hwp`), providing fallback prose/table font families and sizes when the target template's body cells are empty placeholders.

### 2.2. Pure 2D Topological Grid Classifiers ([`./elastic_solver.py`](./elastic_solver.py))
All table and paragraph structural classifications operate purely on 2D matrix geometry and punctuation topology without matching any document-specific text:
- **Stair-Step Sparse Chart Detector**: Detects 2D grids where a single full-width axis row divides upper/lower rows containing sparse single non-empty cells (`1..2` active columns per row) and collapses them into a dense 4-row chronological table.
- **Vertical Step-Arrow Detector**: Detects single-column tables (odd row count $\ge 3$) where every odd-indexed row (`r = 1, 3, ...`) is a short connector/arrow indicator ($\le 25$ chars) and folds the connector inline into the preceding step row.
- **Multi-Column Spacer Card Collapser**: Automatically collapses multi-column tables where every non-empty row contains at most 1 non-empty cell into a clean full-width `1`-column card stack.
- **Standalone Subheading Classifier**: Identifies standalone short subheadings ($\le 40$ chars, bold or bracket/bullet-prefixed, not ending with sentence terminal punctuation) for Stage 3 inline lead-in merging.

### 2.3. 5-Stage Minimum-Invasion Elastic Layout Solver (`MinimumInvasionElasticSolver`)
To fit dense source sections into strict page budgets **without over-squishing paragraphs**, `MinimumInvasionElasticSolver` scales line heights and character capacities proportionally with the resolved body and table font sizes, executing a progressive 5-stage loop and **stopping immediately (`Early-Stop`)** as soon as the estimated height fits within the target page budget:
- **Stage 1 (Spacers Only)**: Removes empty spacer paragraphs and empty table rows/columns, and disables document line-grid snapping (`snapToGrid = False`).
- **Stage 2 (Spacing Interpolation Only)**: Interpolates body line spacing and table line spacing between the resolved maximum and minimum percentage bounds while keeping **100% of original paragraph and table cell line breaks intact**.
- **Stage 3 (Subheading Inline Early-Stop)**: Converts standalone short subheadings into bold inline lead-ins (`[Bold Subheading]  [Body Paragraph]`) using whitespace-only joining (`"  "`) so non-whitespace character hashes remain identical.
- **Stage 4 (Partial Cell Merge Early-Stop)**: Sorts multi-paragraph table cells by shortest combined character length and merges one adjacent pair at a time, checking the height budget after **every single merge**.
- **Stage 5 (Minimal Body Continuation Early-Stop)**: Only if still over budget after Stage 4, merges the shortest adjacent body continuation paragraphs one by one until the page budget is met.

---

## 3. Repository Structure

| Module Path | Description |
| :--- | :--- |
| [`./__init__.py`](./__init__.py) | Package initialization and exported classes (`TdirBetaMcpServer`, `MinimumInvasionElasticSolver`, `SourceDocumentSpecIR`, `TargetDocumentSpecIR`, `UserSpecOverrideIR`). |
| [`./ir_schema.py`](./ir_schema.py) | T-DIR v3.0 data structures (`TextRunIR`, `ParagraphIR`, `TableCellIR`, `TableGridIR`, `MediaObjectIR`, `SectionIR`, `SlotSpecIR`, `SourceDocumentSpecIR`, `TargetDocumentSpecIR`, `UserSpecOverrideIR`, `ConversionJobIR`), dynamic specification priority resolver, unit converters, and cryptographic non-whitespace `SHA-256` verifier. |
| [`./cfb_ole2.py`](./cfb_ole2.py) | Pure-Python MS-CFB (`CFBArchive` OLE2 Compound File Binary v3) reader and writer supporting 512B standard sectors, 64B mini-streams, and balanced Red-Black directory trees. |
| [`./elastic_solver.py`](./elastic_solver.py) | Pure topological 2D grid normalizers and `MinimumInvasionElasticSolver` 5-stage Early-Stop layout engine (`ElasticLayoutState`). |
| [`./codec_hwp_v5.py`](./codec_hwp_v5.py) | HWP v5.0 binary record parser, `DocInfo` dynamic typography extractor, dynamic font face registrar, `Section0` template slot scanner, lifter, and binary projector. |
| [`./codec_ooxml_owpml.py`](./codec_ooxml_owpml.py) | OOXML (`.docx`) style/page scanner, dynamic section lifter, and `TargetDocumentSpecIR`-driven `.docx` / `.hwpx` projectors. |
| [`./codec_tabular_markup.py`](./codec_tabular_markup.py) | Markdown (`.md`) and Tabular (`.csv`, `.tsv`) lifters and projectors. |
| [`./mcp_server.py`](./mcp_server.py) | Standalone JSON-RPC 2.0 stdio MCP server (`TdirBetaMcpServer` / `tdir-document-beta`) exposing 5 tools and the `UserSpecOverrideIR` JSON schema. |
| [`./test_regression_suite.py`](./test_regression_suite.py) | End-to-end regression suite verifying dynamic typography/slot probing, `UserSpecOverrideIR` priority overrides, Early-Stop solving, multi-format projection, `SHA-256` text immutability, and 8-page PDF boundary compliance. |

---

## 4. MCP Tools & `UserSpecOverrideIR` Schema

The server ([`./mcp_server.py`](./mcp_server.py)) communicates over standard JSON-RPC 2.0 (`stdio`) and exposes 5 tools:

```python
# 1. Probe source documents and target template typography, page geometry, and slots dynamically
tdir_probe_files(source_paths: list[str], target_path: str) -> dict

# 2. Resolve effective spec (UserSpecOverrideIR > TargetDocumentSpecIR > SourceDocumentSpecIR) and preview Elastic Solver
tdir_build_mapping_plan(
    job_id: str,
    user_spec_overrides: dict | None = None,
    user_constraints: dict | None = None,
) -> dict

# 3. Execute deterministic projection into ["hwp", "hwpx", "docx", "md"]
tdir_execute_conversion(
    job_id: str,
    output_prefix: str,
    output_formats: list[str] | None = None,
    page_body_height_hu: int | None = None,
    base_hwpx_path: str | None = None,
    user_spec_overrides: dict | None = None,
) -> dict

# 4. Verify 100.00% non-whitespace SHA-256 text immutability and rendered PDF page count
tdir_verify_fidelity_and_layout(job_id: str, export_pdf: bool = True) -> dict

# 5. Pinpoint adjustment of paragraph merge lock or table column widths
tdir_adjust_block_layout(
    job_id: str,
    section_uid: str,
    block_index: int,
    lock_unmerged: bool | None = None,
    col_widths_hu: list[int] | None = None,
) -> dict
```

### Supported `UserSpecOverrideIR` Fields
When calling Tool 2 or Tool 3, users or agents can pass the override dictionary to explicitly override any aspect of the target specification:

```json
{
  "body_font_family": "함초롬바탕",
  "table_font_family": "함초롬돋움",
  "body_font_size_pt": 11.5,
  "table_font_size_pt": 9.5,
  "body_line_spacing_min_pct": 122,
  "body_line_spacing_max_pct": 160,
  "table_line_spacing_min_pct": 112,
  "table_line_spacing_max_pct": 125,
  "page_width_mm": 210.0,
  "page_height_mm": 297.0,
  "margins_mm": {"top": 15.0, "bottom": 15.0, "left": 20.0, "right": 20.0},
  "section_page_budgets": {"Q1": 2.0, "Q2": 2.0},
  "exclude_sections": ["PREAMBLE", "Q3", "Q4", "REFERENCES"],
  "preserve_outer_tables": true,
  "route_media_to_attachment_slots": true,
  "max_attachment_images": 3,
  "attachment_captions": [
    "[붙임 1] 시스템 아키텍처 다이어그램",
    "[붙임 2] 3단계 워크플로우 및 타겟 시장 로드맵"
  ]
}
```

---

## 5. Quickstart & Usage

### Prerequisites
- Python 3.10+ (Zero external third-party pip dependencies required for core `.hwp`, `.hwpx`, `.docx`, `.md`, `.csv` lifting and projection; uses Python standard library modules `zlib`, `struct`, `zipfile`, `hashlib`).
- Optional: Microsoft Word (on Windows) if headless PDF page-count verification is enabled in Tool 4.

### Running as an MCP Server
Add `tdir-document-beta` to your MCP client configuration:

```json
{
  "mcpServers": {
    "tdir-document-beta": {
      "command": "python",
      "args": ["mcp_server.py"],
      "cwd": "/path/to/-document-IR"
    }
  }
}
```

---

## 6. Known Limitations & Development Roadmap (WIP)

Because this project is currently in **Early Beta (`v3.0.0-beta`)**, please note the following incomplete areas:
1. **Standalone `.hwpx` Binary Image Injection**: While `.hwp` (v5.0 Binary) and `.docx` fully embed binary image streams (`BinData/BIN0001` and `word/media/image1`) into `6x2` attachment cells, `.hwpx` projection currently outputs attachment captions and requires `<hp:pic>` + `BinData/` ZIP entry injection in a future update.
2. **Arbitrary Template Topology Generalization**: Dynamic slot discovery currently targets `2x2` section tables and `6x2` attachment tables used in standard Korean startup/R&D grant templates, and is being generalized to arbitrary $N \times M$ form templates.
3. **Math / Equation AST (`OMML` $\leftrightarrow$ `HWP EQN` $\leftrightarrow$ `LaTeX`)**: Equation nodes are currently preserved as text runs; full bidirectional equation AST compilation is planned for `v3.1`.
4. **Native `.xlsx` & `.html` Projectors**: `.csv`/`.tsv` and `.md` lifters/projectors are implemented; full styled `.xlsx` and `.html` DOM projectors are in progress.
