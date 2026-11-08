"""
T-DIR v3.0 Beta Isolated MCP Server Package.
Topological Document Intermediate Representation with 100.00% Text Immutability,
Dynamic File/Template Spec Extraction, User Spec Override Priority, and 5-Stage Minimum-Invasion Elastic Solver.
"""

from .ir_schema import (
    TextRunIR,
    ParagraphIR,
    TableCellIR,
    TableGridIR,
    MediaObjectIR,
    SectionIR,
    SlotSpecIR,
    SourceDocumentSpecIR,
    TargetDocumentSpecIR,
    UserSpecOverrideIR,
    ConversionJobIR,
    resolve_effective_document_spec,
    compute_non_ws_sha256,
    validate_text_immutability,
)
from .elastic_solver import ElasticLayoutState, MinimumInvasionElasticSolver
from .mcp_server import TdirBetaMcpServer

__all__ = [
    "TextRunIR",
    "ParagraphIR",
    "TableCellIR",
    "TableGridIR",
    "MediaObjectIR",
    "SectionIR",
    "SlotSpecIR",
    "SourceDocumentSpecIR",
    "TargetDocumentSpecIR",
    "UserSpecOverrideIR",
    "ConversionJobIR",
    "resolve_effective_document_spec",
    "compute_non_ws_sha256",
    "validate_text_immutability",
    "ElasticLayoutState",
    "MinimumInvasionElasticSolver",
    "TdirBetaMcpServer",
]
__version__ = "3.1.0-beta"
