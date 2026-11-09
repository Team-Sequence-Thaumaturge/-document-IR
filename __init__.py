"""
T-DIR v4.0.0 Isolated MCP Server Package.
Topological Document Intermediate Representation with 100.00% Text Immutability,
Dynamic File/Template Spec Extraction, User Spec Override Priority, 6-Stage Minimum-Invasion Elastic Solver,
3-Way Isomorphic Equation AST Compiler, Standalone OWPML (.hwpx) Synthesis,
8-Format N-to-M Codec Hub (.hwp, .hwpx, .docx, .md, .csv, .xlsx, .html, .pdf),
and Security/Side-Effect Hardening ([SEC-01] ~ [SEC-06]).
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
    parse_latex_to_equation_ast,
    parse_hwpeqn_to_equation_ast,
    parse_omml_to_equation_ast,
    compile_equation_ast_to_latex,
    compile_equation_ast_to_hwpeqn,
    compile_equation_ast_to_omml_element,
    compile_equation_ast_to_linear_text,
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
    "parse_latex_to_equation_ast",
    "parse_hwpeqn_to_equation_ast",
    "parse_omml_to_equation_ast",
    "compile_equation_ast_to_latex",
    "compile_equation_ast_to_hwpeqn",
    "compile_equation_ast_to_omml_element",
    "compile_equation_ast_to_linear_text",
    "ElasticLayoutState",
    "MinimumInvasionElasticSolver",
    "TdirBetaMcpServer",
]
__version__ = "4.0.0"
