"""
SADT (Stella Actuation Document Transfer) v4.1.0 MCP Server Package.
Powered by Topological Document Intermediate Representation (T-DIR) for structure-preserving
conversion across HWP v5.0 (.hwp), OWPML (.hwpx), OOXML (.docx / .xlsx), HTML5, PDF 1.4, Markdown, and CSV,
with non-whitespace SHA-256 text verification, dynamic file/template spec extraction,
user spec override priority, a 6-stage Minimum-Invasion Elastic Solver,
a 3-Way Equation AST converter, standalone .hwpx package synthesis,
and defensive parsing/validation checks ([SEC-01] ~ [SEC-06]).
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
from .mcp_server import TdirBetaMcpServer, SadtMcpServer

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
    "SadtMcpServer",
]
__version__ = "4.1.0"
