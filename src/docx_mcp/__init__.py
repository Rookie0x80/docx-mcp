"""
DOCX-MCP: Word Document Operations MCP Server

A comprehensive Model Context Protocol (MCP) server for Microsoft Word document 
operations, with advanced table manipulation capabilities and planned expansion 
into full document processing features.

Current Focus: Complete table operations (formatting, import/export, search)
Future Scope: Full document content management and media integration
"""

__version__ = "2.0.0"
__author__ = "DOCX-MCP Team"
__email__ = "contact@docx-mcp.org"
__description__ = "Advanced Word document operations via Model Context Protocol"
__url__ = "https://github.com/your-org/docx-mcp"

# Feature flags for development phases
FEATURES = {
    "core_tables": True,          # Phase 1 - Complete
    "table_formatting": False,    # Phase 2 - In development
    "data_import_export": False,  # Phase 2 - Planned
    "table_search": False,        # Phase 2 - Planned
    "document_content": False,    # Phase 4 - Future
    "media_integration": False,   # Phase 5 - Future
    "enterprise_features": False  # Phase 6 - Future
}
