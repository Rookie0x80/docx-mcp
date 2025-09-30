"""MCP Server for Word document operations using FastMCP."""

from typing import Any, Dict, List, Optional
from fastmcp import FastMCP

from .core.document_manager import DocumentManager
from .operations.tables.table_operations import TableOperations
from .models.responses import OperationResponse
from .models.formatting import TextFormat, CellAlignment, CellBorders, CellFormatting


# Initialize FastMCP app
mcp = FastMCP("Word Document MCP Server")

# Initialize managers
document_manager = DocumentManager()
table_operations = TableOperations(document_manager)


# Note: FastMCP automatically handles JSON parameter validation and conversion
# No need for explicit Pydantic request models when using direct parameter signatures


# Document operations
@mcp.tool()
def open_document(
    file_path: str,
    create_if_not_exists: bool = True
) -> Dict[str, Any]:
    """Open or create a Word document.
    
    Args:
        file_path: Path to the document file
        create_if_not_exists: Create document if it doesn't exist
    """
    result = document_manager.open_document(
        file_path,
        create_if_not_exists
    )
    return result.to_dict()


@mcp.tool()
def save_document(
    file_path: str,
    save_as: Optional[str] = None
) -> Dict[str, Any]:
    """Save a Word document.
    
    Args:
        file_path: Path to the document file
        save_as: Optional path to save as a different file
    """
    result = document_manager.save_document(
        file_path,
        save_as
    )
    return result.to_dict()


@mcp.tool()
def get_document_info(file_path: str) -> Dict[str, Any]:
    """Get information about a document.
    
    Args:
        file_path: Path to the document file
    """
    result = document_manager.get_document_info(file_path)
    return result.to_dict()


# Table structure operations
@mcp.tool()
def create_table(
    file_path: str,
    rows: int,
    cols: int,
    position: str = "end",
    paragraph_index: Optional[int] = None,
    headers: Optional[List[str]] = None
) -> Dict[str, Any]:
    """Create a new table in the document.
    
    Args:
        file_path: Path to the document file
        rows: Number of rows (must be >= 1)
        cols: Number of columns (must be >= 1)
        position: Position to insert table ("end", "beginning", "after_paragraph")
        paragraph_index: Paragraph index for after_paragraph position
        headers: Optional header row data
    """
    result = table_operations.create_table(
        file_path,
        rows,
        cols,
        position,
        paragraph_index,
        headers
    )
    return result.to_dict()


@mcp.tool()
def delete_table(
    file_path: str,
    table_index: int
) -> Dict[str, Any]:
    """Delete a table from the document.
    
    Args:
        file_path: Path to the document file
        table_index: Index of the table to delete (>= 0)
    """
    result = table_operations.delete_table(
        file_path,
        table_index
    )
    return result.to_dict()


@mcp.tool()
def add_table_rows(
    file_path: str,
    table_index: int,
    count: int = 1,
    position: str = "end",
    row_index: Optional[int] = None
) -> Dict[str, Any]:
    """Add rows to a table.
    
    Args:
        file_path: Path to the document file
        table_index: Index of the table (>= 0)
        count: Number of rows to add (>= 1)
        position: Position to add rows ("end", "beginning", "at_index")
        row_index: Row index for at_index position
    """
    result = table_operations.add_table_rows(
        file_path,
        table_index,
        count,
        position,
        row_index
    )
    return result.to_dict()


@mcp.tool()
def add_table_columns(
    file_path: str,
    table_index: int,
    count: int = 1,
    position: str = "end",
    column_index: Optional[int] = None
) -> Dict[str, Any]:
    """Add columns to a table.
    
    Args:
        file_path: Path to the document file
        table_index: Index of the table (>= 0)
        count: Number of columns to add (>= 1)
        position: Position to add columns ("end", "beginning", "at_index")
        column_index: Column index for at_index position
    """
    result = table_operations.add_table_columns(
        file_path,
        table_index,
        count,
        position,
        column_index
    )
    return result.to_dict()


@mcp.tool()
def delete_table_rows(
    file_path: str,
    table_index: int,
    row_indices: List[int]
) -> Dict[str, Any]:
    """Delete rows from a table.
    
    Args:
        file_path: Path to the document file
        table_index: Index of the table (>= 0)
        row_indices: List of row indices to delete
    """
    result = table_operations.delete_table_rows(
        file_path,
        table_index,
        row_indices
    )
    return result.to_dict()


# Data operations
@mcp.tool()
def set_cell_value(
    file_path: str,
    table_index: int,
    row_index: int,
    column_index: int,
    value: str
) -> Dict[str, Any]:
    """Set the value of a specific cell.
    
    Args:
        file_path: Path to the document file
        table_index: Index of the table (>= 0)
        row_index: Row index (>= 0)
        column_index: Column index (>= 0)
        value: Value to set
    """
    result = table_operations.set_cell_value(
        file_path,
        table_index,
        row_index,
        column_index,
        value
    )
    return result.to_dict()


@mcp.tool()
def get_cell_value(
    file_path: str,
    table_index: int,
    row_index: int,
    column_index: int
) -> Dict[str, Any]:
    """Get the value of a specific cell.
    
    Args:
        file_path: Path to the document file
        table_index: Index of the table (>= 0)
        row_index: Row index (>= 0)
        column_index: Column index (>= 0)
    """
    result = table_operations.get_cell_value(
        file_path,
        table_index,
        row_index,
        column_index
    )
    return result.to_dict()


@mcp.tool()
def get_table_data(
    file_path: str,
    table_index: int,
    include_headers: bool = True,
    format: str = "array"
) -> Dict[str, Any]:
    """Get all data from a table.
    
    Args:
        file_path: Path to the document file
        table_index: Index of the table (>= 0)
        include_headers: Whether to include headers
        format: Format of returned data ("array", "object", "csv")
    """
    result = table_operations.get_table_data(
        file_path,
        table_index,
        include_headers,
        format
    )
    return result.to_dict()


# Query operations
@mcp.tool()
def list_tables(
    file_path: str,
    include_summary: bool = True
) -> Dict[str, Any]:
    """List all tables in the document.
    
    Args:
        file_path: Path to the document file
        include_summary: Whether to include table summary information
    """
    result = table_operations.list_tables(
        file_path,
        include_summary
    )
    return result.to_dict()


# Table search operations
@mcp.tool()
def search_table_content(
    file_path: str,
    query: str,
    search_mode: str = "contains",
    case_sensitive: bool = False,
    table_indices: Optional[List[int]] = None,
    max_results: Optional[int] = None
) -> Dict[str, Any]:
    """Search for content within table cells across all or specified tables.
    
    Args:
        file_path: Path to the document file
        query: Search query string
        search_mode: Search mode ("exact", "contains", "regex")
        case_sensitive: Whether search is case sensitive
        table_indices: Optional list of table indices to search (None = all tables)
        max_results: Maximum number of results to return (None = no limit)
    """
    result = table_operations.search_table_content(
        file_path,
        query,
        search_mode,
        case_sensitive,
        table_indices,
        max_results
    )
    return result.to_dict()


@mcp.tool()
def search_table_headers(
    file_path: str,
    query: str,
    search_mode: str = "contains",
    case_sensitive: bool = False
) -> Dict[str, Any]:
    """Search specifically in table headers (first row of each table).
    
    Args:
        file_path: Path to the document file
        query: Search query string
        search_mode: Search mode ("exact", "contains", "regex")
        case_sensitive: Whether search is case sensitive
    """
    result = table_operations.search_table_headers(
        file_path,
        query,
        search_mode,
        case_sensitive
    )
    return result.to_dict()


# Cell formatting operations
@mcp.tool()
def format_cell_text(
    file_path: str,
    table_index: int,
    row_index: int,
    column_index: int,
    font_family: Optional[str] = None,
    font_size: Optional[int] = None,
    font_color: Optional[str] = None,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    underline: Optional[bool] = None,
    strikethrough: Optional[bool] = None
) -> Dict[str, Any]:
    """Format text in a specific cell.
    
    Args:
        file_path: Path to the document file
        table_index: Index of the table (>= 0)
        row_index: Row index (>= 0)
        column_index: Column index (>= 0)
        font_family: Font family name (e.g., "Arial", "Times New Roman")
        font_size: Font size in points (8-72)
        font_color: Font color as hex string (e.g., "FF0000" for red)
        bold: Make text bold
        italic: Make text italic
        underline: Make text underlined
        strikethrough: Make text strikethrough
    """
    text_format = TextFormat(
        font_family=font_family,
        font_size=font_size,
        font_color=font_color,
        bold=bold,
        italic=italic,
        underline=underline,
        strikethrough=strikethrough
    )
    
    result = table_operations.formatting.format_cell_text(
        file_path, table_index, row_index, column_index, text_format
    )
    return result.to_dict()


@mcp.tool()
def format_cell_alignment(
    file_path: str,
    table_index: int,
    row_index: int,
    column_index: int,
    horizontal: Optional[str] = None,
    vertical: Optional[str] = None
) -> Dict[str, Any]:
    """Set text alignment for a specific cell.
    
    Args:
        file_path: Path to the document file
        table_index: Index of the table (>= 0)
        row_index: Row index (>= 0)
        column_index: Column index (>= 0)
        horizontal: Horizontal alignment ("left", "center", "right", "justify")
        vertical: Vertical alignment ("top", "middle", "bottom")
    """
    alignment_dict = {}
    if horizontal:
        alignment_dict["horizontal"] = horizontal
    if vertical:
        alignment_dict["vertical"] = vertical
    
    result = table_operations.formatting.format_cell_alignment(
        file_path, table_index, row_index, column_index, alignment_dict
    )
    return result.to_dict()


@mcp.tool()
def format_cell_background(
    file_path: str,
    table_index: int,
    row_index: int,
    column_index: int,
    color: str
) -> Dict[str, Any]:
    """Set background color for a specific cell.
    
    Args:
        file_path: Path to the document file
        table_index: Index of the table (>= 0)
        row_index: Row index (>= 0)
        column_index: Column index (>= 0)
        color: Background color as hex string (e.g., "FFFF00" for yellow)
    """
    result = table_operations.formatting.format_cell_background(
        file_path, table_index, row_index, column_index, color
    )
    return result.to_dict()


@mcp.tool()
def format_cell_borders(
    file_path: str,
    table_index: int,
    row_index: int,
    column_index: int,
    top_style: Optional[str] = None,
    top_width: Optional[str] = None,
    top_color: Optional[str] = None,
    bottom_style: Optional[str] = None,
    bottom_width: Optional[str] = None,
    bottom_color: Optional[str] = None,
    left_style: Optional[str] = None,
    left_width: Optional[str] = None,
    left_color: Optional[str] = None,
    right_style: Optional[str] = None,
    right_width: Optional[str] = None,
    right_color: Optional[str] = None
) -> Dict[str, Any]:
    """Set borders for a specific cell.
    
    Args:
        file_path: Path to the document file
        table_index: Index of the table (>= 0)
        row_index: Row index (>= 0)
        column_index: Column index (>= 0)
        top_style: Top border style ("solid", "dashed", "dotted", "double", "none")
        top_width: Top border width ("thin", "medium", "thick")
        top_color: Top border color as hex string
        bottom_style: Bottom border style
        bottom_width: Bottom border width
        bottom_color: Bottom border color as hex string
        left_style: Left border style
        left_width: Left border width
        left_color: Left border color as hex string
        right_style: Right border style
        right_width: Right border width
        right_color: Right border color as hex string
    """
    borders_dict = {}
    
    # Build border configuration
    if any([top_style, top_width, top_color]):
        borders_dict["top"] = {
            "style": top_style or "solid",
            "width": top_width or "thin",
            "color": top_color or "000000"
        }
    
    if any([bottom_style, bottom_width, bottom_color]):
        borders_dict["bottom"] = {
            "style": bottom_style or "solid",
            "width": bottom_width or "thin",
            "color": bottom_color or "000000"
        }
    
    if any([left_style, left_width, left_color]):
        borders_dict["left"] = {
            "style": left_style or "solid",
            "width": left_width or "thin",
            "color": left_color or "000000"
        }
    
    if any([right_style, right_width, right_color]):
        borders_dict["right"] = {
            "style": right_style or "solid",
            "width": right_width or "thin",
            "color": right_color or "000000"
        }
    
    result = table_operations.formatting.format_cell_borders(
        file_path, table_index, row_index, column_index, borders_dict
    )
    return result.to_dict()


# Table analysis operations
@mcp.tool()
def analyze_table_structure(
    file_path: str,
    table_index: int,
    include_cell_details: bool = True
) -> Dict[str, Any]:
    """Analyze the complete structure and styling of a specific table.
    
    This function provides comprehensive analysis of table structure including:
    - Cell merging information (which cells are merged horizontally/vertically)
    - Text formatting for each cell (font, size, color, bold, italic, etc.)
    - Cell alignment settings (horizontal and vertical alignment)
    - Background colors and border styles
    - Table-level properties and consistency analysis
    
    This is particularly useful for AI models to understand the existing table
    structure and styling before making modifications, ensuring they maintain
    the original formatting and don't break the table layout.
    
    Args:
        file_path: Path to the document file
        table_index: Index of the table to analyze (>= 0)
        include_cell_details: Whether to include detailed cell-level formatting analysis
    """
    result = table_operations.analyze_table_structure(
        file_path,
        table_index,
        include_cell_details
    )
    return result.to_dict()


@mcp.tool()
def analyze_all_tables_structure(
    file_path: str,
    include_cell_details: bool = True
) -> Dict[str, Any]:
    """Analyze the structure and styling of all tables in the document.
    
    This function provides comprehensive analysis of all tables in the document,
    including the same detailed information as analyze_table_structure but for
    every table. This gives AI models a complete understanding of the document's
    table structure and formatting patterns.
    
    The analysis includes:
    - Complete table inventory with structure details
    - Cell merging patterns across all tables
    - Formatting consistency analysis
    - Style summary (unique fonts, colors, etc. used)
    - Header detection and analysis
    
    Args:
        file_path: Path to the document file
        include_cell_details: Whether to include detailed cell-level formatting analysis
    """
    result = table_operations.analyze_all_tables(
        file_path,
        include_cell_details
    )
    return result.to_dict()


def main():
    """Main entry point to run the MCP server."""
    import argparse
    
    parser = argparse.ArgumentParser(description="Word Document MCP Server")
    parser.add_argument(
        "--transport", 
        choices=["stdio", "sse", "streamable-http"],
        default="stdio",
        help="Transport protocol to use (default: stdio)"
    )
    parser.add_argument(
        "--host",
        default="localhost",
        help="Host to bind to for HTTP/SSE transports (default: localhost)"
    )
    parser.add_argument(
        "--port",
        type=int,
        default=8000,
        help="Port to bind to for HTTP/SSE transports (default: 8000)"
    )
    parser.add_argument(
        "--no-banner",
        action="store_true",
        help="Disable startup banner"
    )
    
    args = parser.parse_args()
    
    # Prepare transport kwargs for HTTP/SSE
    transport_kwargs = {}
    if args.transport in ["sse", "streamable-http"]:
        transport_kwargs["host"] = args.host
        transport_kwargs["port"] = args.port
    
    print(f"Starting Word Document MCP Server with {args.transport} transport...")
    if args.transport in ["sse", "streamable-http"]:
        print(f"Server will be available at http://{args.host}:{args.port}")
    
    mcp.run(
        transport=args.transport,
        show_banner=not args.no_banner,
        **transport_kwargs
    )


if __name__ == "__main__":
    main()
