"""MCP Server for Word document operations using FastMCP."""

from typing import Any, Dict, List, Optional
from fastmcp import FastMCP

from .core.document_manager import DocumentManager
from .operations.tables.table_operations import TableOperations
from .models.responses import OperationResponse


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
