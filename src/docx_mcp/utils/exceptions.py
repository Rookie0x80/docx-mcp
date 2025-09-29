"""Custom exceptions for the docx table MCP server."""


class DocxTableMCPError(Exception):
    """Base exception for docx table MCP server."""
    pass


class DocumentNotFoundError(DocxTableMCPError):
    """Raised when a document file is not found."""
    pass


class TableNotFoundError(DocxTableMCPError):
    """Raised when a table is not found in the document."""
    pass


class InvalidTableIndexError(DocxTableMCPError):
    """Raised when an invalid table index is provided."""
    pass


class InvalidCellPositionError(DocxTableMCPError):
    """Raised when an invalid cell position is provided."""
    pass


class TableOperationError(DocxTableMCPError):
    """Raised when a table operation fails."""
    pass


class DataFormatError(DocxTableMCPError):
    """Raised when data format is invalid."""
    pass


class DocumentAccessError(DocxTableMCPError):
    """Raised when document cannot be accessed or opened."""
    pass


class FileSizeError(DocxTableMCPError):
    """Raised when file size exceeds limits."""
    pass
