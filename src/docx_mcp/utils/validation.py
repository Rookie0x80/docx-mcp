"""Validation utilities for the docx table MCP server."""

import os
from pathlib import Path
from typing import List, Any, Optional

from .exceptions import (
    DocumentNotFoundError,
    InvalidTableIndexError,
    InvalidCellPositionError,
    DataFormatError,
    FileSizeError,
)


def validate_file_path(file_path: str, must_exist: bool = True, max_size_mb: int = 50) -> Path:
    """
    Validate file path and return Path object.
    
    Args:
        file_path: Path to the file
        must_exist: Whether the file must exist
        max_size_mb: Maximum file size in MB
        
    Returns:
        Path object
        
    Raises:
        DocumentNotFoundError: If file doesn't exist when required
        FileSizeError: If file is too large
    """
    path = Path(file_path)
    
    # Check if file exists when required
    if must_exist and not path.exists():
        raise DocumentNotFoundError(f"Document not found: {file_path}")
    
    # Check file size if file exists
    if path.exists():
        file_size_mb = path.stat().st_size / (1024 * 1024)
        if file_size_mb > max_size_mb:
            raise FileSizeError(f"File size ({file_size_mb:.1f}MB) exceeds limit ({max_size_mb}MB)")
    
    # Ensure parent directory exists for new files
    if not must_exist:
        path.parent.mkdir(parents=True, exist_ok=True)
    
    return path


def validate_table_index(table_index: int, max_tables: Optional[int] = None) -> None:
    """
    Validate table index.
    
    Args:
        table_index: Index of the table
        max_tables: Maximum number of tables (if known)
        
    Raises:
        InvalidTableIndexError: If table index is invalid
    """
    if table_index < 0:
        raise InvalidTableIndexError(f"Table index must be non-negative, got: {table_index}")
    
    if max_tables is not None and table_index >= max_tables:
        raise InvalidTableIndexError(
            f"Table index {table_index} out of range. Document has {max_tables} tables."
        )


def validate_cell_position(
    row_index: int, 
    column_index: int, 
    max_rows: Optional[int] = None, 
    max_columns: Optional[int] = None
) -> None:
    """
    Validate cell position.
    
    Args:
        row_index: Row index
        column_index: Column index
        max_rows: Maximum number of rows (if known)
        max_columns: Maximum number of columns (if known)
        
    Raises:
        InvalidCellPositionError: If cell position is invalid
    """
    if row_index < 0:
        raise InvalidCellPositionError(f"Row index must be non-negative, got: {row_index}")
    
    if column_index < 0:
        raise InvalidCellPositionError(f"Column index must be non-negative, got: {column_index}")
    
    if max_rows is not None and row_index >= max_rows:
        raise InvalidCellPositionError(
            f"Row index {row_index} out of range. Table has {max_rows} rows."
        )
    
    if max_columns is not None and column_index >= max_columns:
        raise InvalidCellPositionError(
            f"Column index {column_index} out of range. Table has {max_columns} columns."
        )


def validate_table_data(data: List[List[str]]) -> None:
    """
    Validate table data format.
    
    Args:
        data: Table data as 2D list
        
    Raises:
        DataFormatError: If data format is invalid
    """
    if not isinstance(data, list):
        raise DataFormatError("Table data must be a list")
    
    if not data:
        raise DataFormatError("Table data cannot be empty")
    
    # Check that all rows are lists
    for i, row in enumerate(data):
        if not isinstance(row, list):
            raise DataFormatError(f"Row {i} must be a list, got: {type(row)}")
    
    # Check that all rows have the same length
    first_row_length = len(data[0])
    for i, row in enumerate(data[1:], 1):
        if len(row) != first_row_length:
            raise DataFormatError(
                f"All rows must have the same length. Row 0 has {first_row_length} "
                f"columns, but row {i} has {len(row)} columns."
            )
    
    # Convert all values to strings and validate
    for i, row in enumerate(data):
        for j, cell in enumerate(row):
            if not isinstance(cell, (str, int, float, bool, type(None))):
                raise DataFormatError(
                    f"Cell value at row {i}, column {j} must be a basic type "
                    f"(str, int, float, bool, or None), got: {type(cell)}"
                )


def validate_position_parameter(position: str, valid_positions: List[str]) -> None:
    """
    Validate position parameter.
    
    Args:
        position: Position value
        valid_positions: List of valid positions
        
    Raises:
        DataFormatError: If position is invalid
    """
    if position not in valid_positions:
        raise DataFormatError(
            f"Invalid position '{position}'. Valid options: {', '.join(valid_positions)}"
        )


def sanitize_string(value: Any) -> str:
    """
    Sanitize and convert value to string.
    
    Args:
        value: Value to convert
        
    Returns:
        Sanitized string
    """
    if value is None:
        return ""
    
    # Convert to string
    str_value = str(value)
    
    # Remove any control characters except newlines and tabs
    sanitized = "".join(char for char in str_value if ord(char) >= 32 or char in "\n\t")
    
    return sanitized
