"""Table operations for Word documents."""

from typing import List, Optional, Dict, Any, Union
from docx import Document
from docx.table import Table, _Cell
from docx.shared import Inches

from ...models.responses import OperationResponse
from ...models.tables import TableInfo, CellPosition, SearchResult, TableData
from ...utils.exceptions import (
    TableNotFoundError,
    InvalidTableIndexError,
    InvalidCellPositionError,
    TableOperationError,
    DataFormatError,
)
from ...utils.validation import (
    validate_table_index,
    validate_cell_position,
    validate_table_data,
    validate_position_parameter,
    sanitize_string,
)
from ...core.document_manager import DocumentManager


class TableOperations:
    """Handles table operations in Word documents."""
    
    def __init__(self, document_manager: DocumentManager):
        """
        Initialize table operations.
        
        Args:
            document_manager: Document manager instance
        """
        self.document_manager = document_manager
    
    def create_table(
        self,
        file_path: str,
        rows: int,
        cols: int,
        position: str = "end",
        paragraph_index: Optional[int] = None,
        headers: Optional[List[str]] = None,
    ) -> OperationResponse:
        """
        Create a new table in the document.
        
        Args:
            file_path: Path to the document
            rows: Number of rows
            cols: Number of columns
            position: Where to insert the table
            paragraph_index: Paragraph index for 'after_paragraph' position
            headers: Optional header row data
            
        Returns:
            OperationResponse with operation result
        """
        try:
            # Validate inputs
            if rows <= 0 or cols <= 0:
                return OperationResponse.error("Rows and columns must be positive integers")
            
            valid_positions = ["end", "beginning", "after_paragraph"]
            validate_position_parameter(position, valid_positions)
            
            if position == "after_paragraph" and paragraph_index is None:
                return OperationResponse.error("paragraph_index required for 'after_paragraph' position")
            
            if headers and len(headers) != cols:
                return OperationResponse.error(f"Headers length ({len(headers)}) must match columns ({cols})")
            
            # Get document
            document = self.document_manager.get_document(file_path)
            if not document:
                return OperationResponse.error(f"Document not loaded: {file_path}")
            
            # Create table
            table = None
            if position == "end":
                table = document.add_table(rows=rows, cols=cols)
            elif position == "beginning":
                # Insert at beginning by adding after title or first paragraph
                if document.paragraphs:
                    p = document.paragraphs[0]
                    table = p.insert_paragraph_before().add_table(rows=rows, cols=cols)
                else:
                    table = document.add_table(rows=rows, cols=cols)
            elif position == "after_paragraph":
                if paragraph_index < 0 or paragraph_index >= len(document.paragraphs):
                    return OperationResponse.error(f"Invalid paragraph index: {paragraph_index}")
                p = document.paragraphs[paragraph_index]
                new_p = p.insert_paragraph_after()
                table = new_p._element.addnext(document.add_table(rows=rows, cols=cols)._element)
                table = document.tables[-1]  # Get the newly added table
            
            if not table:
                return OperationResponse.error("Failed to create table")
            
            # Set headers if provided
            if headers:
                for col_idx, header in enumerate(headers):
                    table.cell(0, col_idx).text = sanitize_string(header)
            
            table_index = len(document.tables) - 1
            
            data = {
                "table_index": table_index,
                "rows": rows,
                "cols": cols,
                "position": position,
                "has_headers": bool(headers)
            }
            
            return OperationResponse.success(f"Table created with {rows} rows and {cols} columns", data)
            
        except Exception as e:
            return OperationResponse.error(f"Failed to create table: {str(e)}")
    
    def delete_table(self, file_path: str, table_index: int) -> OperationResponse:
        """
        Delete a table from the document.
        
        Args:
            file_path: Path to the document
            table_index: Index of the table to delete
            
        Returns:
            OperationResponse with operation result
        """
        try:
            document = self.document_manager.get_document(file_path)
            if not document:
                return OperationResponse.error(f"Document not loaded: {file_path}")
            
            validate_table_index(table_index, len(document.tables))
            
            # Get table and remove it
            table = document.tables[table_index]
            table._element.getparent().remove(table._element)
            
            return OperationResponse.success(f"Table {table_index} deleted")
            
        except (InvalidTableIndexError, TableNotFoundError) as e:
            return OperationResponse.error(str(e))
        except Exception as e:
            return OperationResponse.error(f"Failed to delete table: {str(e)}")
    
    def add_table_rows(
        self,
        file_path: str,
        table_index: int,
        count: int = 1,
        position: str = "end",
        row_index: Optional[int] = None,
    ) -> OperationResponse:
        """
        Add rows to a table.
        
        Args:
            file_path: Path to the document
            table_index: Index of the table
            count: Number of rows to add
            position: Where to add rows
            row_index: Specific row index for 'at_index' position
            
        Returns:
            OperationResponse with operation result
        """
        try:
            if count <= 0:
                return OperationResponse.error("Count must be a positive integer")
            
            valid_positions = ["end", "beginning", "at_index"]
            validate_position_parameter(position, valid_positions)
            
            document = self.document_manager.get_document(file_path)
            if not document:
                return OperationResponse.error(f"Document not loaded: {file_path}")
            
            validate_table_index(table_index, len(document.tables))
            table = document.tables[table_index]
            
            if position == "at_index" and row_index is None:
                return OperationResponse.error("row_index required for 'at_index' position")
            
            if position == "at_index":
                validate_cell_position(row_index, 0, len(table.rows), len(table.columns))
            
            # Add rows
            for _ in range(count):
                if position == "end":
                    table.add_row()
                elif position == "beginning":
                    # Insert at beginning - not directly supported, need to work around
                    new_row = table.add_row()
                    # Move new row to beginning
                    table._element.insert(0, new_row._element)
                elif position == "at_index":
                    # Insert at specific index
                    new_row = table.add_row()
                    # Move to desired position
                    if row_index < len(table.rows) - 1:
                        target_row = table.rows[row_index]
                        target_row._element.addprevious(new_row._element)
            
            data = {
                "table_index": table_index,
                "rows_added": count,
                "new_row_count": len(table.rows),
                "position": position
            }
            
            return OperationResponse.success(f"Added {count} rows to table {table_index}", data)
            
        except (InvalidTableIndexError, InvalidCellPositionError) as e:
            return OperationResponse.error(str(e))
        except Exception as e:
            return OperationResponse.error(f"Failed to add rows: {str(e)}")
    
    def add_table_columns(
        self,
        file_path: str,
        table_index: int,
        count: int = 1,
        position: str = "end",
        column_index: Optional[int] = None,
    ) -> OperationResponse:
        """
        Add columns to a table.
        
        Args:
            file_path: Path to the document
            table_index: Index of the table
            count: Number of columns to add
            position: Where to add columns
            column_index: Specific column index for 'at_index' position
            
        Returns:
            OperationResponse with operation result
        """
        try:
            if count <= 0:
                return OperationResponse.error("Count must be a positive integer")
            
            valid_positions = ["end", "beginning", "at_index"]
            validate_position_parameter(position, valid_positions)
            
            document = self.document_manager.get_document(file_path)
            if not document:
                return OperationResponse.error(f"Document not loaded: {file_path}")
            
            validate_table_index(table_index, len(document.tables))
            table = document.tables[table_index]
            
            if not table.rows:
                return OperationResponse.error("Cannot add columns to empty table")
            
            original_cols = len(table.columns)
            
            if position == "at_index" and column_index is None:
                return OperationResponse.error("column_index required for 'at_index' position")
            
            if position == "at_index":
                validate_cell_position(0, column_index, len(table.rows), original_cols)
            
            # Add columns by adding cells to each row
            for _ in range(count):
                for row in table.rows:
                    if position == "end":
                        row.cells[-1]._element.addnext(row.cells[0]._element.__class__())
                    elif position == "beginning":
                        row.cells[0]._element.addprevious(row.cells[0]._element.__class__())
                    elif position == "at_index":
                        target_cell = row.cells[column_index]
                        target_cell._element.addprevious(row.cells[0]._element.__class__())
            
            new_cols = len(table.columns)
            
            data = {
                "table_index": table_index,
                "columns_added": count,
                "new_column_count": new_cols,
                "position": position
            }
            
            return OperationResponse.success(f"Added {count} columns to table {table_index}", data)
            
        except (InvalidTableIndexError, InvalidCellPositionError) as e:
            return OperationResponse.error(str(e))
        except Exception as e:
            return OperationResponse.error(f"Failed to add columns: {str(e)}")
    
    def delete_table_rows(
        self, file_path: str, table_index: int, row_indices: List[int]
    ) -> OperationResponse:
        """
        Delete rows from a table.
        
        Args:
            file_path: Path to the document
            table_index: Index of the table
            row_indices: List of row indices to delete
            
        Returns:
            OperationResponse with operation result
        """
        try:
            if not row_indices:
                return OperationResponse.error("No row indices provided")
            
            document = self.document_manager.get_document(file_path)
            if not document:
                return OperationResponse.error(f"Document not loaded: {file_path}")
            
            validate_table_index(table_index, len(document.tables))
            table = document.tables[table_index]
            
            # Validate all row indices
            for row_idx in row_indices:
                validate_cell_position(row_idx, 0, len(table.rows), len(table.columns))
            
            # Sort indices in reverse order to delete from end to beginning
            sorted_indices = sorted(set(row_indices), reverse=True)
            
            # Delete rows
            for row_idx in sorted_indices:
                row = table.rows[row_idx]
                row._element.getparent().remove(row._element)
            
            data = {
                "table_index": table_index,
                "rows_deleted": len(sorted_indices),
                "remaining_rows": len(table.rows)
            }
            
            return OperationResponse.success(
                f"Deleted {len(sorted_indices)} rows from table {table_index}", data
            )
            
        except (InvalidTableIndexError, InvalidCellPositionError) as e:
            return OperationResponse.error(str(e))
        except Exception as e:
            return OperationResponse.error(f"Failed to delete rows: {str(e)}")
    
    def set_cell_value(
        self, file_path: str, table_index: int, row_index: int, column_index: int, value: str
    ) -> OperationResponse:
        """
        Set the value of a specific cell.
        
        Args:
            file_path: Path to the document
            table_index: Index of the table
            row_index: Row index
            column_index: Column index
            value: Value to set
            
        Returns:
            OperationResponse with operation result
        """
        try:
            document = self.document_manager.get_document(file_path)
            if not document:
                return OperationResponse.error(f"Document not loaded: {file_path}")
            
            validate_table_index(table_index, len(document.tables))
            table = document.tables[table_index]
            
            validate_cell_position(row_index, column_index, len(table.rows), len(table.columns))
            
            # Set cell value
            cell = table.cell(row_index, column_index)
            cell.text = sanitize_string(value)
            
            data = {
                "table_index": table_index,
                "row_index": row_index,
                "column_index": column_index,
                "value": cell.text
            }
            
            return OperationResponse.success(
                f"Cell value set at table {table_index}, row {row_index}, column {column_index}",
                data
            )
            
        except (InvalidTableIndexError, InvalidCellPositionError) as e:
            return OperationResponse.error(str(e))
        except Exception as e:
            return OperationResponse.error(f"Failed to set cell value: {str(e)}")
    
    def get_cell_value(
        self, file_path: str, table_index: int, row_index: int, column_index: int
    ) -> OperationResponse:
        """
        Get the value of a specific cell.
        
        Args:
            file_path: Path to the document
            table_index: Index of the table
            row_index: Row index
            column_index: Column index
            
        Returns:
            OperationResponse with cell value
        """
        try:
            document = self.document_manager.get_document(file_path)
            if not document:
                return OperationResponse.error(f"Document not loaded: {file_path}")
            
            validate_table_index(table_index, len(document.tables))
            table = document.tables[table_index]
            
            validate_cell_position(row_index, column_index, len(table.rows), len(table.columns))
            
            # Get cell value
            cell = table.cell(row_index, column_index)
            value = cell.text
            
            data = {
                "table_index": table_index,
                "row_index": row_index,
                "column_index": column_index,
                "value": value
            }
            
            return OperationResponse.success("Cell value retrieved", data)
            
        except (InvalidTableIndexError, InvalidCellPositionError) as e:
            return OperationResponse.error(str(e))
        except Exception as e:
            return OperationResponse.error(f"Failed to get cell value: {str(e)}")
    
    def get_table_data(
        self,
        file_path: str,
        table_index: int,
        include_headers: bool = True,
        format_type: str = "array",
    ) -> OperationResponse:
        """
        Get all data from a table.
        
        Args:
            file_path: Path to the document
            table_index: Index of the table
            include_headers: Whether to include headers
            format_type: Format of returned data ('array', 'object', 'csv')
            
        Returns:
            OperationResponse with table data
        """
        try:
            valid_formats = ["array", "object", "csv"]
            if format_type not in valid_formats:
                return OperationResponse.error(f"Invalid format. Valid options: {', '.join(valid_formats)}")
            
            document = self.document_manager.get_document(file_path)
            if not document:
                return OperationResponse.error(f"Document not loaded: {file_path}")
            
            validate_table_index(table_index, len(document.tables))
            table = document.tables[table_index]
            
            if not table.rows:
                return OperationResponse.success("Table is empty", {"data": []})
            
            # Extract data
            data = []
            headers = None
            
            start_row = 0
            if include_headers and table.rows:
                # Extract headers from first row
                headers = [cell.text for cell in table.rows[0].cells]
                start_row = 1
            
            # Extract data rows
            for row in table.rows[start_row:]:
                row_data = [cell.text for cell in row.cells]
                data.append(row_data)
            
            # Format data according to requested format
            if format_type == "array":
                result_data = data
                if include_headers and headers:
                    result_data = [headers] + data
            elif format_type == "object":
                if headers:
                    result_data = []
                    for row in data:
                        row_dict = {}
                        for i, value in enumerate(row):
                            header = headers[i] if i < len(headers) else f"Column_{i}"
                            row_dict[header] = value
                        result_data.append(row_dict)
                else:
                    result_data = [{"Column_" + str(i): value for i, value in enumerate(row)} for row in data]
            elif format_type == "csv":
                result_data = []
                if include_headers and headers:
                    result_data.append(headers)
                result_data.extend(data)
            
            response_data = {
                "table_index": table_index,
                "format": format_type,
                "rows": len(data),
                "columns": len(data[0]) if data else 0,
                "has_headers": bool(headers),
                "headers": headers,
                "data": result_data
            }
            
            return OperationResponse.success(f"Table data retrieved in {format_type} format", response_data)
            
        except (InvalidTableIndexError,) as e:
            return OperationResponse.error(str(e))
        except Exception as e:
            return OperationResponse.error(f"Failed to get table data: {str(e)}")
    
    def list_tables(self, file_path: str, include_summary: bool = True) -> OperationResponse:
        """
        List all tables in the document.
        
        Args:
            file_path: Path to the document
            include_summary: Whether to include table summary information
            
        Returns:
            OperationResponse with list of tables
        """
        try:
            document = self.document_manager.get_document(file_path)
            if not document:
                return OperationResponse.error(f"Document not loaded: {file_path}")
            
            tables = []
            for i, table in enumerate(document.tables):
                table_info = {
                    "index": i,
                    "rows": len(table.rows),
                    "columns": len(table.columns) if table.rows else 0,
                }
                
                if include_summary:
                    # Check if has headers (simple heuristic)
                    has_headers = False
                    if table.rows:
                        first_row_has_text = all(cell.text.strip() for cell in table.rows[0].cells)
                        has_headers = first_row_has_text
                    
                    table_info.update({
                        "has_headers": has_headers,
                        "style": getattr(table.style, 'name', None) if table.style else None,
                        "first_row_data": [cell.text for cell in table.rows[0].cells] if table.rows else []
                    })
                
                tables.append(table_info)
            
            data = {
                "tables": tables,
                "total_count": len(tables)
            }
            
            return OperationResponse.success(f"Found {len(tables)} tables", data)
            
        except Exception as e:
            return OperationResponse.error(f"Failed to list tables: {str(e)}")
