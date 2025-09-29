"""Table-related data models."""

from dataclasses import dataclass
from typing import List, Optional, Any, Dict


@dataclass
class TableInfo:
    """Information about a table in a Word document."""
    index: int
    rows: int
    columns: int
    has_headers: bool
    style: Optional[str]
    position: int  # Position in the document


@dataclass
class CellPosition:
    """Position of a cell in a table."""
    table_index: int
    row_index: int
    column_index: int


@dataclass
class SearchResult:
    """Result of a search operation in a table."""
    positions: List[CellPosition]
    matches: List[str]
    total_matches: int


@dataclass
class CellFormatting:
    """Cell formatting options."""
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    font_size: Optional[int] = None
    font_name: Optional[str] = None
    alignment: Optional[str] = None  # left, center, right, justify
    background_color: Optional[str] = None
    text_color: Optional[str] = None


@dataclass
class TableData:
    """Container for table data."""
    data: List[List[str]]
    headers: Optional[List[str]] = None
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary format."""
        result = {"data": self.data}
        if self.headers:
            result["headers"] = self.headers
        return result
    
    def to_csv_format(self) -> List[List[str]]:
        """Convert to CSV format (headers + data)."""
        if self.headers:
            return [self.headers] + self.data
        return self.data
