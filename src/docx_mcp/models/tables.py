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
class TableSearchMatch:
    """A single search match in a table cell."""
    table_index: int
    row_index: int
    column_index: int
    cell_value: str
    match_text: str
    match_start: int  # Start position of match within cell
    match_end: int    # End position of match within cell


@dataclass
class TableSearchResult:
    """Result of a table search operation."""
    query: str
    search_mode: str  # "exact", "contains", "regex"
    case_sensitive: bool
    matches: List[TableSearchMatch]
    total_matches: int
    tables_searched: List[int]
    summary: Dict[str, Any]  # Summary statistics
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary format."""
        return {
            "query": self.query,
            "search_mode": self.search_mode,
            "case_sensitive": self.case_sensitive,
            "matches": [
                {
                    "table_index": m.table_index,
                    "row_index": m.row_index,
                    "column_index": m.column_index,
                    "cell_value": m.cell_value,
                    "match_text": m.match_text,
                    "match_start": m.match_start,
                    "match_end": m.match_end
                }
                for m in self.matches
            ],
            "total_matches": self.total_matches,
            "tables_searched": self.tables_searched,
            "summary": self.summary
        }


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
