"""Data models for comprehensive table structure and style analysis."""

from dataclasses import dataclass
from typing import List, Optional, Dict, Any, Tuple
from enum import Enum


class CellMergeType(Enum):
    """Types of cell merging."""
    NONE = "none"
    HORIZONTAL = "horizontal"  # Cell spans multiple columns
    VERTICAL = "vertical"      # Cell spans multiple rows
    BOTH = "both"             # Cell spans both rows and columns


@dataclass
class MergeInfo:
    """Information about cell merging."""
    merge_type: CellMergeType
    start_row: int
    end_row: int
    start_col: int
    end_col: int
    span_rows: int
    span_cols: int


@dataclass
class CellStyleAnalysis:
    """Comprehensive analysis of a single cell's styling."""
    # Position information
    row_index: int
    column_index: int
    
    # Content
    text_content: str
    is_empty: bool
    
    # Merge information
    merge_info: Optional[MergeInfo]
    
    # Text formatting
    font_family: Optional[str]
    font_size: Optional[int]
    font_color: Optional[str]
    is_bold: bool
    is_italic: bool
    is_underlined: bool
    is_strikethrough: bool
    
    # Alignment
    horizontal_alignment: Optional[str]  # left, center, right, justify
    vertical_alignment: Optional[str]    # top, middle, bottom
    
    # Background and borders
    background_color: Optional[str]
    
    # Border information for each side
    top_border: Optional[Dict[str, str]]     # style, width, color
    bottom_border: Optional[Dict[str, str]]
    left_border: Optional[Dict[str, str]]
    right_border: Optional[Dict[str, str]]
    
    # Cell dimensions (if available)
    width: Optional[float]  # In points or inches
    height: Optional[float]
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary representation."""
        result = {
            "position": {
                "row": self.row_index,
                "column": self.column_index
            },
            "content": {
                "text": self.text_content,
                "is_empty": self.is_empty
            },
            "text_format": {
                "font_family": self.font_family,
                "font_size": self.font_size,
                "font_color": self.font_color,
                "bold": self.is_bold,
                "italic": self.is_italic,
                "underlined": self.is_underlined,
                "strikethrough": self.is_strikethrough
            },
            "alignment": {
                "horizontal": self.horizontal_alignment,
                "vertical": self.vertical_alignment
            },
            "background": {
                "color": self.background_color
            },
            "borders": {
                "top": self.top_border,
                "bottom": self.bottom_border,
                "left": self.left_border,
                "right": self.right_border
            }
        }
        
        if self.merge_info:
            result["merge"] = {
                "type": self.merge_info.merge_type.value,
                "start_row": self.merge_info.start_row,
                "end_row": self.merge_info.end_row,
                "start_col": self.merge_info.start_col,
                "end_col": self.merge_info.end_col,
                "span_rows": self.merge_info.span_rows,
                "span_cols": self.merge_info.span_cols
            }
        else:
            result["merge"] = None
            
        if self.width is not None or self.height is not None:
            result["dimensions"] = {
                "width": self.width,
                "height": self.height
            }
            
        return result


@dataclass
class TableStructureAnalysis:
    """Comprehensive analysis of table structure and styling."""
    # Basic table information
    table_index: int
    total_rows: int
    total_columns: int
    
    # Table-level properties
    table_style_name: Optional[str]
    table_alignment: Optional[str]  # left, center, right
    table_width: Optional[float]
    
    # Header information
    has_header_row: bool
    header_row_index: Optional[int]
    header_cells: Optional[List[str]]
    
    # All cell analyses
    cells: List[List[CellStyleAnalysis]]  # [row][column]
    
    # Merge summary
    merged_cells_count: int
    merge_regions: List[MergeInfo]
    
    # Style consistency analysis
    consistent_fonts: bool
    consistent_alignment: bool
    consistent_borders: bool
    
    # Common styles found
    unique_font_families: List[str]
    unique_font_sizes: List[int]
    unique_colors: List[str]
    unique_background_colors: List[str]
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary representation."""
        return {
            "table_info": {
                "index": self.table_index,
                "rows": self.total_rows,
                "columns": self.total_columns,
                "style_name": self.table_style_name,
                "alignment": self.table_alignment,
                "width": self.table_width
            },
            "header_info": {
                "has_header": self.has_header_row,
                "header_row_index": self.header_row_index,
                "header_cells": self.header_cells
            },
            "cells": [
                [cell.to_dict() for cell in row]
                for row in self.cells
            ],
            "merge_analysis": {
                "merged_cells_count": self.merged_cells_count,
                "merge_regions": [
                    {
                        "type": merge.merge_type.value,
                        "start_row": merge.start_row,
                        "end_row": merge.end_row,
                        "start_col": merge.start_col,
                        "end_col": merge.end_col,
                        "span_rows": merge.span_rows,
                        "span_cols": merge.span_cols
                    }
                    for merge in self.merge_regions
                ]
            },
            "style_consistency": {
                "fonts": self.consistent_fonts,
                "alignment": self.consistent_alignment,
                "borders": self.consistent_borders
            },
            "style_summary": {
                "font_families": self.unique_font_families,
                "font_sizes": self.unique_font_sizes,
                "colors": self.unique_colors,
                "background_colors": self.unique_background_colors
            }
        }


@dataclass
class TableAnalysisResult:
    """Result of comprehensive table analysis."""
    file_path: str
    total_tables: int
    analysis_timestamp: str
    tables: List[TableStructureAnalysis]
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary representation."""
        return {
            "file_info": {
                "path": self.file_path,
                "total_tables": self.total_tables,
                "analysis_timestamp": self.analysis_timestamp
            },
            "tables": [table.to_dict() for table in self.tables]
        }


# Helper functions for analysis
def analyze_cell_merge(cell, row_idx: int, col_idx: int) -> Optional[MergeInfo]:
    """Analyze if a cell is part of a merge and return merge information."""
    try:
        # Check if cell is merged
        if hasattr(cell, '_element'):
            tc_element = cell._element
            
            # Check for gridSpan (horizontal merge)
            grid_span = tc_element.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}gridSpan')
            
            # Check for vMerge (vertical merge)
            tc_pr = tc_element.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tcPr')
            v_merge = None
            if tc_pr is not None:
                v_merge_elem = tc_pr.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}vMerge')
                if v_merge_elem is not None:
                    v_merge = v_merge_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
            
            # Determine merge type and span
            span_cols = int(grid_span) if grid_span else 1
            span_rows = 1  # This would need more complex logic to determine vertical span
            
            if span_cols > 1 or v_merge is not None:
                merge_type = CellMergeType.NONE
                if span_cols > 1 and v_merge is not None:
                    merge_type = CellMergeType.BOTH
                elif span_cols > 1:
                    merge_type = CellMergeType.HORIZONTAL
                elif v_merge is not None:
                    merge_type = CellMergeType.VERTICAL
                
                return MergeInfo(
                    merge_type=merge_type,
                    start_row=row_idx,
                    end_row=row_idx + span_rows - 1,
                    start_col=col_idx,
                    end_col=col_idx + span_cols - 1,
                    span_rows=span_rows,
                    span_cols=span_cols
                )
    except Exception:
        pass
    
    return None


def extract_cell_formatting(cell) -> Dict[str, Any]:
    """Extract comprehensive formatting information from a cell."""
    formatting = {
        "font_family": None,
        "font_size": None,
        "font_color": None,
        "is_bold": False,
        "is_italic": False,
        "is_underlined": False,
        "is_strikethrough": False,
        "horizontal_alignment": None,
        "vertical_alignment": None,
        "background_color": None,
        "borders": {
            "top": None,
            "bottom": None,
            "left": None,
            "right": None
        }
    }
    
    try:
        # Get the first paragraph and run for text formatting
        if cell.paragraphs:
            paragraph = cell.paragraphs[0]
            
            # Paragraph alignment
            if paragraph.alignment is not None:
                alignment_map = {
                    0: "left",
                    1: "center", 
                    2: "right",
                    3: "justify"
                }
                formatting["horizontal_alignment"] = alignment_map.get(paragraph.alignment)
            
            # Run formatting (text properties)
            if paragraph.runs:
                run = paragraph.runs[0]
                
                if run.font.name:
                    formatting["font_family"] = run.font.name
                if run.font.size:
                    formatting["font_size"] = run.font.size.pt
                if run.font.color and run.font.color.rgb:
                    formatting["font_color"] = str(run.font.color.rgb)
                    
                formatting["is_bold"] = run.bold or False
                formatting["is_italic"] = run.italic or False
                formatting["is_underlined"] = run.underline or False
        
        # Cell-level formatting (background, borders)
        if hasattr(cell, '_element'):
            tc_element = cell._element
            tc_pr = tc_element.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tcPr')
            
            if tc_pr is not None:
                # Background color (shading)
                shd = tc_pr.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}shd')
                if shd is not None:
                    fill = shd.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fill')
                    if fill:
                        formatting["background_color"] = fill
                
                # Vertical alignment
                v_align = tc_pr.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}vAlign')
                if v_align is not None:
                    val = v_align.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                    formatting["vertical_alignment"] = val
                
                # Borders
                tc_borders = tc_pr.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tcBorders')
                if tc_borders is not None:
                    for border_side in ["top", "bottom", "left", "right"]:
                        border_elem = tc_borders.find(f'.//{{{tc_borders.nsmap[None]}}}{border_side}')
                        if border_elem is not None:
                            border_info = {
                                "style": border_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val'),
                                "width": border_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz'),
                                "color": border_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color')
                            }
                            formatting["borders"][border_side] = border_info
                            
    except Exception:
        # If any error occurs, return the partial formatting extracted
        pass
    
    return formatting
