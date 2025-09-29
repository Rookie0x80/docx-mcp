"""Tests for table and cell formatting functionality."""

import pytest
from pathlib import Path

from docx_mcp.models.responses import ResponseStatus
from docx_mcp.models.formatting import (
    TextFormat, CellAlignment, CellBorders, BorderProperties,
    HorizontalAlignment, VerticalAlignment, BorderStyle, BorderWidth,
    Colors, Fonts
)
from docx_mcp.operations.tables.formatting import TableFormattingOperations


class TestCellTextFormatting:
    """Test cell text formatting operations."""
    
    def test_format_cell_text_basic(self, document_manager, table_operations, test_doc_path):
        """Test basic text formatting."""
        # Create document and table
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        table_operations.create_table(str(test_doc_path), rows=2, cols=2)
        table_operations.set_cell_value(str(test_doc_path), 0, 0, 0, "Test Text")
        
        # Apply text formatting
        formatting_ops = TableFormattingOperations(document_manager)
        text_format = TextFormat(
            font_family=Fonts.ARIAL,
            font_size=14,
            font_color=Colors.RED,
            bold=True,
            italic=True
        )
        
        result = formatting_ops.format_cell_text(
            str(test_doc_path), 0, 0, 0, text_format
        )
        
        assert result.status == ResponseStatus.SUCCESS
        assert "Text formatting applied" in result.message
        assert result.data["table_index"] == 0
        assert result.data["row_index"] == 0
        assert result.data["column_index"] == 0
    
    def test_format_cell_text_invalid_cell(self, document_manager, table_operations, test_doc_path):
        """Test text formatting with invalid cell position."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        table_operations.create_table(str(test_doc_path), rows=2, cols=2)
        
        formatting_ops = TableFormattingOperations(document_manager)
        text_format = TextFormat(bold=True)
        
        result = formatting_ops.format_cell_text(
            str(test_doc_path), 0, 5, 5, text_format  # Invalid position
        )
        
        assert result.status == ResponseStatus.ERROR
        assert "Failed to format cell text" in result.message
    
    def test_format_cell_text_from_dict(self, document_manager, table_operations, test_doc_path):
        """Test text formatting using dictionary input."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        table_operations.create_table(str(test_doc_path), rows=2, cols=2)
        
        formatting_ops = TableFormattingOperations(document_manager)
        format_dict = {
            "font_family": "Times New Roman",
            "font_size": 16,
            "bold": True,
            "underline": True
        }
        
        result = formatting_ops.format_cell_text(
            str(test_doc_path), 0, 0, 0, format_dict
        )
        
        assert result.status == ResponseStatus.SUCCESS


class TestCellAlignment:
    """Test cell alignment operations."""
    
    def test_format_cell_alignment_horizontal(self, document_manager, table_operations, test_doc_path):
        """Test horizontal alignment."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        table_operations.create_table(str(test_doc_path), rows=2, cols=2)
        
        formatting_ops = TableFormattingOperations(document_manager)
        alignment = CellAlignment(horizontal=HorizontalAlignment.CENTER)
        
        result = formatting_ops.format_cell_alignment(
            str(test_doc_path), 0, 0, 0, alignment
        )
        
        assert result.status == ResponseStatus.SUCCESS
        assert "Alignment applied" in result.message
    
    def test_format_cell_alignment_vertical(self, document_manager, table_operations, test_doc_path):
        """Test vertical alignment."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        table_operations.create_table(str(test_doc_path), rows=2, cols=2)
        
        formatting_ops = TableFormattingOperations(document_manager)
        alignment = CellAlignment(vertical=VerticalAlignment.MIDDLE)
        
        result = formatting_ops.format_cell_alignment(
            str(test_doc_path), 0, 0, 0, alignment
        )
        
        assert result.status == ResponseStatus.SUCCESS
    
    def test_format_cell_alignment_both(self, document_manager, table_operations, test_doc_path):
        """Test both horizontal and vertical alignment."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        table_operations.create_table(str(test_doc_path), rows=2, cols=2)
        
        formatting_ops = TableFormattingOperations(document_manager)
        alignment = CellAlignment(
            horizontal=HorizontalAlignment.RIGHT,
            vertical=VerticalAlignment.BOTTOM
        )
        
        result = formatting_ops.format_cell_alignment(
            str(test_doc_path), 0, 0, 0, alignment
        )
        
        assert result.status == ResponseStatus.SUCCESS


class TestCellBackground:
    """Test cell background color operations."""
    
    def test_format_cell_background_valid_color(self, document_manager, table_operations, test_doc_path):
        """Test setting valid background color."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        table_operations.create_table(str(test_doc_path), rows=2, cols=2)
        
        formatting_ops = TableFormattingOperations(document_manager)
        
        result = formatting_ops.format_cell_background(
            str(test_doc_path), 0, 0, 0, Colors.YELLOW
        )
        
        assert result.status == ResponseStatus.SUCCESS
        assert "Background color applied" in result.message
        assert result.data["background_color"] == Colors.YELLOW
    
    def test_format_cell_background_with_hash(self, document_manager, table_operations, test_doc_path):
        """Test setting background color with # prefix."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        table_operations.create_table(str(test_doc_path), rows=2, cols=2)
        
        formatting_ops = TableFormattingOperations(document_manager)
        
        result = formatting_ops.format_cell_background(
            str(test_doc_path), 0, 0, 0, "#00FF00"  # Green with #
        )
        
        assert result.status == ResponseStatus.SUCCESS
        assert result.data["background_color"] == "00FF00"
    
    def test_format_cell_background_invalid_color(self, document_manager, table_operations, test_doc_path):
        """Test setting invalid background color."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        table_operations.create_table(str(test_doc_path), rows=2, cols=2)
        
        formatting_ops = TableFormattingOperations(document_manager)
        
        result = formatting_ops.format_cell_background(
            str(test_doc_path), 0, 0, 0, "invalid_color"
        )
        
        assert result.status == ResponseStatus.ERROR
        assert "Invalid color format" in result.message


class TestCellBorders:
    """Test cell border operations."""
    
    def test_format_cell_borders_single_side(self, document_manager, table_operations, test_doc_path):
        """Test setting border for single side."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        table_operations.create_table(str(test_doc_path), rows=2, cols=2)
        
        formatting_ops = TableFormattingOperations(document_manager)
        borders = CellBorders(
            top=BorderProperties(
                style=BorderStyle.SOLID,
                width=BorderWidth.THICK,
                color=Colors.BLACK
            )
        )
        
        result = formatting_ops.format_cell_borders(
            str(test_doc_path), 0, 0, 0, borders
        )
        
        assert result.status == ResponseStatus.SUCCESS
        assert "Borders applied" in result.message
    
    def test_format_cell_borders_all_sides(self, document_manager, table_operations, test_doc_path):
        """Test setting borders for all sides."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        table_operations.create_table(str(test_doc_path), rows=2, cols=2)
        
        formatting_ops = TableFormattingOperations(document_manager)
        border_prop = BorderProperties(
            style=BorderStyle.DOUBLE,
            width=BorderWidth.MEDIUM,
            color=Colors.BLUE
        )
        borders = CellBorders(
            top=border_prop,
            bottom=border_prop,
            left=border_prop,
            right=border_prop
        )
        
        result = formatting_ops.format_cell_borders(
            str(test_doc_path), 0, 0, 0, borders
        )
        
        assert result.status == ResponseStatus.SUCCESS
    
    def test_format_cell_borders_from_dict(self, document_manager, table_operations, test_doc_path):
        """Test setting borders using dictionary input."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        table_operations.create_table(str(test_doc_path), rows=2, cols=2)
        
        formatting_ops = TableFormattingOperations(document_manager)
        borders_dict = {
            "top": {"style": "solid", "width": "thick", "color": "FF0000"},
            "bottom": {"style": "dashed", "width": "thin", "color": "00FF00"}
        }
        
        result = formatting_ops.format_cell_borders(
            str(test_doc_path), 0, 0, 0, borders_dict
        )
        
        assert result.status == ResponseStatus.SUCCESS


class TestCompleteFormatting:
    """Test complete cell formatting operations."""
    
    def test_format_cell_complete(self, document_manager, table_operations, test_doc_path):
        """Test applying complete formatting to a cell."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        table_operations.create_table(str(test_doc_path), rows=2, cols=2)
        table_operations.set_cell_value(str(test_doc_path), 0, 0, 0, "Formatted Cell")
        
        formatting_ops = TableFormattingOperations(document_manager)
        
        complete_formatting = {
            "text_format": {
                "font_family": "Arial",
                "font_size": 12,
                "font_color": "000080",
                "bold": True,
                "italic": False
            },
            "alignment": {
                "horizontal": "center",
                "vertical": "middle"
            },
            "background_color": "F0F0F0",
            "borders": {
                "top": {"style": "solid", "width": "medium", "color": "000000"},
                "bottom": {"style": "solid", "width": "medium", "color": "000000"},
                "left": {"style": "solid", "width": "medium", "color": "000000"},
                "right": {"style": "solid", "width": "medium", "color": "000000"}
            }
        }
        
        result = formatting_ops.format_cell_complete(
            str(test_doc_path), 0, 0, 0, complete_formatting
        )
        
        assert result.status == ResponseStatus.SUCCESS
        assert "Complete formatting applied" in result.message
        assert len(result.data["applied_formats"]) == 4  # text, alignment, background, borders
    
    def test_format_cell_complete_partial(self, document_manager, table_operations, test_doc_path):
        """Test applying partial complete formatting."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        table_operations.create_table(str(test_doc_path), rows=2, cols=2)
        
        formatting_ops = TableFormattingOperations(document_manager)
        
        partial_formatting = {
            "text_format": {
                "bold": True,
                "font_size": 16
            },
            "background_color": "FFFF00"
        }
        
        result = formatting_ops.format_cell_complete(
            str(test_doc_path), 0, 0, 0, partial_formatting
        )
        
        assert result.status == ResponseStatus.SUCCESS
        assert len(result.data["applied_formats"]) == 2  # text and background only


class TestFormattingModels:
    """Test formatting data models."""
    
    def test_text_format_to_dict(self):
        """Test TextFormat to_dict conversion."""
        text_format = TextFormat(
            font_family="Arial",
            font_size=12,
            bold=True,
            italic=False
        )
        
        result = text_format.to_dict()
        
        assert result["font_family"] == "Arial"
        assert result["font_size"] == 12
        assert result["bold"] is True
        assert result["italic"] is False
        assert "underline" not in result  # None values should be excluded
    
    def test_text_format_from_dict(self):
        """Test TextFormat from_dict conversion."""
        data = {
            "font_family": "Times New Roman",
            "font_size": 14,
            "font_color": "FF0000",
            "bold": True,
            "unknown_field": "ignored"  # Should be ignored
        }
        
        text_format = TextFormat.from_dict(data)
        
        assert text_format.font_family == "Times New Roman"
        assert text_format.font_size == 14
        assert text_format.font_color == "FF0000"
        assert text_format.bold is True
        assert text_format.italic is None  # Not in dict
    
    def test_cell_alignment_to_dict(self):
        """Test CellAlignment to_dict conversion."""
        alignment = CellAlignment(
            horizontal=HorizontalAlignment.CENTER,
            vertical=VerticalAlignment.MIDDLE
        )
        
        result = alignment.to_dict()
        
        assert result["horizontal"] == "center"
        assert result["vertical"] == "middle"
    
    def test_border_properties_from_dict(self):
        """Test BorderProperties from_dict conversion."""
        data = {
            "style": "dashed",
            "width": "thick",
            "color": "FF0000"
        }
        
        border = BorderProperties.from_dict(data)
        
        assert border.style == BorderStyle.DASHED
        assert border.width == BorderWidth.THICK
        assert border.color == "FF0000"


class TestColorValidation:
    """Test color validation utilities."""
    
    def test_validate_color_valid(self):
        """Test validation of valid colors."""
        from docx_mcp.models.formatting import validate_color
        
        assert validate_color("FF0000") is True
        assert validate_color("00FF00") is True
        assert validate_color("0000FF") is True
        assert validate_color("FFFFFF") is True
        assert validate_color("000000") is True
    
    def test_validate_color_invalid(self):
        """Test validation of invalid colors."""
        from docx_mcp.models.formatting import validate_color
        
        assert validate_color("invalid") is False
        assert validate_color("FF00") is False  # Too short
        assert validate_color("FF0000FF") is False  # Too long
        assert validate_color("GGGGGG") is False  # Invalid hex
    
    def test_hex_to_rgb(self):
        """Test hex to RGB conversion."""
        from docx_mcp.models.formatting import hex_to_rgb
        
        assert hex_to_rgb("FF0000") == (255, 0, 0)
        assert hex_to_rgb("00FF00") == (0, 255, 0)
        assert hex_to_rgb("0000FF") == (0, 0, 255)
        assert hex_to_rgb("#FFFFFF") == (255, 255, 255)
    
    def test_rgb_to_hex(self):
        """Test RGB to hex conversion."""
        from docx_mcp.models.formatting import rgb_to_hex
        
        assert rgb_to_hex(255, 0, 0) == "ff0000"
        assert rgb_to_hex(0, 255, 0) == "00ff00"
        assert rgb_to_hex(0, 0, 255) == "0000ff"
        assert rgb_to_hex(255, 255, 255) == "ffffff"
