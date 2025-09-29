"""Data models for table and cell formatting operations."""

from dataclasses import dataclass
from typing import Optional, Dict, Any
from enum import Enum


class HorizontalAlignment(Enum):
    """Horizontal text alignment options."""
    LEFT = "left"
    CENTER = "center"
    RIGHT = "right"
    JUSTIFY = "justify"


class VerticalAlignment(Enum):
    """Vertical text alignment options."""
    TOP = "top"
    MIDDLE = "middle"
    BOTTOM = "bottom"


class BorderStyle(Enum):
    """Border style options."""
    NONE = "none"
    SOLID = "solid"
    DASHED = "dashed"
    DOTTED = "dotted"
    DOUBLE = "double"


class BorderWidth(Enum):
    """Border width options."""
    THIN = "thin"
    MEDIUM = "medium"
    THICK = "thick"


@dataclass
class TextFormat:
    """Text formatting properties for cells."""
    font_family: Optional[str] = None
    font_size: Optional[int] = None  # Points (8-72)
    font_color: Optional[str] = None  # Hex color or RGB
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    underline: Optional[bool] = None
    strikethrough: Optional[bool] = None
    subscript: Optional[bool] = None
    superscript: Optional[bool] = None

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary representation."""
        return {
            k: v for k, v in self.__dict__.items() 
            if v is not None
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "TextFormat":
        """Create from dictionary representation."""
        return cls(**{
            k: v for k, v in data.items() 
            if k in cls.__dataclass_fields__
        })


@dataclass
class CellAlignment:
    """Cell text alignment properties."""
    horizontal: Optional[HorizontalAlignment] = None
    vertical: Optional[VerticalAlignment] = None

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary representation."""
        result = {}
        if self.horizontal:
            result["horizontal"] = self.horizontal.value
        if self.vertical:
            result["vertical"] = self.vertical.value
        return result

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "CellAlignment":
        """Create from dictionary representation."""
        horizontal = None
        vertical = None
        
        if "horizontal" in data:
            try:
                horizontal = HorizontalAlignment(data["horizontal"])
            except ValueError:
                pass
                
        if "vertical" in data:
            try:
                vertical = VerticalAlignment(data["vertical"])
            except ValueError:
                pass
                
        return cls(horizontal=horizontal, vertical=vertical)


@dataclass
class BorderProperties:
    """Properties for a single border (top, bottom, left, right)."""
    style: BorderStyle = BorderStyle.SOLID
    width: BorderWidth = BorderWidth.THIN
    color: str = "000000"  # Hex color without #

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary representation."""
        return {
            "style": self.style.value,
            "width": self.width.value,
            "color": self.color
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "BorderProperties":
        """Create from dictionary representation."""
        style = BorderStyle.SOLID
        width = BorderWidth.THIN
        color = data.get("color", "000000")
        
        if "style" in data:
            try:
                style = BorderStyle(data["style"])
            except ValueError:
                pass
                
        if "width" in data:
            try:
                width = BorderWidth(data["width"])
            except ValueError:
                pass
                
        return cls(style=style, width=width, color=color)


@dataclass
class CellBorders:
    """Border configuration for all sides of a cell."""
    top: Optional[BorderProperties] = None
    bottom: Optional[BorderProperties] = None
    left: Optional[BorderProperties] = None
    right: Optional[BorderProperties] = None

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary representation."""
        result = {}
        if self.top:
            result["top"] = self.top.to_dict()
        if self.bottom:
            result["bottom"] = self.bottom.to_dict()
        if self.left:
            result["left"] = self.left.to_dict()
        if self.right:
            result["right"] = self.right.to_dict()
        return result

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "CellBorders":
        """Create from dictionary representation."""
        borders = {}
        for side in ["top", "bottom", "left", "right"]:
            if side in data:
                borders[side] = BorderProperties.from_dict(data[side])
        return cls(**borders)


@dataclass
class CellFormatting:
    """Complete formatting configuration for a cell."""
    text_format: Optional[TextFormat] = None
    alignment: Optional[CellAlignment] = None
    background_color: Optional[str] = None  # Hex color without #
    borders: Optional[CellBorders] = None

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary representation."""
        result = {}
        if self.text_format:
            result["text_format"] = self.text_format.to_dict()
        if self.alignment:
            result["alignment"] = self.alignment.to_dict()
        if self.background_color:
            result["background_color"] = self.background_color
        if self.borders:
            result["borders"] = self.borders.to_dict()
        return result

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "CellFormatting":
        """Create from dictionary representation."""
        text_format = None
        alignment = None
        borders = None
        
        if "text_format" in data:
            text_format = TextFormat.from_dict(data["text_format"])
        if "alignment" in data:
            alignment = CellAlignment.from_dict(data["alignment"])
        if "borders" in data:
            borders = CellBorders.from_dict(data["borders"])
            
        return cls(
            text_format=text_format,
            alignment=alignment,
            background_color=data.get("background_color"),
            borders=borders
        )


# Utility functions for color handling
def hex_to_rgb(hex_color: str) -> tuple:
    """Convert hex color to RGB tuple."""
    hex_color = hex_color.lstrip('#')
    if len(hex_color) != 6:
        raise ValueError("Invalid hex color format")
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))


def rgb_to_hex(r: int, g: int, b: int) -> str:
    """Convert RGB values to hex color string."""
    return f"{r:02x}{g:02x}{b:02x}"


def validate_color(color: str) -> bool:
    """Validate if color string is a valid hex color."""
    try:
        color = color.lstrip('#')
        if len(color) != 6:
            return False
        int(color, 16)
        return True
    except ValueError:
        return False


# Predefined color constants
class Colors:
    """Common color constants."""
    BLACK = "000000"
    WHITE = "FFFFFF"
    RED = "FF0000"
    GREEN = "00FF00"
    BLUE = "0000FF"
    YELLOW = "FFFF00"
    CYAN = "00FFFF"
    MAGENTA = "FF00FF"
    GRAY = "808080"
    LIGHT_GRAY = "D3D3D3"
    DARK_GRAY = "A9A9A9"


# Font family constants
class Fonts:
    """Common font family constants."""
    ARIAL = "Arial"
    TIMES_NEW_ROMAN = "Times New Roman"
    CALIBRI = "Calibri"
    HELVETICA = "Helvetica"
    GEORGIA = "Georgia"
    VERDANA = "Verdana"
    TAHOMA = "Tahoma"
    COURIER_NEW = "Courier New"
