"""Response models for MCP operations."""

from dataclasses import dataclass
from typing import Any, Optional, Dict, List
from enum import Enum


class ResponseStatus(Enum):
    """Status codes for operation responses."""
    SUCCESS = "success"
    ERROR = "error"
    WARNING = "warning"


@dataclass
class OperationResponse:
    """Standard response format for MCP operations."""
    status: ResponseStatus
    message: str
    data: Optional[Any] = None
    error_code: Optional[str] = None
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary format."""
        result = {
            "status": self.status.value,
            "message": self.message
        }
        if self.data is not None:
            result["data"] = self.data
        if self.error_code:
            result["error_code"] = self.error_code
        return result
    
    @classmethod
    def success(cls, message: str, data: Optional[Any] = None) -> "OperationResponse":
        """Create a success response."""
        return cls(ResponseStatus.SUCCESS, message, data)
    
    @classmethod
    def error(cls, message: str, error_code: Optional[str] = None) -> "OperationResponse":
        """Create an error response."""
        return cls(ResponseStatus.ERROR, message, error_code=error_code)
    
    @classmethod
    def warning(cls, message: str, data: Optional[Any] = None) -> "OperationResponse":
        """Create a warning response."""
        return cls(ResponseStatus.WARNING, message, data)


@dataclass
class TableListResponse:
    """Response for listing tables in a document."""
    tables: List[Dict[str, Any]]
    total_count: int
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary format."""
        return {
            "tables": self.tables,
            "total_count": self.total_count
        }
