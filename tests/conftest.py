"""Pytest configuration and fixtures."""

import os
import sys
import tempfile
import pytest
from pathlib import Path

# Add the src directory to the path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

from docx_mcp.core.document_manager import DocumentManager
from docx_mcp.operations.tables.table_operations import TableOperations


@pytest.fixture
def temp_dir():
    """Create a temporary directory for test files."""
    with tempfile.TemporaryDirectory() as tmp_dir:
        yield Path(tmp_dir)


@pytest.fixture
def document_manager():
    """Create a document manager instance."""
    return DocumentManager()


@pytest.fixture
def table_operations(document_manager):
    """Create a table operations instance."""
    return TableOperations(document_manager)


@pytest.fixture
def test_doc_path(temp_dir):
    """Create a test document path."""
    return temp_dir / "test_document.docx"


@pytest.fixture
def sample_table_data():
    """Sample table data for testing."""
    return {
        "headers": ["Name", "Age", "City", "Occupation"],
        "data": [
            ["Alice", "28", "New York", "Engineer"],
            ["Bob", "35", "San Francisco", "Designer"],
            ["Carol", "42", "Chicago", "Manager"]
        ]
    }
