"""Tests for FastMCP server integration."""

import os
import pytest
from unittest.mock import Mock, patch
import tempfile
from pathlib import Path

from docx_mcp.server import (
    document_manager, table_operations
)


class TestMCPToolSignatures:
    """Test that MCP tools have correct signatures for JSON parameter handling."""

    @pytest.mark.unit
    def test_tool_functions_exist(self):
        """Test that all MCP tool functions exist and are callable."""
        from docx_mcp import server
        
        # Check that the FastMCP app exists
        assert hasattr(server, 'mcp')
        assert server.mcp is not None
        
        # Check that managers are initialized
        assert hasattr(server, 'document_manager')
        assert hasattr(server, 'table_operations')
        assert server.document_manager is not None
        assert server.table_operations is not None

    @pytest.mark.unit  
    def test_tool_function_signatures(self):
        """Test that tool functions have correct parameter signatures."""
        import inspect
        from docx_mcp.server import (
            open_document, save_document, get_document_info,
            create_table, delete_table, add_table_rows,
            set_cell_value, get_cell_value, get_table_data,
            list_tables
        )
        
        # Test open_document signature
        sig = inspect.signature(open_document.fn)  # Access underlying function
        params = list(sig.parameters.keys())
        assert 'file_path' in params
        assert 'create_if_not_exists' in params
        
        # Test create_table signature
        sig = inspect.signature(create_table.fn)
        params = list(sig.parameters.keys())
        expected_params = ['file_path', 'rows', 'cols', 'position', 'paragraph_index', 'headers']
        for param in expected_params:
            assert param in params, f"Parameter '{param}' missing from create_table"
        
        # Test set_cell_value signature
        sig = inspect.signature(set_cell_value.fn)
        params = list(sig.parameters.keys())
        expected_params = ['file_path', 'table_index', 'row_index', 'column_index', 'value']
        for param in expected_params:
            assert param in params, f"Parameter '{param}' missing from set_cell_value"


class TestServerIntegration:
    """Test server integration with underlying operations."""

    @pytest.fixture
    def temp_doc_path(self):
        """Create a temporary document path."""
        # Create a temporary file path without creating the file
        temp_dir = Path(tempfile.gettempdir())
        temp_file = temp_dir / f"test_docx_{os.getpid()}_{id(self)}.docx"
        
        yield str(temp_file)
        
        # Cleanup
        if temp_file.exists():
            temp_file.unlink()

    @pytest.mark.integration
    def test_document_manager_integration(self, temp_doc_path):
        """Test document manager integration."""
        # Test creating document
        result = document_manager.open_document(temp_doc_path, create_if_not_exists=True)
        assert result.status.value == 'success'
        assert result.data['is_new'] is True
        
        # Test saving document
        result = document_manager.save_document(temp_doc_path)
        assert result.status.value == 'success'
        assert Path(temp_doc_path).exists()

    @pytest.mark.integration
    def test_table_operations_integration(self, temp_doc_path):
        """Test table operations integration."""
        # Setup: open document
        document_manager.open_document(temp_doc_path, create_if_not_exists=True)
        
        # Create table
        result = table_operations.create_table(
            temp_doc_path, 
            rows=3, 
            cols=3,
            headers=["A", "B", "C"]
        )
        assert result.status.value == 'success'
        assert result.data['table_index'] == 0
        
        # Set cell value
        result = table_operations.set_cell_value(temp_doc_path, 0, 1, 0, "Test")
        assert result.status.value == 'success'
        
        # Get cell value
        result = table_operations.get_cell_value(temp_doc_path, 0, 1, 0)
        assert result.status.value == 'success'
        assert result.data['value'] == "Test"
        
        # List tables
        result = table_operations.list_tables(temp_doc_path)
        assert result.status.value == 'success'
        assert result.data['total_count'] == 1

    @pytest.mark.integration
    def test_full_workflow_integration(self, temp_doc_path):
        """Test a complete workflow."""
        # 1. Open document
        result = document_manager.open_document(temp_doc_path, create_if_not_exists=True)
        assert result.status.value == 'success'
        
        # 2. Create table with headers
        headers = ["Name", "Age", "City"]
        result = table_operations.create_table(temp_doc_path, 3, 3, headers=headers)
        assert result.status.value == 'success'
        table_index = result.data['table_index']
        
        # 3. Add data to table
        test_data = [
            ["Alice", "25", "New York"],
            ["Bob", "30", "London"]
        ]
        
        for row_idx, row_data in enumerate(test_data, 1):  # Start from row 1 (after headers)
            for col_idx, value in enumerate(row_data):
                result = table_operations.set_cell_value(temp_doc_path, table_index, row_idx, col_idx, value)
                assert result.status.value == 'success'
        
        # 4. Get table data
        result = table_operations.get_table_data(temp_doc_path, table_index, include_headers=True)
        assert result.status.value == 'success'
        assert result.data['headers'] == headers
        assert result.data['rows'] == 2  # Data rows only
        
        # 5. Add a row
        result = table_operations.add_table_rows(temp_doc_path, table_index, count=1)
        assert result.status.value == 'success'
        assert result.data['new_row_count'] == 4  # Original 3 + 1 added
        
        # 6. Save document
        result = document_manager.save_document(temp_doc_path)
        assert result.status.value == 'success'
        
        # Verify file exists
        assert Path(temp_doc_path).exists()

    @pytest.mark.integration
    def test_error_handling_integration(self, temp_doc_path):
        """Test error handling in integration scenarios."""
        # Test operations on non-existent document
        result = table_operations.create_table("nonexistent.docx", 2, 2)
        assert result.status.value == 'error'
        
        # Test invalid table operations
        document_manager.open_document(temp_doc_path, create_if_not_exists=True)
        
        # Try to get cell from non-existent table
        result = table_operations.get_cell_value(temp_doc_path, 999, 0, 0)
        assert result.status.value == 'error'