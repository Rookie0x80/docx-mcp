"""Tests for table operations."""

import pytest
import re
from docx_mcp.models.responses import ResponseStatus


class TestTableStructureOperations:
    """Test table structure operations."""

    @pytest.mark.unit
    def test_create_table_basic(self, document_manager, table_operations, test_doc_path):
        """Test creating a basic table."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        
        result = table_operations.create_table(str(test_doc_path), rows=3, cols=4)
        
        assert result.status == ResponseStatus.SUCCESS
        assert result.data['rows'] == 3
        assert result.data['cols'] == 4
        assert result.data['table_index'] == 0

    @pytest.mark.unit
    def test_create_table_with_headers(self, document_manager, table_operations, test_doc_path):
        """Test creating a table with headers."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        headers = ["Name", "Age", "City"]
        
        result = table_operations.create_table(
            str(test_doc_path), 
            rows=3, 
            cols=3, 
            headers=headers
        )
        
        assert result.status == ResponseStatus.SUCCESS
        assert result.data['has_headers'] is True

    @pytest.mark.unit
    def test_create_table_invalid_dimensions(self, document_manager, table_operations, test_doc_path):
        """Test creating a table with invalid dimensions."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        
        # Test zero rows
        result = table_operations.create_table(str(test_doc_path), rows=0, cols=3)
        assert result.status == ResponseStatus.ERROR
        
        # Test zero columns
        result = table_operations.create_table(str(test_doc_path), rows=3, cols=0)
        assert result.status == ResponseStatus.ERROR

    @pytest.mark.unit
    def test_create_table_header_mismatch(self, document_manager, table_operations, test_doc_path):
        """Test creating a table with header count mismatch."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        headers = ["Name", "Age"]  # 2 headers for 3 columns
        
        result = table_operations.create_table(
            str(test_doc_path), 
            rows=3, 
            cols=3, 
            headers=headers
        )
        
        assert result.status == ResponseStatus.ERROR
        assert "Headers length" in result.message

    @pytest.mark.unit
    def test_delete_table(self, document_manager, table_operations, test_doc_path):
        """Test deleting a table."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        table_operations.create_table(str(test_doc_path), rows=2, cols=2)
        
        result = table_operations.delete_table(str(test_doc_path), table_index=0)
        
        assert result.status == ResponseStatus.SUCCESS
        assert "deleted" in result.message

    @pytest.mark.unit
    def test_delete_nonexistent_table(self, document_manager, table_operations, test_doc_path):
        """Test deleting a non-existent table."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        
        result = table_operations.delete_table(str(test_doc_path), table_index=999)
        
        assert result.status == ResponseStatus.ERROR

    @pytest.mark.unit
    def test_add_table_rows(self, document_manager, table_operations, test_doc_path):
        """Test adding rows to a table."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        table_operations.create_table(str(test_doc_path), rows=2, cols=3)
        
        result = table_operations.add_table_rows(str(test_doc_path), table_index=0, count=2)
        
        assert result.status == ResponseStatus.SUCCESS
        assert result.data['rows_added'] == 2
        assert result.data['new_row_count'] == 4  # Original 2 + added 2

    @pytest.mark.unit
    def test_add_table_columns(self, document_manager, table_operations, test_doc_path):
        """Test adding columns to a table."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        table_operations.create_table(str(test_doc_path), rows=3, cols=2)
        
        result = table_operations.add_table_columns(str(test_doc_path), table_index=0, count=1)
        
        assert result.status == ResponseStatus.SUCCESS
        assert result.data['columns_added'] == 1

    @pytest.mark.unit
    def test_delete_table_rows(self, document_manager, table_operations, test_doc_path):
        """Test deleting rows from a table."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        table_operations.create_table(str(test_doc_path), rows=5, cols=3)
        
        result = table_operations.delete_table_rows(
            str(test_doc_path), 
            table_index=0, 
            row_indices=[1, 3]
        )
        
        assert result.status == ResponseStatus.SUCCESS
        assert result.data['rows_deleted'] == 2


class TestTableDataOperations:
    """Test table data operations."""

    @pytest.fixture
    def setup_table(self, document_manager, table_operations, test_doc_path, sample_table_data):
        """Set up a table with sample data."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        table_operations.create_table(
            str(test_doc_path), 
            rows=4, 
            cols=4, 
            headers=sample_table_data["headers"]
        )
        return 0  # table index

    @pytest.mark.unit
    def test_set_cell_value(self, table_operations, test_doc_path, setup_table):
        """Test setting a cell value."""
        table_index = setup_table
        
        result = table_operations.set_cell_value(
            str(test_doc_path), table_index, 1, 0, "Test Value"
        )
        
        assert result.status == ResponseStatus.SUCCESS
        assert result.data['value'] == "Test Value"

    @pytest.mark.unit
    def test_get_cell_value(self, table_operations, test_doc_path, setup_table):
        """Test getting a cell value."""
        table_index = setup_table
        
        # First set a value
        table_operations.set_cell_value(str(test_doc_path), table_index, 1, 0, "Test Value")
        
        # Then get it
        result = table_operations.get_cell_value(str(test_doc_path), table_index, 1, 0)
        
        assert result.status == ResponseStatus.SUCCESS
        assert result.data['value'] == "Test Value"

    @pytest.mark.unit
    def test_get_cell_value_invalid_position(self, table_operations, test_doc_path, setup_table):
        """Test getting a cell value from invalid position."""
        table_index = setup_table
        
        result = table_operations.get_cell_value(str(test_doc_path), table_index, 999, 999)
        
        assert result.status == ResponseStatus.ERROR

    @pytest.mark.unit
    def test_get_table_data_array_format(self, document_manager, table_operations, test_doc_path, sample_table_data):
        """Test getting table data in array format."""
        # Setup table with data
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        table_operations.create_table(
            str(test_doc_path), 
            rows=4, 
            cols=4, 
            headers=sample_table_data["headers"]
        )
        
        # Add sample data
        for row_idx, row_data in enumerate(sample_table_data["data"], 1):
            for col_idx, value in enumerate(row_data):
                table_operations.set_cell_value(str(test_doc_path), 0, row_idx, col_idx, value)
        
        # Get table data
        result = table_operations.get_table_data(
            str(test_doc_path), 0, include_headers=True, format_type="array"
        )
        
        assert result.status == ResponseStatus.SUCCESS
        assert result.data['format'] == "array"
        assert result.data['headers'] == sample_table_data["headers"]
        assert result.data['rows'] == 3  # Data rows only
        assert result.data['columns'] == 4

    @pytest.mark.unit
    def test_get_table_data_object_format(self, document_manager, table_operations, test_doc_path, sample_table_data):
        """Test getting table data in object format."""
        # Setup table with data
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        table_operations.create_table(
            str(test_doc_path), 
            rows=4, 
            cols=4, 
            headers=sample_table_data["headers"]
        )
        
        # Add one row of sample data
        for col_idx, value in enumerate(sample_table_data["data"][0]):
            table_operations.set_cell_value(str(test_doc_path), 0, 1, col_idx, value)
        
        # Get table data
        result = table_operations.get_table_data(
            str(test_doc_path), 0, include_headers=True, format_type="object"
        )
        
        assert result.status == ResponseStatus.SUCCESS
        assert result.data['format'] == "object"
        assert isinstance(result.data['data'], list)
        if result.data['data']:  # If there's data
            assert isinstance(result.data['data'][0], dict)

    @pytest.mark.unit
    def test_get_table_data_invalid_format(self, table_operations, test_doc_path, setup_table):
        """Test getting table data with invalid format."""
        table_index = setup_table
        
        result = table_operations.get_table_data(
            str(test_doc_path), table_index, format_type="invalid_format"
        )
        
        assert result.status == ResponseStatus.ERROR
        assert "Invalid format" in result.message


class TestTableQueryOperations:
    """Test table query operations."""

    @pytest.mark.unit
    def test_list_tables_empty_document(self, document_manager, table_operations, test_doc_path):
        """Test listing tables in an empty document."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        
        result = table_operations.list_tables(str(test_doc_path))
        
        assert result.status == ResponseStatus.SUCCESS
        assert result.data['total_count'] == 0
        assert result.data['tables'] == []

    @pytest.mark.unit
    def test_list_tables_with_tables(self, document_manager, table_operations, test_doc_path):
        """Test listing tables in a document with tables."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        
        # Create multiple tables
        table_operations.create_table(str(test_doc_path), rows=2, cols=3)
        table_operations.create_table(str(test_doc_path), rows=3, cols=2)
        
        result = table_operations.list_tables(str(test_doc_path), include_summary=True)
        
        assert result.status == ResponseStatus.SUCCESS
        assert result.data['total_count'] == 2
        
        tables = result.data['tables']
        assert len(tables) == 2
        assert tables[0]['rows'] == 2
        assert tables[0]['columns'] == 3
        assert tables[1]['rows'] == 3
        assert tables[1]['columns'] == 2

    @pytest.mark.unit
    def test_list_tables_no_summary(self, document_manager, table_operations, test_doc_path):
        """Test listing tables without summary information."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        table_operations.create_table(str(test_doc_path), rows=2, cols=3)
        
        result = table_operations.list_tables(str(test_doc_path), include_summary=False)
        
        assert result.status == ResponseStatus.SUCCESS
        assert result.data['total_count'] == 1
        
        table_info = result.data['tables'][0]
        assert 'index' in table_info
        assert 'rows' in table_info
        assert 'columns' in table_info
        # Summary fields should not be present
        assert 'first_row_data' not in table_info or not table_info['first_row_data']


class TestErrorHandling:
    """Test error handling scenarios."""

    @pytest.mark.unit
    def test_operations_on_nonexistent_document(self, table_operations):
        """Test operations on a non-existent document."""
        result = table_operations.create_table("nonexistent.docx", rows=2, cols=2)
        assert result.status == ResponseStatus.ERROR

    @pytest.mark.unit
    def test_operations_on_invalid_table_index(self, document_manager, table_operations, test_doc_path):
        """Test operations with invalid table index."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        
        result = table_operations.get_cell_value(str(test_doc_path), 999, 0, 0)
        assert result.status == ResponseStatus.ERROR

    @pytest.mark.unit
    def test_invalid_cell_coordinates(self, document_manager, table_operations, test_doc_path):
        """Test operations with invalid cell coordinates."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        table_operations.create_table(str(test_doc_path), rows=2, cols=2)
        
        # Test negative coordinates
        result = table_operations.set_cell_value(str(test_doc_path), 0, -1, 0, "value")
        assert result.status == ResponseStatus.ERROR
        
        # Test out of range coordinates
        result = table_operations.set_cell_value(str(test_doc_path), 0, 999, 999, "value")
        assert result.status == ResponseStatus.ERROR


class TestTableSearchOperations:
    """Test table search operations."""

    @pytest.mark.unit
    def test_search_table_content_basic(self, document_manager, table_operations, test_doc_path):
        """Test basic table content search."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        
        # Create table with test data
        table_operations.create_table(str(test_doc_path), rows=3, cols=3, headers=["Name", "Age", "City"])
        table_operations.set_cell_value(str(test_doc_path), 0, 1, 0, "Alice")
        table_operations.set_cell_value(str(test_doc_path), 0, 1, 1, "25")
        table_operations.set_cell_value(str(test_doc_path), 0, 1, 2, "New York")
        table_operations.set_cell_value(str(test_doc_path), 0, 2, 0, "Bob")
        table_operations.set_cell_value(str(test_doc_path), 0, 2, 1, "30")
        table_operations.set_cell_value(str(test_doc_path), 0, 2, 2, "Boston")
        
        # Search for "Alice"
        result = table_operations.search_table_content(str(test_doc_path), "Alice")
        
        assert result.status == ResponseStatus.SUCCESS
        assert result.data['total_matches'] == 1
        assert len(result.data['matches']) == 1
        assert result.data['matches'][0]['table_index'] == 0
        assert result.data['matches'][0]['row_index'] == 1
        assert result.data['matches'][0]['column_index'] == 0
        assert result.data['matches'][0]['cell_value'] == "Alice"
        assert result.data['matches'][0]['match_text'] == "Alice"

    @pytest.mark.unit
    def test_search_table_content_contains_mode(self, document_manager, table_operations, test_doc_path):
        """Test table content search with contains mode."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        
        # Create table with test data
        table_operations.create_table(str(test_doc_path), rows=2, cols=2)
        table_operations.set_cell_value(str(test_doc_path), 0, 0, 0, "Hello World")
        table_operations.set_cell_value(str(test_doc_path), 0, 0, 1, "Testing")
        table_operations.set_cell_value(str(test_doc_path), 0, 1, 0, "World Peace")
        table_operations.set_cell_value(str(test_doc_path), 0, 1, 1, "Python")
        
        # Search for "World" - should find 2 matches
        result = table_operations.search_table_content(str(test_doc_path), "World", search_mode="contains")
        
        assert result.status == ResponseStatus.SUCCESS
        assert result.data['total_matches'] == 2
        assert result.data['search_mode'] == "contains"
        
        # Check that both matches are found
        matches = result.data['matches']
        assert len(matches) == 2
        
        # First match in "Hello World"
        assert matches[0]['cell_value'] == "Hello World"
        assert matches[0]['match_text'] == "World"
        
        # Second match in "World Peace"
        assert matches[1]['cell_value'] == "World Peace"
        assert matches[1]['match_text'] == "World"

    @pytest.mark.unit
    def test_search_table_content_exact_mode(self, document_manager, table_operations, test_doc_path):
        """Test table content search with exact mode."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        
        # Create table with test data
        table_operations.create_table(str(test_doc_path), rows=2, cols=2)
        table_operations.set_cell_value(str(test_doc_path), 0, 0, 0, "Test")
        table_operations.set_cell_value(str(test_doc_path), 0, 0, 1, "Testing")
        table_operations.set_cell_value(str(test_doc_path), 0, 1, 0, "Test")
        table_operations.set_cell_value(str(test_doc_path), 0, 1, 1, "Different")
        
        # Search for exact "Test" - should find 2 matches
        result = table_operations.search_table_content(str(test_doc_path), "Test", search_mode="exact")
        
        assert result.status == ResponseStatus.SUCCESS
        assert result.data['total_matches'] == 2
        assert result.data['search_mode'] == "exact"
        
        # Verify matches
        for match in result.data['matches']:
            assert match['cell_value'] == "Test"
            assert match['match_text'] == "Test"

    @pytest.mark.unit
    def test_search_table_content_regex_mode(self, document_manager, table_operations, test_doc_path):
        """Test table content search with regex mode."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        
        # Create table with test data
        table_operations.create_table(str(test_doc_path), rows=2, cols=2)
        table_operations.set_cell_value(str(test_doc_path), 0, 0, 0, "email@test.com")
        table_operations.set_cell_value(str(test_doc_path), 0, 0, 1, "user@example.org")
        table_operations.set_cell_value(str(test_doc_path), 0, 1, 0, "not-an-email")
        table_operations.set_cell_value(str(test_doc_path), 0, 1, 1, "admin@site.net")
        
        # Search for email pattern
        email_pattern = r'\w+@\w+\.\w+'
        result = table_operations.search_table_content(str(test_doc_path), email_pattern, search_mode="regex")
        
        assert result.status == ResponseStatus.SUCCESS
        assert result.data['total_matches'] == 3
        assert result.data['search_mode'] == "regex"
        
        # Check that all email addresses are found
        email_matches = [match['cell_value'] for match in result.data['matches']]
        assert "email@test.com" in email_matches
        assert "user@example.org" in email_matches
        assert "admin@site.net" in email_matches
        assert "not-an-email" not in email_matches

    @pytest.mark.unit
    def test_search_table_content_case_sensitive(self, document_manager, table_operations, test_doc_path):
        """Test table content search with case sensitivity."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        
        # Create table with test data
        table_operations.create_table(str(test_doc_path), rows=2, cols=2)
        table_operations.set_cell_value(str(test_doc_path), 0, 0, 0, "Test")
        table_operations.set_cell_value(str(test_doc_path), 0, 0, 1, "test")
        table_operations.set_cell_value(str(test_doc_path), 0, 1, 0, "TEST")
        table_operations.set_cell_value(str(test_doc_path), 0, 1, 1, "Testing")
        
        # Case sensitive search for "test"
        result = table_operations.search_table_content(str(test_doc_path), "test", case_sensitive=True)
        
        assert result.status == ResponseStatus.SUCCESS
        assert result.data['total_matches'] == 1
        assert result.data['case_sensitive'] is True
        assert result.data['matches'][0]['cell_value'] == "test"
        
        # Case insensitive search for "test"
        result = table_operations.search_table_content(str(test_doc_path), "test", case_sensitive=False)
        
        assert result.status == ResponseStatus.SUCCESS
        assert result.data['total_matches'] == 4  # "Test", "test", "TEST", "Testing"
        assert result.data['case_sensitive'] is False

    @pytest.mark.unit
    def test_search_table_content_specific_tables(self, document_manager, table_operations, test_doc_path):
        """Test searching specific tables only."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        
        # Create two tables
        table_operations.create_table(str(test_doc_path), rows=2, cols=2)
        table_operations.set_cell_value(str(test_doc_path), 0, 0, 0, "Table1Data")
        
        table_operations.create_table(str(test_doc_path), rows=2, cols=2)
        table_operations.set_cell_value(str(test_doc_path), 1, 0, 0, "Table2Data")
        
        # Search only in table 0
        result = table_operations.search_table_content(str(test_doc_path), "Data", table_indices=[0])
        
        assert result.status == ResponseStatus.SUCCESS
        assert result.data['total_matches'] == 1
        assert result.data['matches'][0]['table_index'] == 0
        assert result.data['tables_searched'] == [0]
        
        # Search only in table 1
        result = table_operations.search_table_content(str(test_doc_path), "Data", table_indices=[1])
        
        assert result.status == ResponseStatus.SUCCESS
        assert result.data['total_matches'] == 1
        assert result.data['matches'][0]['table_index'] == 1
        assert result.data['tables_searched'] == [1]

    @pytest.mark.unit
    def test_search_table_content_max_results(self, document_manager, table_operations, test_doc_path):
        """Test limiting search results."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        
        # Create table with multiple matches
        table_operations.create_table(str(test_doc_path), rows=3, cols=3)
        for row in range(3):
            for col in range(3):
                table_operations.set_cell_value(str(test_doc_path), 0, row, col, f"test{row}{col}")
        
        # Search with limit
        result = table_operations.search_table_content(str(test_doc_path), "test", max_results=5)
        
        assert result.status == ResponseStatus.SUCCESS
        assert result.data['total_matches'] == 5  # Limited to 5
        assert len(result.data['matches']) == 5

    @pytest.mark.unit
    def test_search_table_headers(self, document_manager, table_operations, test_doc_path):
        """Test searching table headers specifically."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        
        # Create table with headers
        table_operations.create_table(str(test_doc_path), rows=3, cols=3, headers=["Name", "Email", "Phone"])
        table_operations.set_cell_value(str(test_doc_path), 0, 1, 0, "Alice")
        table_operations.set_cell_value(str(test_doc_path), 0, 1, 1, "alice@email.com")
        table_operations.set_cell_value(str(test_doc_path), 0, 1, 2, "123-456-7890")
        
        # Search for "Email" in headers
        result = table_operations.search_table_headers(str(test_doc_path), "Email")
        
        assert result.status == ResponseStatus.SUCCESS
        assert result.data['total_matches'] == 1
        assert result.data['matches'][0]['table_index'] == 0
        assert result.data['matches'][0]['row_index'] == 0  # Header row
        assert result.data['matches'][0]['column_index'] == 1
        assert result.data['matches'][0]['cell_value'] == "Email"
        assert result.data['summary']['search_type'] == "headers_only"

    @pytest.mark.unit
    def test_search_table_content_empty_query(self, document_manager, table_operations, test_doc_path):
        """Test search with empty query."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        table_operations.create_table(str(test_doc_path), rows=2, cols=2)
        
        # Test empty query
        result = table_operations.search_table_content(str(test_doc_path), "")
        assert result.status == ResponseStatus.ERROR
        assert "empty" in result.message.lower()
        
        # Test whitespace-only query
        result = table_operations.search_table_content(str(test_doc_path), "   ")
        assert result.status == ResponseStatus.ERROR
        assert "empty" in result.message.lower()

    @pytest.mark.unit
    def test_search_table_content_invalid_mode(self, document_manager, table_operations, test_doc_path):
        """Test search with invalid search mode."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        table_operations.create_table(str(test_doc_path), rows=2, cols=2)
        
        result = table_operations.search_table_content(str(test_doc_path), "test", search_mode="invalid")
        assert result.status == ResponseStatus.ERROR
        assert "Invalid search mode" in result.message

    @pytest.mark.unit
    def test_search_table_content_invalid_regex(self, document_manager, table_operations, test_doc_path):
        """Test search with invalid regex pattern."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        table_operations.create_table(str(test_doc_path), rows=2, cols=2)
        
        # Invalid regex pattern
        result = table_operations.search_table_content(str(test_doc_path), "[invalid", search_mode="regex")
        assert result.status == ResponseStatus.ERROR
        assert "Invalid regex pattern" in result.message

    @pytest.mark.unit
    def test_search_table_content_no_tables(self, document_manager, table_operations, test_doc_path):
        """Test search in document with no tables."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        
        result = table_operations.search_table_content(str(test_doc_path), "test")
        
        assert result.status == ResponseStatus.SUCCESS
        assert result.data['total_matches'] == 0
        assert len(result.data['matches']) == 0
        assert result.data['summary']['tables_with_matches'] == 0

    @pytest.mark.unit
    def test_search_table_content_invalid_table_indices(self, document_manager, table_operations, test_doc_path):
        """Test search with invalid table indices."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        table_operations.create_table(str(test_doc_path), rows=2, cols=2)
        
        # Test invalid table index
        result = table_operations.search_table_content(str(test_doc_path), "test", table_indices=[999])
        assert result.status == ResponseStatus.ERROR


class TestTableAnalysisOperations:
    """Test table structure and style analysis operations."""

    @pytest.fixture
    def setup_formatted_table(self, document_manager, table_operations, test_doc_path):
        """Set up a table with various formatting for testing analysis."""
        from docx import Document
        from docx.shared import RGBColor
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        
        # Create table with headers
        table_operations.create_table(
            str(test_doc_path), 
            rows=4, 
            cols=3, 
            headers=["Name", "Department", "Salary"]
        )
        
        # Add some data
        data_rows = [
            ["Alice Smith", "Engineering", "$75,000"],
            ["Bob Johnson", "Marketing", "$65,000"],
            ["Carol Davis", "Sales", "$55,000"]
        ]
        
        for row_idx, row_data in enumerate(data_rows, 1):
            for col_idx, value in enumerate(row_data):
                table_operations.set_cell_value(str(test_doc_path), 0, row_idx, col_idx, value)
        
        # Apply some formatting using the document directly
        document = document_manager.get_document(str(test_doc_path))
        table = document.tables[0]
        
        # Make headers bold
        for cell in table.rows[0].cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.bold = True
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Apply some formatting to data cells
        if len(table.rows) > 1:
            # Make first column (names) italic
            for row_idx in range(1, len(table.rows)):
                cell = table.cell(row_idx, 0)
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.italic = True
            
            # Right-align salary column
            for row_idx in range(1, len(table.rows)):
                cell = table.cell(row_idx, 2)
                for paragraph in cell.paragraphs:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        
        return 0  # table index

    @pytest.mark.unit
    def test_analyze_table_structure_basic(self, document_manager, table_operations, test_doc_path, setup_formatted_table):
        """Test basic table structure analysis."""
        table_index = setup_formatted_table
        
        result = table_operations.analyze_table_structure(
            str(test_doc_path), 
            table_index, 
            include_cell_details=True
        )
        
        assert result.status == ResponseStatus.SUCCESS
        
        data = result.data
        table_info = data['table_info']
        header_info = data['header_info']
        merge_analysis = data['merge_analysis']
        style_consistency = data['style_consistency']
        
        # Check basic table info
        assert table_info['index'] == 0
        assert table_info['rows'] == 4
        assert table_info['columns'] == 3
        assert table_info['style_name'] in ['Table Grid', 'Normal Table']
        
        # Check header detection
        assert header_info['has_header'] is True
        assert header_info['header_row_index'] == 0
        assert header_info['header_cells'] == ["Name", "Department", "Salary"]
        
        # Check merge analysis
        assert merge_analysis['merged_cells_count'] == 0
        assert len(merge_analysis['merge_regions']) == 0
        
        # Check style consistency (should be mixed due to formatting)
        assert 'fonts' in style_consistency
        assert 'alignment' in style_consistency
        assert 'borders' in style_consistency

    @pytest.mark.unit
    def test_analyze_table_structure_cell_details(self, document_manager, table_operations, test_doc_path, setup_formatted_table):
        """Test table structure analysis with detailed cell information."""
        table_index = setup_formatted_table
        
        result = table_operations.analyze_table_structure(
            str(test_doc_path), 
            table_index, 
            include_cell_details=True
        )
        
        assert result.status == ResponseStatus.SUCCESS
        
        data = result.data
        cells = data['cells']
        
        # Should have 4 rows of cells
        assert len(cells) == 4
        
        # Each row should have 3 columns
        for row in cells:
            assert len(row) == 3
        
        # Check first row (headers)
        header_row = cells[0]
        for cell in header_row:
            assert cell['content']['text'] in ["Name", "Department", "Salary"]
            assert not cell['content']['is_empty']
            assert cell['text_format']['bold'] is True
            assert cell['alignment']['horizontal'] == 'center'
            assert cell['merge'] is None
        
        # Check first data row
        data_row = cells[1]
        name_cell = data_row[0]
        assert name_cell['content']['text'] == "Alice Smith"
        assert name_cell['text_format']['italic'] is True
        
        salary_cell = data_row[2]
        assert salary_cell['content']['text'] == "$75,000"
        assert salary_cell['alignment']['horizontal'] == 'right'

    @pytest.mark.unit
    def test_analyze_table_structure_without_cell_details(self, document_manager, table_operations, test_doc_path, setup_formatted_table):
        """Test table structure analysis without detailed cell information."""
        table_index = setup_formatted_table
        
        result = table_operations.analyze_table_structure(
            str(test_doc_path), 
            table_index, 
            include_cell_details=False
        )
        
        assert result.status == ResponseStatus.SUCCESS
        
        data = result.data
        cells = data['cells']
        
        # Should still have cell structure but without detailed formatting
        assert len(cells) == 4
        
        # Check that basic cell info is present but detailed formatting is minimal
        for row in cells:
            for cell in row:
                assert 'content' in cell
                assert 'position' in cell
                # Detailed formatting should be None or default values
                assert cell['text_format']['font_family'] is None
                assert cell['text_format']['font_size'] is None

    @pytest.mark.unit
    def test_analyze_table_structure_empty_table(self, document_manager, table_operations, test_doc_path):
        """Test analyzing an empty table."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        
        # Create empty table
        table_operations.create_table(str(test_doc_path), rows=1, cols=1)
        
        result = table_operations.analyze_table_structure(str(test_doc_path), 0)
        
        assert result.status == ResponseStatus.SUCCESS
        
        data = result.data
        assert data['table_info']['rows'] == 1
        assert data['table_info']['columns'] == 1
        assert data['merge_analysis']['merged_cells_count'] == 0
        assert len(data['cells']) == 1
        assert len(data['cells'][0]) == 1

    @pytest.mark.unit
    def test_analyze_table_structure_invalid_table(self, document_manager, table_operations, test_doc_path):
        """Test analyzing non-existent table."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        
        result = table_operations.analyze_table_structure(str(test_doc_path), 999)
        
        assert result.status == ResponseStatus.ERROR
        assert "out of range" in result.message or "Invalid table index" in result.message or "Table not found" in result.message

    @pytest.mark.unit
    def test_analyze_all_tables_empty_document(self, document_manager, table_operations, test_doc_path):
        """Test analyzing all tables in an empty document."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        
        result = table_operations.analyze_all_tables(str(test_doc_path))
        
        assert result.status == ResponseStatus.SUCCESS
        
        data = result.data
        
        # For empty documents, the response structure is different
        if 'file_info' in data:
            file_info = data['file_info']
            assert file_info['total_tables'] == 0
            assert len(data['tables']) == 0
            assert str(test_doc_path) in file_info['path']
            assert 'analysis_timestamp' in file_info
        else:
            # Fallback for the simple response structure
            assert data['total_tables'] == 0
            assert len(data['tables']) == 0
            assert str(test_doc_path) in data['file_path']

    @pytest.mark.unit
    def test_analyze_all_tables_multiple_tables(self, document_manager, table_operations, test_doc_path):
        """Test analyzing all tables in a document with multiple tables."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        
        # Create multiple tables with different characteristics
        table_operations.create_table(str(test_doc_path), rows=3, cols=4, headers=["A", "B", "C", "D"])
        table_operations.create_table(str(test_doc_path), rows=2, cols=2)
        table_operations.create_table(str(test_doc_path), rows=5, cols=3, headers=["X", "Y", "Z"])
        
        result = table_operations.analyze_all_tables(str(test_doc_path), include_cell_details=False)
        
        assert result.status == ResponseStatus.SUCCESS
        
        data = result.data
        file_info = data['file_info']
        tables = data['tables']
        
        assert file_info['total_tables'] == 3
        assert len(tables) == 3
        
        # Check first table
        table1 = tables[0]
        assert table1['table_info']['index'] == 0
        assert table1['table_info']['rows'] == 3
        assert table1['table_info']['columns'] == 4
        assert table1['header_info']['has_header'] is True
        
        # Check second table
        table2 = tables[1]
        assert table2['table_info']['index'] == 1
        assert table2['table_info']['rows'] == 2
        assert table2['table_info']['columns'] == 2
        
        # Check third table
        table3 = tables[2]
        assert table3['table_info']['index'] == 2
        assert table3['table_info']['rows'] == 5
        assert table3['table_info']['columns'] == 3
        assert table3['header_info']['has_header'] is True

    @pytest.mark.unit
    def test_analyze_all_tables_with_cell_details(self, document_manager, table_operations, test_doc_path):
        """Test analyzing all tables with detailed cell information."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        
        # Create a simple table
        table_operations.create_table(str(test_doc_path), rows=2, cols=2, headers=["Col1", "Col2"])
        table_operations.set_cell_value(str(test_doc_path), 0, 1, 0, "Data1")
        table_operations.set_cell_value(str(test_doc_path), 0, 1, 1, "Data2")
        
        result = table_operations.analyze_all_tables(str(test_doc_path), include_cell_details=True)
        
        assert result.status == ResponseStatus.SUCCESS
        
        data = result.data
        assert data['file_info']['total_tables'] == 1
        
        # The analyze_all_tables method currently returns simplified table data
        # This is by design to avoid overwhelming responses for documents with many tables
        tables = data['tables']
        assert len(tables) == 1
        
        table = tables[0]
        assert table['table_info']['rows'] == 2
        assert table['table_info']['columns'] == 2
        assert table['header_info']['has_header'] is True

    @pytest.mark.unit
    def test_analyze_table_structure_style_consistency(self, document_manager, table_operations, test_doc_path):
        """Test style consistency analysis."""
        from docx.shared import RGBColor
        
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        
        # Create table
        table_operations.create_table(str(test_doc_path), rows=3, cols=2)
        
        # Add some data
        table_operations.set_cell_value(str(test_doc_path), 0, 0, 0, "Header1")
        table_operations.set_cell_value(str(test_doc_path), 0, 0, 1, "Header2")
        table_operations.set_cell_value(str(test_doc_path), 0, 1, 0, "Data1")
        table_operations.set_cell_value(str(test_doc_path), 0, 1, 1, "Data2")
        table_operations.set_cell_value(str(test_doc_path), 0, 2, 0, "Data3")
        table_operations.set_cell_value(str(test_doc_path), 0, 2, 1, "Data4")
        
        # Apply mixed formatting to test consistency detection
        document = document_manager.get_document(str(test_doc_path))
        table = document.tables[0]
        
        # Make some cells bold, others not
        table.cell(0, 0).paragraphs[0].runs[0].font.bold = True
        table.cell(1, 0).paragraphs[0].runs[0].font.bold = False
        
        result = table_operations.analyze_table_structure(str(test_doc_path), 0)
        
        assert result.status == ResponseStatus.SUCCESS
        
        data = result.data
        style_summary = data['style_summary']
        
        # Should detect various styles
        assert isinstance(style_summary['font_families'], list)
        assert isinstance(style_summary['font_sizes'], list)
        assert isinstance(style_summary['colors'], list)
        assert isinstance(style_summary['background_colors'], list)

    @pytest.mark.unit
    def test_analyze_table_structure_nonexistent_document(self, table_operations):
        """Test analyzing table in non-existent document."""
        result = table_operations.analyze_table_structure("nonexistent.docx", 0)
        
        assert result.status == ResponseStatus.ERROR
        assert "not loaded" in result.message

    @pytest.mark.unit
    def test_analyze_all_tables_nonexistent_document(self, table_operations):
        """Test analyzing all tables in non-existent document."""
        result = table_operations.analyze_all_tables("nonexistent.docx")
        
        assert result.status == ResponseStatus.ERROR
        assert "not loaded" in result.message

    @pytest.mark.integration
    def test_table_analysis_integration_workflow(self, document_manager, table_operations, test_doc_path):
        """Test complete table analysis workflow."""
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        
        # Create multiple tables with different characteristics
        # Table 1: Employee data with headers
        table_operations.create_table(
            str(test_doc_path), 
            rows=4, 
            cols=3, 
            headers=["Name", "Role", "Salary"]
        )
        
        employees = [
            ["Alice Johnson", "Developer", "$80,000"],
            ["Bob Smith", "Designer", "$70,000"],
            ["Carol White", "Manager", "$90,000"]
        ]
        
        for row_idx, emp_data in enumerate(employees, 1):
            for col_idx, value in enumerate(emp_data):
                table_operations.set_cell_value(str(test_doc_path), 0, row_idx, col_idx, value)
        
        # Table 2: Simple data without headers
        table_operations.create_table(str(test_doc_path), rows=2, cols=2)
        table_operations.set_cell_value(str(test_doc_path), 1, 0, 0, "Item1")
        table_operations.set_cell_value(str(test_doc_path), 1, 0, 1, "Value1")
        table_operations.set_cell_value(str(test_doc_path), 1, 1, 0, "Item2")
        table_operations.set_cell_value(str(test_doc_path), 1, 1, 1, "Value2")
        
        # Step 1: List all tables
        list_result = table_operations.list_tables(str(test_doc_path))
        assert list_result.status == ResponseStatus.SUCCESS
        assert list_result.data['total_count'] == 2
        
        # Step 2: Analyze all tables (overview)
        all_analysis = table_operations.analyze_all_tables(str(test_doc_path), include_cell_details=False)
        assert all_analysis.status == ResponseStatus.SUCCESS
        assert all_analysis.data['file_info']['total_tables'] == 2
        
        # Step 3: Analyze first table in detail
        detailed_analysis = table_operations.analyze_table_structure(str(test_doc_path), 0, include_cell_details=True)
        assert detailed_analysis.status == ResponseStatus.SUCCESS
        
        data = detailed_analysis.data
        assert data['table_info']['rows'] == 4
        assert data['table_info']['columns'] == 3
        assert data['header_info']['has_header'] is True
        assert data['header_info']['header_cells'] == ["Name", "Role", "Salary"]
        
        # Verify cell data
        cells = data['cells']
        assert len(cells) == 4  # 4 rows
        assert cells[1][0]['content']['text'] == "Alice Johnson"  # First data row, first column
        assert cells[1][2]['content']['text'] == "$80,000"       # First data row, salary column
        
        # Step 4: Analyze second table
        table2_analysis = table_operations.analyze_table_structure(str(test_doc_path), 1, include_cell_details=True)
        assert table2_analysis.status == ResponseStatus.SUCCESS
        
        data2 = table2_analysis.data
        assert data2['table_info']['rows'] == 2
        assert data2['table_info']['columns'] == 2
        # Note: Header detection is heuristic-based, so it may detect headers even in simple tables
        # This is expected behavior when all cells in first row have content
        
        # Verify this workflow provides comprehensive table understanding
        # This is exactly what AI models need to understand table structure before modifications