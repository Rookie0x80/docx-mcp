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
