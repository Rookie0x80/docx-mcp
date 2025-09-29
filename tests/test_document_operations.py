"""Tests for document operations."""

import pytest
from docx_mcp.models.responses import ResponseStatus


class TestDocumentOperations:
    """Test document management operations."""

    @pytest.mark.unit
    def test_create_new_document(self, document_manager, test_doc_path):
        """Test creating a new document."""
        result = document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        
        assert result.status == ResponseStatus.SUCCESS
        assert "Created new document" in result.message
        assert result.data['is_new'] is True
        assert result.data['table_count'] == 0
        assert result.data['paragraph_count'] >= 0

    @pytest.mark.unit
    def test_open_existing_document(self, document_manager, test_doc_path):
        """Test opening an existing document."""
        # First create the document
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        document_manager.save_document(str(test_doc_path))
        
        # Then open the existing document
        result = document_manager.open_document(str(test_doc_path), create_if_not_exists=False)
        
        assert result.status == ResponseStatus.SUCCESS
        assert "Opened existing document" in result.message

    @pytest.mark.unit
    def test_open_nonexistent_document_no_create(self, document_manager, test_doc_path):
        """Test opening a non-existent document without creating it."""
        result = document_manager.open_document(str(test_doc_path), create_if_not_exists=False)
        
        assert result.status == ResponseStatus.ERROR
        assert "Document not found" in result.message

    @pytest.mark.unit
    def test_save_document(self, document_manager, test_doc_path):
        """Test saving a document."""
        # Create and open document
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        
        # Save document
        result = document_manager.save_document(str(test_doc_path))
        
        assert result.status == ResponseStatus.SUCCESS
        assert "Document saved" in result.message
        assert test_doc_path.exists()

    @pytest.mark.unit
    def test_save_document_as(self, document_manager, test_doc_path, temp_dir):
        """Test saving a document with a different name."""
        # Create and open document
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        
        # Save as different file
        save_as_path = temp_dir / "saved_as_document.docx"
        result = document_manager.save_document(str(test_doc_path), str(save_as_path))
        
        assert result.status == ResponseStatus.SUCCESS
        assert "Document saved" in result.message
        assert save_as_path.exists()

    @pytest.mark.unit
    def test_get_document_info(self, document_manager, table_operations, test_doc_path):
        """Test getting document information."""
        # Create document with a table
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        table_operations.create_table(str(test_doc_path), rows=3, cols=3)
        
        # Get document info
        result = document_manager.get_document_info(str(test_doc_path))
        
        assert result.status == ResponseStatus.SUCCESS
        assert result.data['table_count'] == 1
        assert len(result.data['tables']) == 1
        assert result.data['tables'][0]['rows'] == 3
        assert result.data['tables'][0]['columns'] == 3

    @pytest.mark.unit
    def test_close_document(self, document_manager, test_doc_path):
        """Test closing a document."""
        # Open document
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        
        # Close document
        result = document_manager.close_document(str(test_doc_path))
        
        assert result.status == ResponseStatus.SUCCESS
        assert "Document closed" in result.message

    @pytest.mark.unit
    def test_close_nonexistent_document(self, document_manager, test_doc_path):
        """Test closing a document that wasn't loaded."""
        result = document_manager.close_document(str(test_doc_path))
        
        assert result.status == ResponseStatus.WARNING
        assert "Document not loaded" in result.message

    @pytest.mark.unit
    def test_list_loaded_documents(self, document_manager, test_doc_path, temp_dir):
        """Test listing loaded documents."""
        # Initially no documents
        result = document_manager.list_loaded_documents()
        assert result.data['count'] == 0
        
        # Load some documents
        doc2_path = temp_dir / "doc2.docx"
        document_manager.open_document(str(test_doc_path), create_if_not_exists=True)
        document_manager.open_document(str(doc2_path), create_if_not_exists=True)
        
        # Check loaded documents
        result = document_manager.list_loaded_documents()
        assert result.status == ResponseStatus.SUCCESS
        assert result.data['count'] == 2
        assert str(test_doc_path) in result.data['loaded_documents']
        assert str(doc2_path) in result.data['loaded_documents']
