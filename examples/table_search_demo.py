#!/usr/bin/env python3
"""
Table Search Demo - Demonstrates table search functionality in DOCX-MCP

This example shows how to use the table search features to find content
within Word document tables.
"""

import os
import sys
from pathlib import Path

# Add the src directory to Python path
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from docx_mcp.core.document_manager import DocumentManager
from docx_mcp.operations.tables.table_operations import TableOperations


def create_sample_document(doc_path: str):
    """Create a sample document with tables for search demonstration."""
    print("Creating sample document with test data...")
    
    doc_manager = DocumentManager()
    table_ops = TableOperations(doc_manager)
    
    # Create document
    doc_manager.open_document(doc_path, create_if_not_exists=True)
    
    # Create first table - Employee Data
    print("Creating Employee Data table...")
    table_ops.create_table(
        doc_path, 
        rows=4, 
        cols=4, 
        headers=["Name", "Email", "Department", "Salary"]
    )
    
    # Add employee data
    employees = [
        ["Alice Johnson", "alice@company.com", "Engineering", "$75,000"],
        ["Bob Smith", "bob.smith@company.com", "Marketing", "$65,000"],
        ["Carol Davis", "carol.davis@company.com", "Engineering", "$80,000"]
    ]
    
    for i, employee in enumerate(employees):
        for j, value in enumerate(employee):
            table_ops.set_cell_value(doc_path, 0, i + 1, j, value)
    
    # Create second table - Project Status
    print("Creating Project Status table...")
    table_ops.create_table(
        doc_path, 
        rows=4, 
        cols=3, 
        headers=["Project", "Status", "Lead"]
    )
    
    projects = [
        ["Website Redesign", "In Progress", "Alice Johnson"],
        ["Mobile App", "Planning", "Bob Smith"],
        ["API Integration", "Completed", "Carol Davis"]
    ]
    
    for i, project in enumerate(projects):
        for j, value in enumerate(project):
            table_ops.set_cell_value(doc_path, 1, i + 1, j, value)
    
    # Save document
    doc_manager.save_document(doc_path)
    print(f"Sample document created: {doc_path}")
    return doc_manager, table_ops


def demonstrate_search_features(doc_path: str, table_ops: TableOperations):
    """Demonstrate various table search features."""
    print("\n" + "="*60)
    print("TABLE SEARCH DEMONSTRATIONS")
    print("="*60)
    
    # 1. Basic content search
    print("\n1. Basic Content Search:")
    print("   Searching for 'Alice' in all tables...")
    result = table_ops.search_table_content(doc_path, "Alice")
    if result.status.value == "success":
        print(f"   Found {result.data['total_matches']} matches:")
        for match in result.data['matches']:
            print(f"     - Table {match['table_index']}, Row {match['row_index']}, Col {match['column_index']}: '{match['cell_value']}'")
    
    # 2. Case-sensitive search
    print("\n2. Case-Sensitive Search:")
    print("   Searching for 'engineering' (case-sensitive)...")
    result = table_ops.search_table_content(doc_path, "engineering", case_sensitive=True)
    print(f"   Found {result.data['total_matches']} matches (case-sensitive)")
    
    print("   Searching for 'engineering' (case-insensitive)...")
    result = table_ops.search_table_content(doc_path, "engineering", case_sensitive=False)
    print(f"   Found {result.data['total_matches']} matches (case-insensitive)")
    
    # 3. Exact match search
    print("\n3. Exact Match Search:")
    print("   Searching for exact match 'Completed'...")
    result = table_ops.search_table_content(doc_path, "Completed", search_mode="exact")
    if result.status.value == "success":
        print(f"   Found {result.data['total_matches']} exact matches:")
        for match in result.data['matches']:
            print(f"     - Table {match['table_index']}, Row {match['row_index']}, Col {match['column_index']}: '{match['cell_value']}'")
    
    # 4. Regex search
    print("\n4. Regular Expression Search:")
    print("   Searching for email addresses using regex...")
    result = table_ops.search_table_content(doc_path, r'\w+@\w+\.\w+', search_mode="regex")
    if result.status.value == "success":
        print(f"   Found {result.data['total_matches']} email addresses:")
        for match in result.data['matches']:
            print(f"     - Table {match['table_index']}, Row {match['row_index']}, Col {match['column_index']}: '{match['cell_value']}'")
    
    # 5. Header search
    print("\n5. Header-Specific Search:")
    print("   Searching for 'Status' in table headers...")
    result = table_ops.search_table_headers(doc_path, "Status")
    if result.status.value == "success":
        print(f"   Found {result.data['total_matches']} header matches:")
        for match in result.data['matches']:
            print(f"     - Table {match['table_index']}, Header Column {match['column_index']}: '{match['cell_value']}'")
    
    # 6. Search specific table
    print("\n6. Search Specific Table:")
    print("   Searching for 'Johnson' only in table 0 (Employee Data)...")
    result = table_ops.search_table_content(doc_path, "Johnson", table_indices=[0])
    if result.status.value == "success":
        print(f"   Found {result.data['total_matches']} matches in table 0:")
        for match in result.data['matches']:
            print(f"     - Row {match['row_index']}, Col {match['column_index']}: '{match['cell_value']}'")
    
    # 7. Limited results
    print("\n7. Limited Search Results:")
    print("   Searching for any cell containing 'a' (limited to 3 results)...")
    result = table_ops.search_table_content(doc_path, "a", max_results=3)
    if result.status.value == "success":
        print(f"   Found {result.data['total_matches']} matches (limited to 3):")
        for match in result.data['matches']:
            print(f"     - Table {match['table_index']}, Row {match['row_index']}, Col {match['column_index']}: '{match['cell_value']}'")
    
    # 8. Search summary
    print("\n8. Search Summary Information:")
    print("   Searching for 'com' to show summary stats...")
    result = table_ops.search_table_content(doc_path, "com")
    if result.status.value == "success":
        summary = result.data['summary']
        print(f"   Tables searched: {result.data['tables_searched']}")
        print(f"   Tables with matches: {summary['tables_with_matches']}")
        print(f"   Total cells searched: {summary['total_cells_searched']}")
        print(f"   Matches per table: {summary['matches_per_table']}")


def main():
    """Main demonstration function."""
    print("DOCX-MCP Table Search Feature Demo")
    print("=" * 40)
    
    # Create demo document
    demo_doc = "table_search_demo.docx"
    
    try:
        # Create sample document with test data
        doc_manager, table_ops = create_sample_document(demo_doc)
        
        # Demonstrate search features
        demonstrate_search_features(demo_doc, table_ops)
        
        print(f"\n" + "="*60)
        print("DEMO COMPLETED SUCCESSFULLY!")
        print(f"Demo document saved as: {demo_doc}")
        print("You can open this document in Word to see the created tables.")
        print("="*60)
        
    except Exception as e:
        print(f"Error during demo: {e}")
        return 1
    
    finally:
        # Clean up - optionally remove demo file
        # if os.path.exists(demo_doc):
        #     os.remove(demo_doc)
        #     print(f"Cleaned up demo file: {demo_doc}")
        pass
    
    return 0


if __name__ == "__main__":
    sys.exit(main())
