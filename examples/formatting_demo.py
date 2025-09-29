#!/usr/bin/env python3
"""
DOCX-MCP Cell Formatting Demo

This example demonstrates the new cell formatting capabilities in Phase 2,
including text formatting, alignment, background colors, and borders.
"""

import sys
import os
from pathlib import Path

# Add the src directory to the path for development
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

from docx_mcp.core.document_manager import DocumentManager
from docx_mcp.operations.tables.table_operations import TableOperations
from docx_mcp.models.formatting import (
    TextFormat, CellAlignment, CellBorders, BorderProperties,
    HorizontalAlignment, VerticalAlignment, BorderStyle, BorderWidth,
    Colors, Fonts
)


def main():
    """Demonstrate cell formatting capabilities."""
    
    print("🎨 DOCX-MCP Cell Formatting Demo")
    print("=" * 50)
    
    # Initialize managers
    doc_manager = DocumentManager()
    table_ops = TableOperations(doc_manager)
    
    document_path = "formatting_demo.docx"
    
    try:
        # Step 1: Create document and basic table
        print("\n📄 Step 1: Creating document and table...")
        result = doc_manager.open_document(document_path, create_if_not_exists=True)
        if result.status.value == "success":
            print(f"✅ {result.message}")
        
        # Create a demo table
        headers = ["Product", "Price", "Stock", "Status"]
        result = table_ops.create_table(document_path, rows=5, cols=4, headers=headers)
        if result.status.value == "success":
            print("✅ Demo table created with headers")
        
        # Add sample data
        sample_data = [
            ["Widget A", "$29.99", "150", "In Stock"],
            ["Widget B", "$45.00", "75", "Low Stock"],
            ["Widget C", "$12.50", "0", "Out of Stock"],
            ["Widget D", "$89.99", "200", "In Stock"]
        ]
        
        for row_idx, row_data in enumerate(sample_data):
            for col_idx, value in enumerate(row_data):
                table_ops.set_cell_value(document_path, 0, row_idx + 1, col_idx, value)
        
        print("✅ Sample data added")
        
        # Step 2: Format header row
        print("\n🎨 Step 2: Formatting header row...")
        
        # Header formatting: Bold, white text on blue background
        for col_idx in range(4):
            # Text formatting
            result = table_ops.formatting.format_cell_text(
                document_path, 0, 0, col_idx,
                TextFormat(
                    font_family=Fonts.ARIAL,
                    font_size=12,
                    font_color=Colors.WHITE,
                    bold=True
                )
            )
            
            # Background color
            result = table_ops.formatting.format_cell_background(
                document_path, 0, 0, col_idx, Colors.BLUE
            )
            
            # Center alignment
            result = table_ops.formatting.format_cell_alignment(
                document_path, 0, 0, col_idx,
                CellAlignment(
                    horizontal=HorizontalAlignment.CENTER,
                    vertical=VerticalAlignment.MIDDLE
                )
            )
        
        print("✅ Header row formatted (blue background, white bold text, centered)")
        
        # Step 3: Format price column
        print("\n💰 Step 3: Formatting price column...")
        
        for row_idx in range(1, 5):  # Skip header
            # Right-align prices and make them green
            result = table_ops.formatting.format_cell_text(
                document_path, 0, row_idx, 1,  # Column 1 is price
                TextFormat(
                    font_family=Fonts.CALIBRI,
                    font_size=11,
                    font_color="006400",  # Dark green
                    bold=True
                )
            )
            
            result = table_ops.formatting.format_cell_alignment(
                document_path, 0, row_idx, 1,
                CellAlignment(horizontal=HorizontalAlignment.RIGHT)
            )
        
        print("✅ Price column formatted (right-aligned, green, bold)")
        
        # Step 4: Format stock column with conditional colors
        print("\n📦 Step 4: Conditional formatting for stock levels...")
        
        stock_values = ["150", "75", "0", "200"]
        for row_idx, stock_value in enumerate(stock_values, 1):
            stock_num = int(stock_value)
            
            if stock_num == 0:
                # Red background for out of stock
                bg_color = Colors.RED
                text_color = Colors.WHITE
            elif stock_num < 100:
                # Yellow background for low stock
                bg_color = Colors.YELLOW
                text_color = Colors.BLACK
            else:
                # Light green for good stock
                bg_color = "90EE90"  # Light green
                text_color = Colors.BLACK
            
            result = table_ops.formatting.format_cell_background(
                document_path, 0, row_idx, 2, bg_color
            )
            
            result = table_ops.formatting.format_cell_text(
                document_path, 0, row_idx, 2,
                TextFormat(
                    font_color=text_color,
                    bold=True
                )
            )
            
            result = table_ops.formatting.format_cell_alignment(
                document_path, 0, row_idx, 2,
                CellAlignment(horizontal=HorizontalAlignment.CENTER)
            )
        
        print("✅ Stock levels with conditional colors (red=out, yellow=low, green=good)")
        
        # Step 5: Format status column with borders
        print("\n🔲 Step 5: Adding borders to status column...")
        
        status_values = ["In Stock", "Low Stock", "Out of Stock", "In Stock"]
        for row_idx, status in enumerate(status_values, 1):
            # Different border styles based on status
            if "Out of Stock" in status:
                border_color = Colors.RED
                border_style = BorderStyle.DOUBLE
            elif "Low Stock" in status:
                border_color = Colors.YELLOW
                border_style = BorderStyle.DASHED
            else:
                border_color = Colors.GREEN
                border_style = BorderStyle.SOLID
            
            border_props = BorderProperties(
                style=border_style,
                width=BorderWidth.MEDIUM,
                color=border_color
            )
            
            borders = CellBorders(
                top=border_props,
                bottom=border_props,
                left=border_props,
                right=border_props
            )
            
            result = table_ops.formatting.format_cell_borders(
                document_path, 0, row_idx, 3, borders
            )
            
            # Center align status
            result = table_ops.formatting.format_cell_alignment(
                document_path, 0, row_idx, 3,
                CellAlignment(horizontal=HorizontalAlignment.CENTER)
            )
        
        print("✅ Status column with colored borders (solid=good, dashed=low, double=out)")
        
        # Step 6: Create a summary table with complete formatting
        print("\n📊 Step 6: Creating formatted summary table...")
        
        result = table_ops.create_table(document_path, rows=4, cols=2, headers=["Metric", "Value"])
        if result.status.value == "success":
            print("✅ Summary table created")
        
        # Add summary data
        summary_data = [
            ["Total Products", "4"],
            ["In Stock", "2"],
            ["Low Stock", "1"],
            ["Out of Stock", "1"]
        ]
        
        for row_idx, (metric, value) in enumerate(summary_data):
            table_ops.set_cell_value(document_path, 1, row_idx + 1, 0, metric)
            table_ops.set_cell_value(document_path, 1, row_idx + 1, 1, value)
        
        # Apply complete formatting to summary table
        for row_idx in range(4):  # All rows including header
            for col_idx in range(2):
                if row_idx == 0:  # Header row
                    formatting = {
                        "text_format": {
                            "font_family": Fonts.ARIAL,
                            "font_size": 12,
                            "font_color": Colors.WHITE,
                            "bold": True
                        },
                        "alignment": {
                            "horizontal": "center",
                            "vertical": "middle"
                        },
                        "background_color": "4472C4",  # Dark blue
                        "borders": {
                            "top": {"style": "solid", "width": "medium", "color": Colors.BLACK},
                            "bottom": {"style": "solid", "width": "medium", "color": Colors.BLACK},
                            "left": {"style": "solid", "width": "medium", "color": Colors.BLACK},
                            "right": {"style": "solid", "width": "medium", "color": Colors.BLACK}
                        }
                    }
                else:  # Data rows
                    formatting = {
                        "text_format": {
                            "font_family": Fonts.CALIBRI,
                            "font_size": 10,
                            "font_color": Colors.BLACK
                        },
                        "alignment": {
                            "horizontal": "center" if col_idx == 1 else "left",
                            "vertical": "middle"
                        },
                        "background_color": "F2F2F2",  # Light gray
                        "borders": {
                            "top": {"style": "solid", "width": "thin", "color": Colors.GRAY},
                            "bottom": {"style": "solid", "width": "thin", "color": Colors.GRAY},
                            "left": {"style": "solid", "width": "thin", "color": Colors.GRAY},
                            "right": {"style": "solid", "width": "thin", "color": Colors.GRAY}
                        }
                    }
                
                result = table_ops.formatting.format_cell_complete(
                    document_path, 1, row_idx, col_idx, formatting
                )
        
        print("✅ Summary table with complete formatting applied")
        
        # Step 7: Save the document
        print("\n💾 Step 7: Saving formatted document...")
        result = doc_manager.save_document(document_path)
        if result.status.value == "success":
            print(f"✅ Document saved: {result.message}")
        
        # Step 8: Document summary
        print("\n📊 Step 8: Final document summary...")
        result = doc_manager.get_document_info(document_path)
        if result.status.value == "success":
            info = result.data
            print("✅ Final document stats:")
            print(f"   📄 File: {info['file_path']}")
            print(f"   📊 Tables: {info['table_count']}")
            print(f"   📝 Paragraphs: {info['paragraph_count']}")
        
        print("\n🎉 Cell formatting demonstration completed!")
        print(f"📁 Check the formatted document: {Path(document_path).absolute()}")
        
        # Show what was demonstrated
        print("\n💡 Formatting features demonstrated:")
        print("   🎨 Text formatting (fonts, sizes, colors, styles)")
        print("   📐 Cell alignment (horizontal and vertical)")
        print("   🌈 Background colors (solid colors)")
        print("   🔲 Cell borders (styles, widths, colors)")
        print("   🎯 Conditional formatting based on data")
        print("   📋 Complete cell formatting (all options together)")
        
    except Exception as e:
        print(f"\n❌ Error during formatting demo: {str(e)}")
        import traceback
        traceback.print_exc()
    
    finally:
        print("\n" + "=" * 50)
        print("🚀 Phase 2 formatting capabilities are now available!")
        print("🔮 Next: Data import/export and search functionality!")


if __name__ == "__main__":
    main()
