# DOCX-MCP: Word Document MCP Server

A comprehensive Model Context Protocol (MCP) server for Microsoft Word document operations, built with FastMCP and python-docx. Currently focused on advanced table operations with plans to expand into full document manipulation capabilities.

## 🚀 Features

### Phase 1 - Core Table Operations ✅

**Document Management**
- Open/create Word documents
- Save documents with optional rename
- Get document information and metadata

**Table Structure Operations**
- Create tables with customizable dimensions
- Delete tables by index
- Add/remove rows and columns at any position
- Support for header rows

**Table Data Operations**
- Set/get individual cell values with optional styling
- Bulk table data retrieval (array, object, CSV formats)
- List all tables in document with metadata
- Comprehensive table structure and style analysis

### Phase 2.1 - Cell Formatting ✅ (New!)

**Text Formatting**
- Font family, size, and color customization
- Bold, italic, underline, strikethrough styling
- Subscript and superscript support

**Cell Alignment**
- Horizontal alignment (left, center, right, justify)
- Vertical alignment (top, middle, bottom)

**Visual Styling**
- Cell background colors (hex color support)
- Cell borders with customizable styles, widths, and colors
- Complete formatting (apply all options at once)

### Phase 2.8 - Table Structure Analysis ✅ (New!)

**Comprehensive Table Analysis**
- Complete table structure analysis with cell-by-cell details
- Automatic detection of merged cells and their ranges
- Full style extraction (fonts, colors, alignment, borders)
- Header row identification using intelligent heuristics
- Support for analyzing single tables or all tables in document

**LLM-Friendly Formatting**
- Enhanced cell operations that preserve existing styles
- Optional style application when setting cell values
- Detailed formatting information when reading cell values
- Perfect for maintaining document consistency during LLM operations

### Phase 2 - Advanced Table Features 🔄 (In Progress)

**Table Formatting & Styling**
- Cell formatting (bold, italic, font, color, alignment)
- Table borders and shading
- Row height and column width adjustment
- Table positioning and text wrapping

**Data Import/Export**
- CSV import to tables
- Excel data import
- JSON data mapping to tables
- Bulk data operations

**Table Search & Query**
- Search cell content across tables
- Filter table data by criteria
- Sort table rows by column values
- Table data validation

### Phase 3 - Extended Table Features 🔮 (Planned)

**Table Templates & Automation**
- Predefined table templates
- Table style libraries
- Automated table generation from data
- Table cloning and duplication

**Advanced Operations**
- Table merging and splitting
- Cross-table references
- Calculated fields and formulas
- Table relationship management

### Phase 4 - Document Operations 🔮 (Future)

**Content Management**
- Text insertion and formatting
- Paragraph operations
- Heading and outline management
- Document structure manipulation

**Media & Objects**
- Image insertion and positioning
- Shape and drawing objects
- Charts and graphs integration
- Hyperlinks and bookmarks

**Document Formatting**
- Page layout and margins
- Headers and footers
- Styles and themes
- Document properties and metadata

## 📦 Installation

### Prerequisites
- Python 3.8+
- pip package manager

### Install from Source

1. Clone the repository:
```bash
git clone <repository-url>
cd docx_table
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Install the package in development mode:
```bash
pip install -e .
```

## 🖥️ Usage

### As MCP Server

Run the MCP server with default STDIO transport:
```bash
python -m docx_mcp.server
```

> **✅ Independent Tool Design**: Each MCP tool now works independently without requiring document pre-loading. You can directly call any tool with a file path, and the document will be automatically loaded as needed. This makes the tools more suitable for AI model integration.

#### Transport Protocols

The server supports multiple transport protocols:

**STDIO (default)** - Standard input/output for direct integration:
```bash
python -m docx_mcp.server --transport stdio
```

**SSE (Server-Sent Events)** - HTTP-based streaming:
```bash
python -m docx_mcp.server --transport sse --host localhost --port 8000
```

**Streamable HTTP** - HTTP with streaming support:
```bash
python -m docx_mcp.server --transport streamable-http --host localhost --port 8000
```

#### Command Line Options

```bash
python -m docx_mcp.server --help
```

Available options:
- `--transport {stdio,sse,streamable-http}` - Transport protocol (default: stdio)
- `--host HOST` - Host to bind to for HTTP/SSE transports (default: localhost)
- `--port PORT` - Port to bind to for HTTP/SSE transports (default: 8000)
- `--no-banner` - Disable startup banner

### Direct Usage

```python
from docx_mcp.core.document_manager import DocumentManager
from docx_mcp.operations.tables.table_operations import TableOperations

# Initialize
doc_manager = DocumentManager()
table_ops = TableOperations(doc_manager)

# No need to pre-load documents! Tools work independently.
# Create table with headers (document auto-loaded if needed)
result = table_ops.create_table(
    "document.docx", 
    rows=3, 
    cols=4, 
    headers=["Name", "Age", "City", "Occupation"]
)

# Set cell value (document auto-loaded if not already cached)
result = table_ops.set_cell_value("document.docx", 0, 1, 0, "Alice")

# Get table data (document auto-loaded if needed)
result = table_ops.get_table_data("document.docx", 0, include_headers=True)

# Save document (document auto-loaded and saved)
result = doc_manager.save_document("document.docx")
```

> **🚀 Automatic Document Loading**: The new design automatically loads documents when needed, eliminating the need to explicitly call `open_document()` before using other operations. Documents are cached for performance, and you can still manually manage document loading if preferred.

## 🔧 Available MCP Tools

All tools accept JSON parameters and return JSON responses, making them compatible with language models.

### Document Operations
- `open_document(file_path, create_if_not_exists=True)` - Open or create a Word document (optional - tools auto-load)
- `save_document(file_path, save_as=None)` - Save a Word document (auto-loads if needed)
- `get_document_info(file_path)` - Get document information (auto-loads if needed)

### Table Structure Operations
- `create_table(file_path, rows, cols, position="end", paragraph_index=None, headers=None)` - Create a new table
- `delete_table(file_path, table_index)` - Delete a table
- `add_table_rows(file_path, table_index, count=1, position="end", row_index=None)` - Add rows to a table
- `add_table_columns(file_path, table_index, count=1, position="end", column_index=None)` - Add columns to a table
- `delete_table_rows(file_path, table_index, row_indices)` - Delete rows from a table

### Data Operations
- `set_cell_value(file_path, table_index, row_index, column_index, value, ...)` - Set cell value with optional formatting
- `get_cell_value(file_path, table_index, row_index, column_index, include_formatting=True)` - Get cell value with formatting info
- `get_table_data(file_path, table_index, include_headers=True, format="array")` - Get entire table data

### Query Operations
- `list_tables(file_path, include_summary=True)` - List all tables in document

### Table Structure Analysis Operations (New in Phase 2.8!)
- `analyze_table_structure(file_path, table_index)` - Comprehensive analysis of single table structure and styles
- `analyze_all_tables_structure(file_path)` - Analyze all tables in document with complete details

### Table Search Operations (New in Phase 2.7!)
- `search_table_content(file_path, query, search_mode="contains", case_sensitive=False, table_indices=None, max_results=None)` - Search for content within table cells
- `search_table_headers(file_path, query, search_mode="contains", case_sensitive=False)` - Search specifically in table headers

### Cell Formatting Operations (New in Phase 2.1!)
- `format_cell_text(file_path, table_index, row_index, column_index, ...)` - Format text in cell
- `format_cell_alignment(file_path, table_index, row_index, column_index, horizontal, vertical)` - Set cell alignment
- `format_cell_background(file_path, table_index, row_index, column_index, color)` - Set cell background color
- `format_cell_borders(file_path, table_index, row_index, column_index, ...)` - Set cell borders

### Example Language Model Usage

Language models can call these tools with JSON parameters. **No pre-loading required** - each tool call is independent:

```json
{
  "tool": "create_table",
  "parameters": {
    "file_path": "report.docx",
    "rows": 5,
    "cols": 3,
    "headers": ["Product", "Sales", "Growth"]
  }
}
```

```json
{
  "tool": "set_cell_value",
  "parameters": {
    "file_path": "report.docx",
    "table_index": 0,
    "row_index": 1,
    "column_index": 0,
    "value": "Widget A"
  }
}
```

**New! Cell Formatting Examples:**
```json
{
  "tool": "format_cell_text",
  "parameters": {
    "file_path": "report.docx",
    "table_index": 0,
    "row_index": 0,
    "column_index": 0,
    "font_family": "Arial",
    "font_size": 14,
    "font_color": "FF0000",
    "bold": true,
    "italic": true
  }
}
```

```json
{
  "tool": "format_cell_background",
  "parameters": {
    "file_path": "report.docx",
    "table_index": 0,
    "row_index": 0,
    "column_index": 0,
    "color": "FFFF00"
  }
}
```

**New! Enhanced Cell Operations Examples:**
```json
{
  "tool": "set_cell_value",
  "parameters": {
    "file_path": "report.docx",
    "table_index": 0,
    "row_index": 1,
    "column_index": 0,
    "value": "New Value",
    "font_family": "Arial",
    "font_size": 12,
    "bold": true,
    "background_color": "FFFF00",
    "preserve_existing_format": false
  }
}
```

```json
{
  "tool": "get_cell_value",
  "parameters": {
    "file_path": "report.docx",
    "table_index": 0,
    "row_index": 1,
    "column_index": 0,
    "include_formatting": true
  }
}
```

**New! Table Structure Analysis Examples:**
```json
{
  "tool": "analyze_table_structure",
  "parameters": {
    "file_path": "report.docx",
    "table_index": 0
  }
}
```

```json
{
  "tool": "analyze_all_tables_structure",
  "parameters": {
    "file_path": "report.docx"
  }
}
```

**New! Table Search Examples:**
```json
{
  "tool": "search_table_content",
  "parameters": {
    "file_path": "report.docx",
    "query": "Alice",
    "search_mode": "contains",
    "case_sensitive": false,
    "max_results": 10
  }
}
```

```json
{
  "tool": "search_table_headers",
  "parameters": {
    "file_path": "report.docx",
    "query": "Email",
    "search_mode": "exact"
  }
}
```

## 🧪 Testing

The project uses pytest for comprehensive testing with 93 test cases covering all functionality.

Run all tests:
```bash
pytest
```

Run tests with coverage:
```bash
pytest --cov=src/docx_mcp
```

Run specific test categories:
```bash
pytest -m unit          # Unit tests only
pytest -m integration   # Integration tests only
pytest tests/test_table_operations.py  # Specific test file
```

Run tests with verbose output:
```bash
pytest -v
```

## 📋 Development Roadmap

### Phase 2: Advanced Table Features (In Progress)
**Priority: High** - Completing table functionality before expanding scope

- [x] **Cell Formatting & Styling** ✅ **COMPLETED**
  - [x] Cell text formatting (bold, italic, underline, font family/size, color)
  - [x] Cell background colors with hex color support
  - [x] Cell borders with customizable styles, widths, and colors
  - [x] Text alignment (horizontal: left, center, right, justify)
  - [x] Vertical alignment (top, middle, bottom)
  - [x] Complete formatting (apply all options at once)
  - [ ] Row height and column width controls

- [ ] **Data Import/Export**
  - [ ] CSV file import to tables
  - [ ] Excel file data import (.xlsx)
  - [ ] JSON data structure mapping
  - [ ] Enhanced export with formatting preservation
  - [ ] Bulk cell data operations

- [x] **Table Search & Query** ✅ **COMPLETED**
  - [x] Search content across all table cells
  - [x] Search specifically in table headers
  - [x] Multiple search modes (exact, contains, regex)
  - [x] Case-sensitive and case-insensitive search
  - [x] Search specific tables or all tables
  - [x] Limit search results
  - [ ] Filter table rows by column criteria
  - [ ] Sort table data by column values
  - [ ] Find and replace in table content

- [x] **Table Structure Analysis** ✅ **COMPLETED**
  - [x] Comprehensive table structure analysis
  - [x] Cell-by-cell style and content extraction
  - [x] Automatic merged cell detection
  - [x] Header row identification heuristics
  - [x] Enhanced cell operations with style preservation
  - [x] LLM-friendly formatting information

### Phase 3: Extended Table Features
**Priority: Medium** - Advanced table manipulation

- [ ] **Table Templates & Automation**
  - Predefined table styles and layouts
  - Table template library system
  - Auto-generate tables from data schemas
  - Table duplication and cloning

- [ ] **Advanced Table Operations**
  - Merge and split table cells
  - Table-to-table data relationships
  - Calculated fields and basic formulas
  - Cross-reference table data

- [ ] **Performance & Optimization**
  - Batch operations for large tables
  - Memory optimization for big documents
  - Caching for frequently accessed data
  - Async operations support

### Phase 4: Document Content Operations
**Priority: Medium** - Expanding beyond tables

- [ ] **Text & Paragraph Management**
  - Insert and format text content
  - Paragraph styling and spacing
  - Bullet points and numbering
  - Text search and replace

- [ ] **Document Structure**
  - Heading hierarchy management
  - Table of contents generation
  - Section breaks and page layout
  - Document outline operations

### Phase 5: Media & Advanced Features
**Priority: Low** - Rich document features

- [ ] **Media Integration**
  - Image insertion and positioning
  - Charts and graphs from table data
  - Shape and drawing objects
  - Embedded object support

- [ ] **Advanced Document Features**
  - Headers and footers
  - Page numbering and layout
  - Document properties and metadata
  - Track changes and comments

### Phase 6: Enterprise Features
**Priority: Low** - Production-ready enhancements

- [ ] **Security & Compliance**
  - Document encryption support
  - Access control and permissions
  - Audit logging and tracking
  - Data validation and sanitization

- [ ] **Integration & Extensibility**
  - Plugin system architecture
  - External data source connections
  - API rate limiting and throttling
  - Multi-document operations

## 🔧 Architecture

### Project Structure
```
src/docx_mcp/
├── __init__.py                    # Package initialization
├── server.py                      # FastMCP server and tool definitions
├── core/
│   └── document_manager.py       # Core document operations
├── operations/
│   └── tables/
│       └── table_operations.py   # Table-specific operations
├── models/
│   ├── responses.py              # Response data models
│   └── tables.py                 # Table-related data models
└── utils/
    ├── exceptions.py             # Custom exception definitions
    └── validation.py             # Input validation utilities
```

### Key Design Principles
- **Modular Architecture**: Clear separation between document, table, and future operations
- **Type Safety**: Full type hints and Pydantic models for data validation
- **Error Handling**: Comprehensive exception handling with detailed error messages
- **Extensibility**: Easy to add new operation types and document features
- **Testing**: High test coverage with unit and integration tests

## 🐛 Error Handling

The library includes comprehensive error handling with custom exceptions:

- `DocumentNotFoundError` - Document file not found
- `TableNotFoundError` - Table not found in document
- `InvalidTableIndexError` - Invalid table index
- `InvalidCellPositionError` - Invalid cell position
- `TableOperationError` - Table operation failed
- `DataFormatError` - Invalid data format

## 🤝 Contributing

We welcome contributions! Here's how to get started:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Make your changes following our coding standards
4. Add tests for new functionality
5. Ensure all tests pass (`pytest`)
6. Commit your changes (`git commit -m 'Add amazing feature'`)
7. Push to the branch (`git push origin feature/amazing-feature`)
8. Open a Pull Request

### Development Guidelines
- Follow PEP 8 coding standards
- Add type hints to all functions
- Write comprehensive tests for new features
- Update documentation for API changes
- Ensure backward compatibility when possible

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 📞 Support

- **Issues**: Report bugs and request features on GitHub Issues
- **Documentation**: Check the examples/ directory for usage examples
- **Testing**: Run the test suite to verify functionality

## 🔗 Dependencies

- **Python 3.8+** - Core runtime
- **python-docx ≥ 1.1.0** - Word document manipulation
- **fastmcp ≥ 0.4.0** - MCP server framework
- **pytest** - Testing framework (development)

## 📊 Project Status

### ✅ Current Capabilities (Phase 1 + 2.1 + 2.7 + 2.8)
- **🧪 Test Coverage**: 93/93 tests passing (100%)
- **🛠️ MCP Tools**: 19 available tools (11 core + 4 formatting + 2 search + 2 analysis)
- **📦 Modules**: 8 core modules with clean architecture
- **🎨 Formatting**: Complete cell formatting support with style preservation
- **🔍 Search**: Comprehensive table search capabilities
- **🔬 Analysis**: Deep table structure and style analysis for LLMs
- **📚 Documentation**: Comprehensive API docs and examples

### 🚀 Recent Additions (Phase 2.8 + Tool Independence)
- ✅ **Independent Tool Design**: Each MCP tool works independently without requiring document pre-loading
- ✅ **Automatic Document Loading**: Documents are loaded automatically when needed and cached for performance
- ✅ **Table Structure Analysis**: Complete analysis of table structure, styles, and merged cells
- ✅ **Enhanced Cell Operations**: Set/get cell values with optional formatting preservation
- ✅ **Style Detection**: Automatic extraction of fonts, colors, alignment, borders
- ✅ **Merge Detection**: Identify and report merged cell regions
- ✅ **LLM Integration**: Perfect for maintaining document consistency during AI operations
- ✅ **Background Color Fix**: Resolved background color setting and extraction issues

### 🎯 Next Milestones
- **Phase 2.2**: Data import/export (CSV, Excel, JSON)
- **Phase 3**: Advanced table features and templates

---

**DOCX-MCP** - Making Word document automation accessible through the Model Context Protocol! 🚀