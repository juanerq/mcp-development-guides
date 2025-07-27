# Read Me First
The following was discovered as part of building this project:

* The original package name 'mcp.development-guides.project' is invalid and this project uses 'mcp.development_guides.project' instead.

# Excel MCP Service - Comprehensive Documentation

This project provides a Model Context Protocol (MCP) server for Excel file manipulation with comprehensive functionality for reading, writing, formatting, and managing Excel files.

## Project Information

* The original package name 'mcp.development-guides.project' is invalid and this project uses 'mcp.development_guides.project' instead.

## Available MCP Methods

The Excel MCP Service provides over 40 specialized methods organized into the following categories:

### 📂 File Loading and Information

#### `excel_load_file`
Load and read an Excel file returning its complete structure
- **Parameters**: `filePath` (String)
- **Returns**: Complete file structure with sheets summary

#### `excel_get_file_info`
Get basic information about an Excel file
- **Parameters**: `filePath` (String)
- **Returns**: File metadata including size, sheet count, and sheet names

#### `excel_validate_file`
Validate if an Excel file exists and is accessible
- **Parameters**: `filePath` (String)
- **Returns**: Boolean validation result

### 📖 Reading Operations

#### Sheet Reading
- **`excel_read_sheet`**: Read a specific sheet using modern record-based structure
- **`excel_read_sheet_by_index`**: Read sheet by index instead of name
- **`excel_get_sheet_names`**: Get all sheet names from an Excel file
- **`excel_get_sheets_summary`**: Get summary information of all sheets

#### Cell and Range Reading
- **`excel_read_cell`**: Read a specific cell value
- **`excel_read_cell_data`**: Read detailed cell information (type, formatting, etc.)
- **`excel_read_range`**: Read a range of cells
- **`excel_read_row`**: Read a complete row
- **`excel_read_column`**: Read a complete column

### ✏️ Writing and Modification

#### Cell Modification
- **`excel_modify_cells`**: Modify multiple cells with different content types (text, numbers, formulas, booleans) in a single operation
- **`excel_write_rows`**: Write multiple rows of data to a sheet
- **`excel_clear_range`**: Clear content from a range of cells

### 🎨 Formatting and Styling

#### Text Formatting
- **`excel_format_text`**: Apply text formatting (bold, italic, color)
- **`excel_set_background_color`**: Set cell background color
- **`excel_set_borders`**: Set borders around cells
- **`excel_set_number_format`**: Set number format (currency, percentage, etc.)
- **`excel_set_alignment`**: Set text alignment (horizontal/vertical)

### 🔧 Structure Operations

#### Row and Column Management
- **`excel_insert_row`**: Insert a new row at specified position
- **`excel_delete_row`**: Delete a row at specified position
- **`excel_insert_column`**: Insert a new column at specified position
- **`excel_delete_column`**: Delete a column at specified position

#### Range Operations
- **`excel_copy_range`**: Copy a range of cells to another location
- **`excel_move_range`**: Move a range of cells to another location
- **`excel_merge_cells`**: Merge cells in specified range
- **`excel_unmerge_cells`**: Unmerge cells in specified range

### 📁 File Management

#### File Creation and Copying
- **`excel_create_new_file`**: Create a new Excel file
- **`excel_create_file_with_sheets`**: Create a new Excel file with specific sheets
- **`excel_copy_file`**: Copy an Excel file to another location

#### File Operations
- **`excel_merge_files`**: Merge multiple Excel files into one
- **`excel_split_file_by_sheets`**: Split an Excel file into separate files by sheets
- **`excel_convert_to_csv`**: Convert an Excel sheet to CSV format
- **`excel_protect_file`**: Protect an Excel file with password

### 📋 Sheet Management

#### Sheet Operations
- **`excel_create_sheet`**: Create a new sheet in an Excel file
- **`excel_delete_sheet`**: Delete a sheet from an Excel file
- **`excel_rename_sheet`**: Rename a sheet in an Excel file
- **`excel_copy_sheet`**: Copy a complete sheet within the same Excel file
- **`excel_copy_sheet_between_files`**: Copy a sheet from one Excel file to another
- **`excel_clear_sheet`**: Clear all content from a sheet

### 🔍 Specialized Functions

#### Data Analysis
- **`excel_analyze_data_types`**: Analyze and return data types for all cells in a sheet
- **`excel_find_value`**: Find all occurrences of a specific value in a sheet

## Usage Examples

### Basic File Operations
```
# Load a file and get its structure
excel_load_file("/path/to/file.xlsx")

# Read a specific sheet
excel_read_sheet("/path/to/file.xlsx", "Sheet1")

# Read a cell value
excel_read_cell("/path/to/file.xlsx", "Sheet1", 0, 0)
```

### Data Modification
```
# Modify multiple cells at once
excel_modify_cells("/path/to/file.xlsx", "Sheet1", [
    {position: {row: 0, column: 0}, type: "TEXT", value: "Header"},
    {position: {row: 1, column: 0}, type: "NUMBER", value: 123}
])

# Format text with styling
excel_format_text("/path/to/file.xlsx", "Sheet1", 0, 0, true, false, "#FF0000")
```

### File Management
```
# Create a new file with sheets
excel_create_file_with_sheets("/path/to/new.xlsx", ["Data", "Summary"])

# Convert sheet to CSV
excel_convert_to_csv("/path/to/file.xlsx", "Sheet1", "/path/to/output.csv")
```

## Technical Architecture

The service is built with a modular architecture:

- **Core**: File handling and workbook operations
- **Reader**: Specialized reading operations for cells, sheets, and ranges
- **Writer**: Comprehensive writing capabilities including formatting and structure editing
- **Specialized**: High-level MCP service that coordinates all operations

## Getting Started

### Reference Documentation
For further reference, please consider the following sections:

* [Official Apache Maven documentation](https://maven.apache.org/guides/index.html)
* [Spring Boot Maven Plugin Reference Guide](https://docs.spring.io/spring-boot/3.5.3/maven-plugin)
* [Create an OCI image](https://docs.spring.io/spring-boot/3.5.3/maven-plugin/build-image.html)
* [Model Context Protocol Server](https://docs.spring.io/spring-ai/reference/api/mcp/mcp-server-boot-starter-docs.html)

### Maven Parent overrides

Due to Maven's design, elements are inherited from the parent POM to the project POM.
While most of the inheritance is fine, it also inherits unwanted elements like `<license>` and `<developers>` from the parent.
To prevent this, the project POM contains empty overrides for these elements.
If you manually switch to a different parent and actually want the inheritance, you need to remove those overrides.
