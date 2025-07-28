package mcp.development_guides.project.infrastructure.excel.specialized;

import mcp.development_guides.project.domain.model.CellModification;
import mcp.development_guides.project.domain.model.CellPosition;
import mcp.development_guides.project.domain.model.ExcelCellData;
import mcp.development_guides.project.domain.model.ExcelRange;
import mcp.development_guides.project.domain.model.ExcelSheetData;
import mcp.development_guides.project.domain.model.ExcelSheetInfo;
import mcp.development_guides.project.domain.model.Variable;
import mcp.development_guides.project.application.service.VariableService;
import mcp.development_guides.project.infrastructure.excel.core.ExcelFileHandler;
import mcp.development_guides.project.infrastructure.excel.reader.ExcelCellReader;
import mcp.development_guides.project.infrastructure.excel.reader.ExcelSheetReader;
import mcp.development_guides.project.infrastructure.excel.writer.ExcelCellWriter;
import mcp.development_guides.project.infrastructure.excel.writer.ExcelSheetWriter;
import mcp.development_guides.project.infrastructure.excel.writer.ExcelFormatWriter;
import mcp.development_guides.project.infrastructure.excel.writer.ExcelStructureEditor;
import mcp.development_guides.project.infrastructure.excel.writer.ExcelFileWriter;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.springframework.ai.tool.annotation.Tool;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.awt.Color;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Servicio principal MCP para Excel - Lógica especializada para casos específicos
 * Actúa como fachada que coordina todas las operaciones Excel usando la nueva estructura organizativa
 */
@Service
public class ExcelMCPService {

    // DEPENDENCIAS CORE
    @Autowired
    private ExcelFileHandler fileHandler;

    // DEPENDENCIAS READER
    @Autowired
    private ExcelCellReader cellReader;

    @Autowired
    private ExcelSheetReader sheetReader;

    // DEPENDENCIAS WRITER (AMPLIADAS)
    @Autowired
    private ExcelCellWriter cellWriter;

    @Autowired
    private ExcelSheetWriter sheetWriter;

    @Autowired
    private ExcelFormatWriter formatWriter;

    @Autowired
    private ExcelStructureEditor structureEditor;

    @Autowired
    private ExcelFileWriter fileWriter;

    // DEPENDENCIAS VARIABLES
    @Autowired
    private VariableService variableService;

    // ==================== HERRAMIENTAS DE VARIABLES Y CONFIGURACIÓN ====================

    @Tool(name = "read_variables", description = "Read all variables from the JSON configuration file")
    public List<Variable> readVariables() {
        return variableService.getAllVariables();
    }

    @Tool(name = "get_variable_by_name", description = "Get a specific variable by its name")
    public Variable getVariableByName(String name) {
        return variableService.findVariableByName(name).orElse(null);
    }

    @Tool(name = "validate_variables_config", description = "Validate all variables in the configuration")
    public Map<String, Object> validateVariablesConfig() {
        Map<String, Object> result = new HashMap<>();
        boolean isValid = variableService.validateAllVariables();
        List<Variable> invalidVariables = variableService.getInvalidVariables();

        result.put("isValid", isValid);
        result.put("totalVariables", variableService.countVariables());
        result.put("invalidVariables", invalidVariables);
        result.put("invalidCount", invalidVariables.size());

        return result;
    }

    // ==================== HERRAMIENTAS DE CARGA Y INFORMACIÓN ====================

    @Tool(name = "excel_load_file", description = "Load and read an Excel file returning its complete structure")
    public Map<String, Object> loadExcelFile(String filePath) {
        return fileHandler.executeWithWorkbook(filePath, "Loading Excel file", (workbook, path) -> {
            Map<String, Object> result = new HashMap<>();
            result.put("fileName", new java.io.File(filePath).getName());
            result.put("filePath", filePath);
            result.put("numberOfSheets", workbook.getNumberOfSheets());
            result.put("sheets", sheetReader.getSheetsSummaryAsRecords(filePath));
            return result;
        });
    }

    @Tool(name = "excel_get_file_info", description = "Get basic information about an Excel file")
    public Map<String, Object> getFileInfo(String filePath) {
        return fileHandler.executeWithWorkbook(filePath, "Getting file info", (workbook, path) -> {
            Map<String, Object> info = new HashMap<>();
            info.put("fileName", new java.io.File(filePath).getName());
            info.put("filePath", filePath);
            info.put("fileSize", new java.io.File(filePath).length());
            info.put("numberOfSheets", workbook.getNumberOfSheets());

            // Extract sheet names directly from the workbook to avoid nested operations
            List<String> sheetNames = new ArrayList<>();
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                sheetNames.add(workbook.getSheetAt(i).getSheetName());
            }
            info.put("sheetNames", sheetNames);

            return info;
        });
    }

    // ==================== HERRAMIENTAS DE LECTURA ====================

    @Tool(name = "excel_read_sheet", description = "Read a specific sheet from an Excel file using modern record-based structure")
    public ExcelSheetData readSheetData(String filePath, String sheetName) {
        return sheetReader.readSheetData(filePath, sheetName);
    }

    @Tool(name = "excel_read_sheet_by_index", description = "Read a specific sheet from an Excel file by index using modern record-based structure")
    public ExcelSheetData readSheetDataByIndex(String filePath, int sheetIndex) {
        return sheetReader.readSheetDataByIndex(filePath, sheetIndex);
    }

    @Tool(name = "excel_read_cell", description = "Read a specific cell value from an Excel file")
    public String readCellValue(String filePath, String sheetName, int row, int column) {
        return cellReader.readCellValue(filePath, sheetName, row, column);
    }

    @Tool(name = "excel_read_cell_data", description = "Read detailed cell information from an Excel file")
    public ExcelCellData readCellData(String filePath, String sheetName, int row, int column) {
        return cellReader.readCellData(filePath, sheetName, row, column);
    }

    @Tool(name = "excel_read_range", description = "Read a range of cells from an Excel file")
    public Object[][] readRange(String filePath, String sheetName, int startRow, int startColumn, int endRow, int endColumn) {
        return cellReader.readRange(filePath, sheetName, startRow, startColumn, endRow, endColumn);
    }

    @Tool(name = "excel_read_row", description = "Read a complete row from an Excel file")
    public String[] readRow(String filePath, String sheetName, int rowIndex) {
        return cellReader.readRow(filePath, sheetName, rowIndex);
    }

    @Tool(name = "excel_read_column", description = "Read a complete column from an Excel file")
    public String[] readColumn(String filePath, String sheetName, int columnIndex) {
        return cellReader.readColumn(filePath, sheetName, columnIndex);
    }

    @Tool(name = "excel_get_sheet_names", description = "Get all sheet names from an Excel file")
    public List<String> getSheetNames(String filePath) {
        return sheetReader.getSheetNames(filePath);
    }

    @Tool(name = "excel_get_sheets_summary", description = "Get summary information of all sheets using modern record structure")
    public List<ExcelSheetInfo> getSheetsSummary(String filePath) {
        return sheetReader.getSheetsSummaryAsRecords(filePath);
    }

    // ==================== HERRAMIENTAS DE ESCRITURA DE CELDAS ====================

    @Tool(name = "excel_modify_cells", description = "Modify multiple cells in a sheet with different types of content (text, numbers, formulas, booleans) in a single operation")
    public boolean modifyCells(String filePath, String sheetName, List<CellModification> modifications) {
        return cellWriter.modifyCells(filePath, sheetName, modifications);
    }

    // ==================== HERRAMIENTAS DE FORMATO Y ESTILO ====================

    @Tool(name = "excel_format_text", description = "Apply text formatting (bold, italic, color) to a cell")
    public boolean formatCellText(String filePath, String sheetName, int row, int column,
                                 boolean bold, boolean italic, String textColorHex) {
        Color textColor = textColorHex != null ? Color.decode(textColorHex) : null;
        return formatWriter.formatText(filePath, sheetName, row, column, bold, italic, textColor);
    }

    @Tool(name = "excel_set_background_color", description = "Set background color of a cell")
    public boolean setCellBackgroundColor(String filePath, String sheetName, int row, int column, String colorHex) {
        Color backgroundColor = Color.decode(colorHex);
        return formatWriter.setBackgroundColor(filePath, sheetName, row, column, backgroundColor);
    }

    @Tool(name = "excel_set_borders", description = "Set borders around a cell")
    public boolean setCellBorders(String filePath, String sheetName, int row, int column,
                                 String borderStyle, String borderColorHex) {
        BorderStyle style = BorderStyle.valueOf(borderStyle.toUpperCase());
        Color borderColor = borderColorHex != null ? Color.decode(borderColorHex) : null;
        return formatWriter.setBorders(filePath, sheetName, row, column, style, borderColor);
    }

    @Tool(name = "excel_set_number_format", description = "Set number format for a cell (e.g., currency, percentage)")
    public boolean setCellNumberFormat(String filePath, String sheetName, int row, int column, String formatPattern) {
        return formatWriter.setNumberFormat(filePath, sheetName, row, column, formatPattern);
    }

    @Tool(name = "excel_set_alignment", description = "Set text alignment for a cell")
    public boolean setCellAlignment(String filePath, String sheetName, int row, int column,
                                   String horizontal, String vertical) {
        HorizontalAlignment hAlign = horizontal != null ? HorizontalAlignment.valueOf(horizontal.toUpperCase()) : null;
        VerticalAlignment vAlign = vertical != null ? VerticalAlignment.valueOf(vertical.toUpperCase()) : null;
        return formatWriter.setAlignment(filePath, sheetName, row, column, hAlign, vAlign);
    }

    // ==================== HERRAMIENTAS DE ESTRUCTURA ====================

    @Tool(name = "excel_insert_row", description = "Insert a new row at the specified position")
    public boolean insertRow(String filePath, String sheetName, int rowIndex) {
        return structureEditor.insertRow(filePath, sheetName, rowIndex);
    }

    @Tool(name = "excel_delete_row", description = "Delete a row at the specified position")
    public boolean deleteRow(String filePath, String sheetName, int rowIndex) {
        return structureEditor.deleteRow(filePath, sheetName, rowIndex);
    }

    @Tool(name = "excel_insert_column", description = "Insert a new column at the specified position")
    public boolean insertColumn(String filePath, String sheetName, int columnIndex) {
        return structureEditor.insertColumn(filePath, sheetName, columnIndex);
    }

    @Tool(name = "excel_delete_column", description = "Delete a column at the specified position")
    public boolean deleteColumn(String filePath, String sheetName, int columnIndex) {
        return structureEditor.deleteColumn(filePath, sheetName, columnIndex);
    }

    @Tool(name = "excel_copy_range", description = "Copy a range of cells to another location")
    public boolean copyRange(String filePath, String sheetName, int startRow, int startCol,
                            int endRow, int endCol, int targetRow, int targetCol) {
        ExcelRange sourceRange = new ExcelRange(startRow, startCol, endRow, endCol);
        return structureEditor.copyRange(filePath, sheetName, sourceRange, targetRow, targetCol);
    }

    @Tool(name = "excel_move_range", description = "Move a range of cells to another location")
    public boolean moveRange(String filePath, String sheetName, int startRow, int startCol,
                            int endRow, int endCol, int targetRow, int targetCol) {
        ExcelRange sourceRange = new ExcelRange(startRow, startCol, endRow, endCol);
        return structureEditor.moveRange(filePath, sheetName, sourceRange, targetRow, targetCol);
    }

    @Tool(name = "excel_merge_cells", description = "Merge cells in the specified range")
    public boolean mergeCells(String filePath, String sheetName, int startRow, int startCol, int endRow, int endCol) {
        ExcelRange range = new ExcelRange(startRow, startCol, endRow, endCol);
        return structureEditor.mergeCells(filePath, sheetName, range);
    }

    @Tool(name = "excel_unmerge_cells", description = "Unmerge cells in the specified range")
    public boolean unmergeCells(String filePath, String sheetName, int startRow, int startCol, int endRow, int endCol) {
        ExcelRange range = new ExcelRange(startRow, startCol, endRow, endCol);
        return structureEditor.unmergeCells(filePath, sheetName, range);
    }

    // ==================== HERRAMIENTAS DE GESTIÓN DE ARCHIVOS ====================

    @Tool(name = "excel_create_new_file", description = "Create a new Excel file")
    public boolean createNewExcelFile(String filePath) {
        return fileWriter.createNewFile(filePath);
    }

    @Tool(name = "excel_create_file_with_sheets", description = "Create a new Excel file with specific sheets")
    public boolean createNewFileWithSheets(String filePath, List<String> sheetNames) {
        return fileWriter.createNewFileWithSheets(filePath, sheetNames.toArray(new String[0]));
    }

    @Tool(name = "excel_copy_file", description = "Copy an Excel file to another location")
    public boolean copyExcelFile(String sourceFilePath, String targetFilePath) {
        return fileWriter.copyFile(sourceFilePath, targetFilePath);
    }

    @Tool(name = "excel_merge_files", description = "Merge multiple Excel files into one")
    public boolean mergeExcelFiles(List<String> sourceFilePaths, String targetFilePath) {
        return fileWriter.mergeFiles(sourceFilePaths.toArray(new String[0]), targetFilePath);
    }

    @Tool(name = "excel_split_file_by_sheets", description = "Split an Excel file into separate files by sheets")
    public boolean splitFileBySheets(String sourceFilePath, String outputDirectory) {
        return fileWriter.splitFileBySheets(sourceFilePath, outputDirectory);
    }

    @Tool(name = "excel_convert_to_csv", description = "Convert an Excel sheet to CSV format")
    public boolean convertSheetToCSV(String excelFilePath, String sheetName, String csvFilePath) {
        return fileWriter.convertToCSV(excelFilePath, sheetName, csvFilePath);
    }

    @Tool(name = "excel_protect_file", description = "Protect an Excel file with password")
    public boolean protectExcelFile(String filePath, String password) {
        return fileWriter.protectFile(filePath, password);
    }

    @Tool(name = "excel_clear_range", description = "Clear content from a range of cells")
    public boolean clearCellRange(String filePath, String sheetName, int startRow, int startCol, int endRow, int endCol) {
        ExcelRange range = new ExcelRange(startRow, startCol, endRow, endCol);
        return cellWriter.clearRange(filePath, sheetName, range);
    }

    // ==================== HERRAMIENTAS DE GESTIÓN DE HOJAS ====================

    @Tool(name = "excel_create_sheet", description = "Create a new sheet in an Excel file")
    public boolean createSheet(String filePath, String sheetName) {
        return sheetWriter.createSheet(filePath, sheetName);
    }

    @Tool(name = "excel_delete_sheet", description = "Delete a sheet from an Excel file")
    public boolean deleteSheet(String filePath, String sheetName) {
        return sheetWriter.deleteSheet(filePath, sheetName);
    }

    @Tool(name = "excel_rename_sheet", description = "Rename a sheet in an Excel file")
    public boolean renameSheet(String filePath, String oldName, String newName) {
        return sheetWriter.renameSheet(filePath, oldName, newName);
    }

    @Tool(name = "excel_copy_sheet", description = "Copy a complete sheet within the same Excel file")
    public boolean copySheet(String filePath, String sourceSheetName, String targetSheetName) {
        return sheetWriter.copySheet(filePath, sourceSheetName, targetSheetName);
    }

    @Tool(name = "excel_copy_sheet_between_files", description = "Copy a complete sheet from one Excel file to another")
    public boolean copySheetBetweenFiles(String sourceFilePath, String sourceSheetName,
                                        String targetFilePath, String targetSheetName) {
        return sheetWriter.copySheetBetweenFiles(sourceFilePath, sourceSheetName, targetFilePath, targetSheetName);
    }

    @Tool(name = "excel_clear_sheet", description = "Clear all content from a sheet in an Excel file")
    public boolean clearSheet(String filePath, String sheetName) {
        return sheetWriter.clearSheet(filePath, sheetName);
    }

    @Tool(name = "excel_write_rows", description = "Write multiple rows of data to a sheet in an Excel file")
    public boolean writeRows(String filePath, String sheetName, int startRow, List<List<Object>> data) {
        return sheetWriter.writeRows(filePath, sheetName, startRow, data);
    }

    // ==================== FUNCIONES ESPECIALIZADAS ESPECÍFICAS ====================

    @Tool(name = "excel_analyze_data_types", description = "Analyze and return data types for all cells in a sheet")
    public Map<String, Object> analyzeDataTypes(String filePath, String sheetName) {
        ExcelSheetData sheetData = sheetReader.readSheetData(filePath, sheetName);
        Map<String, Object> analysis = new HashMap<>();

        Map<String, Integer> typeCount = new HashMap<>();
        int totalCells = 0;

        for (List<ExcelCellData> row : sheetData.rows()) {
            for (ExcelCellData cell : row) {
                if (!cell.value().isEmpty()) {
                    typeCount.merge(cell.type(), 1, Integer::sum);
                    totalCells++;
                }
            }
        }

        analysis.put("sheetName", sheetName);
        analysis.put("totalCells", totalCells);
        analysis.put("typeDistribution", typeCount);
        analysis.put("sheetInfo", sheetData.info());

        return analysis;
    }

    @Tool(name = "excel_find_value", description = "Find all occurrences of a specific value in a sheet")
    public List<CellPosition> findValue(String filePath, String sheetName, String searchValue) {
        ExcelSheetData sheetData = sheetReader.readSheetData(filePath, sheetName);
        return sheetData.rows().stream()
            .flatMap(List::stream)
            .filter(cell -> cell.value().equals(searchValue))
            .map(cell -> new CellPosition(cell.row(), cell.column()))
            .toList();
    }

    @Tool(name = "excel_validate_file", description = "Validate if an Excel file exists and is accessible")
    public boolean validateExcelFile(String filePath) {
        return fileHandler.validateExcelFile(filePath);
    }
}
