package mcp.development_guides.project.infrastructure.excel.reader;

import mcp.development_guides.project.domain.model.CellPosition;
import mcp.development_guides.project.domain.model.ExcelCellData;
import mcp.development_guides.project.domain.model.ExcelRange;
import mcp.development_guides.project.infrastructure.excel.core.ExcelDataConverter;
import mcp.development_guides.project.infrastructure.excel.core.ExcelFileHandler;
import org.apache.poi.ss.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

/**
 * Lector especializado en celdas individuales y rangos de celdas
 */
@Component
public class ExcelCellReader {

    @Autowired
    private ExcelFileHandler fileHandler;

    @Autowired
    private ExcelDataConverter dataConverter;

    /**
     * Lee el valor de una celda específica
     */
    public String readCellValue(String filePath, String sheetName, int row, int column) {
        return fileHandler.executeWithWorkbook(filePath,
            String.format("Reading cell [%d,%d] from sheet '%s'", row, column, sheetName),
            (workbook, path) -> {
                Sheet sheet = fileHandler.getSheetByName(workbook, sheetName);
                return getCellValue(sheet, row, column);
            });
    }

    /**
     * Lee el valor de una celda específica usando CellPosition
     */
    public String readCellValue(String filePath, String sheetName, CellPosition position) {
        return readCellValue(filePath, sheetName, position.row(), position.column());
    }

    /**
     * Lee información detallada de una celda específica
     */
    public ExcelCellData readCellData(String filePath, String sheetName, int row, int column) {
        return fileHandler.executeWithWorkbook(filePath,
            String.format("Reading cell data [%d,%d] from sheet '%s'", row, column, sheetName),
            (workbook, path) -> {
                Sheet sheet = fileHandler.getSheetByName(workbook, sheetName);
                Row sheetRow = sheet.getRow(row);
                if (sheetRow == null) {
                    return new ExcelCellData(row, column, "", "BLANK",
                        new CellPosition(row, column).toExcelNotation());
                }
                Cell cell = sheetRow.getCell(column);
                if (cell == null) {
                    return new ExcelCellData(row, column, "", "BLANK",
                        new CellPosition(row, column).toExcelNotation());
                }
                return dataConverter.getCellData(cell);
            });
    }

    /**
     * Lee información detallada de una celda usando CellPosition
     */
    public ExcelCellData readCellData(String filePath, String sheetName, CellPosition position) {
        return readCellData(filePath, sheetName, position.row(), position.column());
    }

    /**
     * Lee un rango de celdas usando ExcelRange
     */
    public Object[][] readRange(String filePath, String sheetName, ExcelRange range) {
        return readRange(filePath, sheetName,
            range.startPosition().row(), range.startPosition().column(),
            range.endPosition().row(), range.endPosition().column());
    }

    /**
     * Lee un rango de celdas
     */
    public Object[][] readRange(String filePath, String sheetName, int startRow, int startColumn, int endRow, int endColumn) {
        return fileHandler.executeWithWorkbook(filePath,
            String.format("Reading range [%d,%d to %d,%d] from sheet '%s'", startRow, startColumn, endRow, endColumn, sheetName),
            (workbook, path) -> {
                Sheet sheet = fileHandler.getSheetByName(workbook, sheetName);

                int rowCount = endRow - startRow + 1;
                int colCount = endColumn - startColumn + 1;
                Object[][] result = new Object[rowCount][colCount];

                for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
                    for (int colIndex = startColumn; colIndex <= endColumn; colIndex++) {
                        String value = getCellValue(sheet, rowIndex, colIndex);
                        result[rowIndex - startRow][colIndex - startColumn] = value;
                    }
                }

                return result;
            });
    }

    /**
     * Lee una fila completa
     */
    public String[] readRow(String filePath, String sheetName, int rowIndex) {
        return fileHandler.executeWithWorkbook(filePath,
            String.format("Reading row %d from sheet '%s'", rowIndex, sheetName),
            (workbook, path) -> {
                Sheet sheet = fileHandler.getSheetByName(workbook, sheetName);
                Row row = sheet.getRow(rowIndex);

                if (row == null) {
                    return new String[0];
                }

                int lastColumn = row.getLastCellNum();
                String[] result = new String[lastColumn];

                for (int colIndex = 0; colIndex < lastColumn; colIndex++) {
                    result[colIndex] = getCellValue(sheet, rowIndex, colIndex);
                }

                return result;
            });
    }

    /**
     * Lee una columna completa
     */
    public String[] readColumn(String filePath, String sheetName, int columnIndex) {
        return fileHandler.executeWithWorkbook(filePath,
            String.format("Reading column %d from sheet '%s'", columnIndex, sheetName),
            (workbook, path) -> {
                Sheet sheet = fileHandler.getSheetByName(workbook, sheetName);
                int lastRow = sheet.getLastRowNum();
                String[] result = new String[lastRow + 1];

                for (int rowIndex = 0; rowIndex <= lastRow; rowIndex++) {
                    result[rowIndex] = getCellValue(sheet, rowIndex, columnIndex);
                }

                return result;
            });
    }

    /**
     * Obtiene el valor de una celda específica dentro de una hoja
     */
    private String getCellValue(Sheet sheet, int rowIndex, int columnIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            return "";
        }

        Cell cell = row.getCell(columnIndex);
        if (cell == null) {
            return "";
        }

        return dataConverter.getCellValueAsString(cell);
    }
}
