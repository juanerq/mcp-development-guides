package mcp.development_guides.project.infrastructure.excel.writer;

import mcp.development_guides.project.domain.model.CellModification;
import mcp.development_guides.project.domain.model.CellPosition;
import mcp.development_guides.project.domain.model.ExcelRange;
import mcp.development_guides.project.infrastructure.excel.core.ExcelFileHandler;
import org.apache.poi.ss.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Date;
import java.util.List;

/**
 * Editor especializado en celdas individuales
 */
@Component
public class ExcelCellWriter {

    @Autowired
    private ExcelFileHandler fileHandler;

    /**
     * Escribe un valor String en una celda específica
     */
    public boolean writeCellValue(String filePath, String sheetName, int row, int column, String value) {
        return modifyWorkbook(filePath,
            String.format("Writing value '%s' to cell [%d,%d] in sheet '%s'", value, row, column, sheetName),
            (workbook, path) -> {
                Sheet sheet = getOrCreateSheet(workbook, sheetName);
                Cell cell = getOrCreateCell(sheet, row, column);
                cell.setCellValue(value);
                return true;
            });
    }

    /**
     * Escribe un valor usando CellPosition
     */
    public boolean writeCellValue(String filePath, String sheetName, CellPosition position, String value) {
        return writeCellValue(filePath, sheetName, position.row(), position.column(), value);
    }

    /**
     * Escribe un valor numérico en una celda
     */
    public boolean writeCellNumber(String filePath, String sheetName, int row, int column, double value) {
        return modifyWorkbook(filePath,
            String.format("Writing number %f to cell [%d,%d] in sheet '%s'", value, row, column, sheetName),
            (workbook, path) -> {
                Sheet sheet = getOrCreateSheet(workbook, sheetName);
                Cell cell = getOrCreateCell(sheet, row, column);
                cell.setCellValue(value);
                return true;
            });
    }

    /**
     * Escribe una fecha en una celda
     */
    public boolean writeCellDate(String filePath, String sheetName, int row, int column, Date date) {
        return modifyWorkbook(filePath,
            String.format("Writing date %s to cell [%d,%d] in sheet '%s'", date, row, column, sheetName),
            (workbook, path) -> {
                Sheet sheet = getOrCreateSheet(workbook, sheetName);
                Cell cell = getOrCreateCell(sheet, row, column);
                cell.setCellValue(date);

                // Aplicar formato de fecha
                CellStyle dateStyle = workbook.createCellStyle();
                CreationHelper createHelper = workbook.getCreationHelper();
                dateStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd/mm/yyyy"));
                cell.setCellStyle(dateStyle);

                return true;
            });
    }

    /**
     * Escribe una fórmula en una celda
     */
    public boolean writeCellFormula(String filePath, String sheetName, int row, int column, String formula) {
        return modifyWorkbook(filePath,
            String.format("Writing formula '%s' to cell [%d,%d] in sheet '%s'", formula, row, column, sheetName),
            (workbook, path) -> {
                Sheet sheet = getOrCreateSheet(workbook, sheetName);
                Cell cell = getOrCreateCell(sheet, row, column);
                cell.setCellFormula(formula);
                return true;
            });
    }

    /**
     * Escribe un valor boolean en una celda
     */
    public boolean writeCellBoolean(String filePath, String sheetName, int row, int column, boolean value) {
        return modifyWorkbook(filePath,
            String.format("Writing boolean %b to cell [%d,%d] in sheet '%s'", value, row, column, sheetName),
            (workbook, path) -> {
                Sheet sheet = getOrCreateSheet(workbook, sheetName);
                Cell cell = getOrCreateCell(sheet, row, column);
                cell.setCellValue(value);
                return true;
            });
    }

    /**
     * Limpia el contenido de una celda
     */
    public boolean clearCell(String filePath, String sheetName, int row, int column) {
        return modifyWorkbook(filePath,
            String.format("Clearing cell [%d,%d] in sheet '%s'", row, column, sheetName),
            (workbook, path) -> {
                Sheet sheet = fileHandler.getSheetByName(workbook, sheetName);
                Row sheetRow = sheet.getRow(row);
                if (sheetRow != null) {
                    Cell cell = sheetRow.getCell(column);
                    if (cell != null) {
                        sheetRow.removeCell(cell);
                    }
                }
                return true;
            });
    }

    /**
     * Limpia un rango de celdas
     */
    public boolean clearRange(String filePath, String sheetName, ExcelRange range) {
        return modifyWorkbook(filePath,
            String.format("Clearing range %s in sheet '%s'", range.toExcelNotation(), sheetName),
            (workbook, path) -> {
                Sheet sheet = fileHandler.getSheetByName(workbook, sheetName);

                for (int rowIndex = range.startPosition().row(); rowIndex <= range.endPosition().row(); rowIndex++) {
                    Row row = sheet.getRow(rowIndex);
                    if (row != null) {
                        for (int colIndex = range.startPosition().column(); colIndex <= range.endPosition().column(); colIndex++) {
                            Cell cell = row.getCell(colIndex);
                            if (cell != null) {
                                row.removeCell(cell);
                            }
                        }
                    }
                }
                return true;
            });
    }

    /**
     * Modifica múltiples celdas en una sola operación
     */
    public boolean modifyCells(String filePath, String sheetName, List<CellModification> modifications) {
        return modifyWorkbook(filePath,
            String.format("Modifying %d cells in sheet '%s'", modifications.size(), sheetName),
            (workbook, path) -> {
                Sheet sheet = getOrCreateSheet(workbook, sheetName);

                for (CellModification modification : modifications) {
                    CellPosition position = modification.position();
                    Object value = modification.value();
                    CellModification.ModificationType type = modification.type();

                    switch (type) {
                        case TEXT:
                            Cell textCell = getOrCreateCell(sheet, position.row(), position.column());
                            textCell.setCellValue((String) value);
                            break;

                        case NUMBER:
                            Cell numberCell = getOrCreateCell(sheet, position.row(), position.column());
                            numberCell.setCellValue(((Number) value).doubleValue());
                            break;

                        case FORMULA:
                            Cell formulaCell = getOrCreateCell(sheet, position.row(), position.column());
                            formulaCell.setCellFormula((String) value);
                            break;

                        case BOOLEAN:
                            Cell booleanCell = getOrCreateCell(sheet, position.row(), position.column());
                            booleanCell.setCellValue((Boolean) value);
                            break;

                        case CLEAR:
                            Row sheetRow = sheet.getRow(position.row());
                            if (sheetRow != null) {
                                Cell cellToClear = sheetRow.getCell(position.column());
                                if (cellToClear != null) {
                                    sheetRow.removeCell(cellToClear);
                                }
                            }
                            break;
                    }
                }

                return true;
            });
    }

    // MÉTODOS PRIVADOS DE SOPORTE

    /**
     * Ejecuta una operación de modificación en un workbook y guarda los cambios
     */
    private boolean modifyWorkbook(String filePath, String operation, ExcelFileHandler.WorkbookOperation<Boolean> workbookOperation) {
        try {
            System.out.println("✏️ " + operation);

            // Crear una copia temporal del archivo para evitar conflictos
            File originalFile = new File(filePath);
            File tempFile = new File(filePath + ".tmp");

            boolean result;

            // Abrir el archivo original, hacer modificaciones, y guardar en temporal
            try (Workbook workbook = WorkbookFactory.create(originalFile)) {
                result = workbookOperation.execute(workbook, filePath);

                // Guardar en archivo temporal
                try (FileOutputStream outputStream = new FileOutputStream(tempFile)) {
                    workbook.write(outputStream);
                }
            }

            // Si todo salió bien, reemplazar el archivo original con el temporal
            if (result && tempFile.exists()) {
                if (originalFile.delete()) {
                    if (tempFile.renameTo(originalFile)) {
                        System.out.println("💾 Changes saved successfully to: " + filePath);
                        return true;
                    } else {
                        System.err.println("❌ Could not rename temporary file");
                        tempFile.delete(); // Limpiar archivo temporal
                        return false;
                    }
                } else {
                    System.err.println("❌ Could not delete original file");
                    tempFile.delete(); // Limpiar archivo temporal
                    return false;
                }
            } else {
                // Limpiar archivo temporal si algo salió mal
                if (tempFile.exists()) {
                    tempFile.delete();
                }
                return false;
            }

        } catch (Exception e) {
            System.err.println("❌ Error during operation: " + operation);
            System.err.println("   File: " + filePath);
            System.err.println("   Error: " + e.getMessage());

            // Limpiar archivo temporal si existe
            File tempFile = new File(filePath + ".tmp");
            if (tempFile.exists()) {
                tempFile.delete();
            }

            return false;
        }
    }

    /**
     * Obtiene una hoja existente o la crea si no existe
     */
    private Sheet getOrCreateSheet(Workbook workbook, String sheetName) {
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            sheet = workbook.createSheet(sheetName);
            System.out.println("📄 Created new sheet: " + sheetName);
        }
        return sheet;
    }

    /**
     * Obtiene una celda existente o la crea si no existe
     */
    private Cell getOrCreateCell(Sheet sheet, int rowIndex, int columnIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }

        Cell cell = row.getCell(columnIndex);
        if (cell == null) {
            cell = row.createCell(columnIndex);
        }

        return cell;
    }
}
