package mcp.development_guides.project.infrastructure.excel.writer;

import mcp.development_guides.project.domain.model.ExcelRange;
import mcp.development_guides.project.infrastructure.excel.core.ExcelFileHandler;
import org.apache.poi.ss.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import java.io.File;
import java.io.FileOutputStream;
import java.util.List;

/**
 * Editor especializado en operaciones a nivel de hoja completa
 */
@Component
public class ExcelSheetWriter {

    @Autowired
    private ExcelFileHandler fileHandler;

    /**
     * Crea una nueva hoja en el workbook
     */
    public boolean createSheet(String filePath, String sheetName) {
        return modifyWorkbook(filePath,
            String.format("Creating new sheet '%s'", sheetName),
            (workbook, path) -> {
                if (workbook.getSheet(sheetName) != null) {
                    System.out.println("⚠️ Sheet '" + sheetName + "' already exists");
                    return false;
                }
                workbook.createSheet(sheetName);
                return true;
            });
    }

    /**
     * Elimina una hoja del workbook
     */
    public boolean deleteSheet(String filePath, String sheetName) {
        return modifyWorkbook(filePath,
            String.format("Deleting sheet '%s'", sheetName),
            (workbook, path) -> {
                int sheetIndex = workbook.getSheetIndex(sheetName);
                if (sheetIndex == -1) {
                    System.out.println("⚠️ Sheet '" + sheetName + "' not found");
                    return false;
                }
                workbook.removeSheetAt(sheetIndex);
                return true;
            });
    }

    /**
     * Renombra una hoja existente
     */
    public boolean renameSheet(String filePath, String oldName, String newName) {
        return modifyWorkbook(filePath,
            String.format("Renaming sheet from '%s' to '%s'", oldName, newName),
            (workbook, path) -> {
                int sheetIndex = workbook.getSheetIndex(oldName);
                if (sheetIndex == -1) {
                    System.out.println("⚠️ Sheet '" + oldName + "' not found");
                    return false;
                }
                workbook.setSheetName(sheetIndex, newName);
                return true;
            });
    }

    /**
     * Escribe múltiples filas de datos
     */
    public boolean writeRows(String filePath, String sheetName, int startRow, List<List<Object>> data) {
        return modifyWorkbook(filePath,
            String.format("Writing %d rows to sheet '%s' starting at row %d", data.size(), sheetName, startRow),
            (workbook, path) -> {
                Sheet sheet = getOrCreateSheet(workbook, sheetName);

                for (int rowIndex = 0; rowIndex < data.size(); rowIndex++) {
                    List<Object> rowData = data.get(rowIndex);
                    Row row = getOrCreateRow(sheet, startRow + rowIndex);

                    for (int colIndex = 0; colIndex < rowData.size(); colIndex++) {
                        Cell cell = getOrCreateCell(row, colIndex);
                        setCellValue(cell, rowData.get(colIndex));
                    }
                }
                return true;
            });
    }

    /**
     * Escribe datos en un rango específico
     */
    public boolean writeRange(String filePath, String sheetName, ExcelRange range, Object[][] data) {
        return modifyWorkbook(filePath,
            String.format("Writing data to range %s in sheet '%s'", range.toExcelNotation(), sheetName),
            (workbook, path) -> {
                Sheet sheet = getOrCreateSheet(workbook, sheetName);

                int dataRows = data.length;
                int dataCols = dataRows > 0 ? data[0].length : 0;

                // Verificar que los datos caben en el rango
                if (dataRows > range.getRowCount() || dataCols > range.getColumnCount()) {
                    System.out.println("⚠️ Data size exceeds range dimensions");
                    return false;
                }

                for (int rowIndex = 0; rowIndex < dataRows; rowIndex++) {
                    Row row = getOrCreateRow(sheet, range.startPosition().row() + rowIndex);

                    for (int colIndex = 0; colIndex < dataCols; colIndex++) {
                        Cell cell = getOrCreateCell(row, range.startPosition().column() + colIndex);
                        setCellValue(cell, data[rowIndex][colIndex]);
                    }
                }
                return true;
            });
    }

    /**
     * Copia una hoja completa a otro archivo
     */
    public boolean copySheetToFile(String sourceFilePath, String sourceSheetName,
                                  String targetFilePath, String targetSheetName) {
        // Esta implementación es más compleja y requiere manejo de dos workbooks
        // Por ahora, devolvemos false como placeholder
        System.out.println("⚠️ Copy sheet functionality not yet implemented");
        return false;
    }

    /**
     * Copia una hoja completa dentro del mismo archivo
     */
    public boolean copySheet(String filePath, String sourceSheetName, String targetSheetName) {
        return modifyWorkbook(filePath,
            String.format("Copying sheet '%s' to '%s'", sourceSheetName, targetSheetName),
            (workbook, path) -> {
                Sheet sourceSheet = fileHandler.getSheetByName(workbook, sourceSheetName);
                Sheet targetSheet = workbook.createSheet(targetSheetName);
                copySheetData(sourceSheet, targetSheet);
                return true;
            });
    }

    /**
     * Copia una hoja completa entre diferentes archivos
     */
    public boolean copySheetBetweenFiles(String sourceFilePath, String sourceSheetName,
                                        String targetFilePath, String targetSheetName) {
        try {
            System.out.println("📋 Copying sheet '" + sourceSheetName + "' from " + sourceFilePath +
                             " to '" + targetSheetName + "' in " + targetFilePath);

            // Leer datos de la hoja origen
            return fileHandler.executeWithWorkbook(sourceFilePath, "Reading source sheet", (sourceWorkbook, sourcePath) -> {
                Sheet sourceSheet = fileHandler.getSheetByName(sourceWorkbook, sourceSheetName);

                // Escribir a archivo destino
                try (Workbook targetWorkbook = WorkbookFactory.create(new java.io.File(targetFilePath))) {
                    Sheet targetSheet = targetWorkbook.createSheet(targetSheetName);
                    copySheetData(sourceSheet, targetSheet);

                    // Guardar cambios en archivo destino
                    try (FileOutputStream outputStream = new FileOutputStream(targetFilePath)) {
                        targetWorkbook.write(outputStream);
                        System.out.println("✅ Sheet copied between files successfully");
                        return true;
                    }
                } catch (Exception e) {
                    System.err.println("❌ Error writing to target file: " + e.getMessage());
                    return false;
                }
            });
        } catch (Exception e) {
            System.err.println("❌ Error copying sheet between files: " + e.getMessage());
            return false;
        }
    }

    /**
     * Limpia todo el contenido de una hoja
     */
    public boolean clearSheet(String filePath, String sheetName) {
        return modifyWorkbook(filePath,
            String.format("Clearing all content from sheet '%s'", sheetName),
            (workbook, path) -> {
                Sheet sheet = fileHandler.getSheetByName(workbook, sheetName);

                // Eliminar todas las filas
                for (int i = sheet.getLastRowNum(); i >= 0; i--) {
                    Row row = sheet.getRow(i);
                    if (row != null) {
                        sheet.removeRow(row);
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
            System.out.println("📊 " + operation);

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
                        System.out.println("💾 Sheet operation completed successfully: " + filePath);
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
            System.err.println("❌ Error during sheet operation: " + operation);
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
     * Obtiene una fila existente o la crea si no existe
     */
    private Row getOrCreateRow(Sheet sheet, int rowIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        return row;
    }

    /**
     * Obtiene una celda existente o la crea si no existe
     */
    private Cell getOrCreateCell(Row row, int columnIndex) {
        Cell cell = row.getCell(columnIndex);
        if (cell == null) {
            cell = row.createCell(columnIndex);
        }
        return cell;
    }

    /**
     * Establece el valor de una celda según el tipo de dato
     */
    private void setCellValue(Cell cell, Object value) {
        switch (value) {
            case null -> cell.setBlank();
            case String s -> cell.setCellValue(s);
            case Number n -> cell.setCellValue(n.doubleValue());
            case Boolean b -> cell.setCellValue(b);
            case java.util.Date d -> cell.setCellValue(d);
            default -> cell.setCellValue(value.toString());
        }
    }

    /**
     * Copia los datos de una hoja a otra con mejor manejo de tipos
     */
    private void copySheetData(Sheet sourceSheet, Sheet targetSheet) {
        // Copiar configuraciones básicas de la hoja
        targetSheet.setDisplayGridlines(sourceSheet.isDisplayGridlines());

        for (int i = 0; i <= sourceSheet.getLastRowNum(); i++) {
            Row sourceRow = sourceSheet.getRow(i);
            if (sourceRow != null) {
                Row targetRow = targetSheet.createRow(i);

                // Copiar altura de fila
                targetRow.setHeight(sourceRow.getHeight());

                // Copiar celdas de la fila
                for (int j = 0; j < sourceRow.getLastCellNum(); j++) {
                    Cell sourceCell = sourceRow.getCell(j);
                    if (sourceCell != null) {
                        Cell targetCell = targetRow.createCell(j);
                        copyCellValueAndStyle(sourceCell, targetCell);
                    }
                }
            }
        }

        // Copiar anchos de columnas
        for (int i = 0; i < 256; i++) { // Máximo común de columnas
            try {
                int columnWidth = sourceSheet.getColumnWidth(i);
                if (columnWidth != sourceSheet.getDefaultColumnWidth() * 256) {
                    targetSheet.setColumnWidth(i, columnWidth);
                }
            } catch (Exception e) {
                // Ignorar si la columna no existe
                break;
            }
        }
    }

    /**
     * Copia el valor y estilo de una celda a otra
     */
    private void copyCellValueAndStyle(Cell sourceCell, Cell targetCell) {
        // Copiar valor según el tipo
        switch (sourceCell.getCellType()) {
            case STRING -> targetCell.setCellValue(sourceCell.getStringCellValue());
            case NUMERIC -> {
                if (DateUtil.isCellDateFormatted(sourceCell)) {
                    targetCell.setCellValue(sourceCell.getDateCellValue());
                } else {
                    targetCell.setCellValue(sourceCell.getNumericCellValue());
                }
            }
            case BOOLEAN -> targetCell.setCellValue(sourceCell.getBooleanCellValue());
            case FORMULA -> targetCell.setCellFormula(sourceCell.getCellFormula());
            case BLANK -> targetCell.setBlank();
            case ERROR -> targetCell.setCellErrorValue(sourceCell.getErrorCellValue());
        }

        // Copiar estilo si existe
        if (sourceCell.getCellStyle() != null) {
            targetCell.setCellStyle(sourceCell.getCellStyle());
        }
    }
}
