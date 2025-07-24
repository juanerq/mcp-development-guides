package mcp.development_guides.project.infrastructure.excel.writer;

import mcp.development_guides.project.domain.model.ExcelRange;
import mcp.development_guides.project.infrastructure.excel.core.ExcelFileHandler;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import java.io.FileOutputStream;

/**
 * Editor especializado en estructura de Excel (insertar/eliminar filas/columnas, copiar/mover)
 */
@Component
public class ExcelStructureEditor {

    @Autowired
    private ExcelFileHandler fileHandler;

    /**
     * Inserta una nueva fila en la posición especificada
     */
    public boolean insertRow(String filePath, String sheetName, int rowIndex) {
        return modifyWorkbook(filePath,
            String.format("Inserting row at index %d in sheet '%s'", rowIndex, sheetName),
            (workbook, path) -> {
                Sheet sheet = fileHandler.getSheetByName(workbook, sheetName);
                
                // Desplazar filas existentes hacia abajo
                int lastRow = sheet.getLastRowNum();
                if (rowIndex <= lastRow) {
                    sheet.shiftRows(rowIndex, lastRow, 1);
                }
                
                // Crear la nueva fila
                sheet.createRow(rowIndex);
                return true;
            });
    }

    /**
     * Elimina una fila en la posición especificada
     */
    public boolean deleteRow(String filePath, String sheetName, int rowIndex) {
        return modifyWorkbook(filePath,
            String.format("Deleting row at index %d in sheet '%s'", rowIndex, sheetName),
            (workbook, path) -> {
                Sheet sheet = fileHandler.getSheetByName(workbook, sheetName);
                
                Row row = sheet.getRow(rowIndex);
                if (row != null) {
                    sheet.removeRow(row);
                    
                    // Desplazar filas hacia arriba
                    int lastRow = sheet.getLastRowNum();
                    if (rowIndex < lastRow) {
                        sheet.shiftRows(rowIndex + 1, lastRow, -1);
                    }
                }
                return true;
            });
    }

    /**
     * Inserta una nueva columna en la posición especificada
     */
    public boolean insertColumn(String filePath, String sheetName, int columnIndex) {
        return modifyWorkbook(filePath,
            String.format("Inserting column at index %d in sheet '%s'", columnIndex, sheetName),
            (workbook, path) -> {
                Sheet sheet = fileHandler.getSheetByName(workbook, sheetName);
                
                // Desplazar celdas existentes hacia la derecha
                for (Row row : sheet) {
                    for (int colIndex = row.getLastCellNum(); colIndex >= columnIndex; colIndex--) {
                        Cell oldCell = row.getCell(colIndex);
                        if (oldCell != null) {
                            Cell newCell = row.createCell(colIndex + 1);
                            copyCellValue(oldCell, newCell);
                            row.removeCell(oldCell);
                        }
                    }
                }
                return true;
            });
    }

    /**
     * Elimina una columna en la posición especificada
     */
    public boolean deleteColumn(String filePath, String sheetName, int columnIndex) {
        return modifyWorkbook(filePath,
            String.format("Deleting column at index %d in sheet '%s'", columnIndex, sheetName),
            (workbook, path) -> {
                Sheet sheet = fileHandler.getSheetByName(workbook, sheetName);
                
                // Eliminar celdas de la columna y desplazar hacia la izquierda
                for (Row row : sheet) {
                    Cell cellToDelete = row.getCell(columnIndex);
                    if (cellToDelete != null) {
                        row.removeCell(cellToDelete);
                    }
                    
                    // Desplazar celdas hacia la izquierda
                    for (int colIndex = columnIndex + 1; colIndex <= row.getLastCellNum(); colIndex++) {
                        Cell oldCell = row.getCell(colIndex);
                        if (oldCell != null) {
                            Cell newCell = row.createCell(colIndex - 1);
                            copyCellValue(oldCell, newCell);
                            row.removeCell(oldCell);
                        }
                    }
                }
                return true;
            });
    }

    /**
     * Copia un rango de celdas a otra ubicación
     */
    public boolean copyRange(String filePath, String sheetName, ExcelRange sourceRange, 
                            int targetRow, int targetColumn) {
        return modifyWorkbook(filePath,
            String.format("Copying range %s to [%d,%d] in sheet '%s'", 
                sourceRange.toExcelNotation(), targetRow, targetColumn, sheetName),
            (workbook, path) -> {
                Sheet sheet = fileHandler.getSheetByName(workbook, sheetName);
                
                int rowOffset = targetRow - sourceRange.startPosition().row();
                int colOffset = targetColumn - sourceRange.startPosition().column();
                
                for (int rowIndex = sourceRange.startPosition().row(); 
                     rowIndex <= sourceRange.endPosition().row(); rowIndex++) {
                    
                    Row sourceRow = sheet.getRow(rowIndex);
                    if (sourceRow != null) {
                        Row targetRowObj = getOrCreateRow(sheet, rowIndex + rowOffset);
                        
                        for (int colIndex = sourceRange.startPosition().column(); 
                             colIndex <= sourceRange.endPosition().column(); colIndex++) {
                            
                            Cell sourceCell = sourceRow.getCell(colIndex);
                            if (sourceCell != null) {
                                Cell targetCell = targetRowObj.createCell(colIndex + colOffset);
                                copyCellValue(sourceCell, targetCell);
                                targetCell.setCellStyle(sourceCell.getCellStyle());
                            }
                        }
                    }
                }
                return true;
            });
    }

    /**
     * Mueve un rango de celdas a otra ubicación
     */
    public boolean moveRange(String filePath, String sheetName, ExcelRange sourceRange, 
                            int targetRow, int targetColumn) {
        return modifyWorkbook(filePath,
            String.format("Moving range %s to [%d,%d] in sheet '%s'", 
                sourceRange.toExcelNotation(), targetRow, targetColumn, sheetName),
            (workbook, path) -> {
                // Primero copiar
                Sheet sheet = fileHandler.getSheetByName(workbook, sheetName);
                
                int rowOffset = targetRow - sourceRange.startPosition().row();
                int colOffset = targetColumn - sourceRange.startPosition().column();
                
                // Copiar datos
                for (int rowIndex = sourceRange.startPosition().row(); 
                     rowIndex <= sourceRange.endPosition().row(); rowIndex++) {
                    
                    Row sourceRow = sheet.getRow(rowIndex);
                    if (sourceRow != null) {
                        Row targetRowObj = getOrCreateRow(sheet, rowIndex + rowOffset);
                        
                        for (int colIndex = sourceRange.startPosition().column(); 
                             colIndex <= sourceRange.endPosition().column(); colIndex++) {
                            
                            Cell sourceCell = sourceRow.getCell(colIndex);
                            if (sourceCell != null) {
                                Cell targetCell = targetRowObj.createCell(colIndex + colOffset);
                                copyCellValue(sourceCell, targetCell);
                                targetCell.setCellStyle(sourceCell.getCellStyle());
                            }
                        }
                    }
                }
                
                // Luego limpiar origen
                clearRange(sheet, sourceRange);
                return true;
            });
    }

    /**
     * Combina celdas en un rango especificado
     */
    public boolean mergeCells(String filePath, String sheetName, ExcelRange range) {
        return modifyWorkbook(filePath,
            String.format("Merging cells in range %s of sheet '%s'", range.toExcelNotation(), sheetName),
            (workbook, path) -> {
                Sheet sheet = fileHandler.getSheetByName(workbook, sheetName);
                
                CellRangeAddress cellRangeAddress = new CellRangeAddress(
                    range.startPosition().row(),
                    range.endPosition().row(),
                    range.startPosition().column(),
                    range.endPosition().column()
                );
                
                sheet.addMergedRegion(cellRangeAddress);
                return true;
            });
    }

    /**
     * Deshace la combinación de celdas en un rango
     */
    public boolean unmergeCells(String filePath, String sheetName, ExcelRange range) {
        return modifyWorkbook(filePath,
            String.format("Unmerging cells in range %s of sheet '%s'", range.toExcelNotation(), sheetName),
            (workbook, path) -> {
                Sheet sheet = fileHandler.getSheetByName(workbook, sheetName);
                
                // Buscar y eliminar regiones combinadas que coincidan
                for (int i = sheet.getNumMergedRegions() - 1; i >= 0; i--) {
                    CellRangeAddress mergedRegion = sheet.getMergedRegion(i);
                    
                    if (mergedRegion.getFirstRow() == range.startPosition().row() &&
                        mergedRegion.getLastRow() == range.endPosition().row() &&
                        mergedRegion.getFirstColumn() == range.startPosition().column() &&
                        mergedRegion.getLastColumn() == range.endPosition().column()) {
                        
                        sheet.removeMergedRegion(i);
                        break;
                    }
                }
                return true;
            });
    }

    // MÉTODOS PRIVADOS DE SOPORTE

    private boolean modifyWorkbook(String filePath, String operation, ExcelFileHandler.WorkbookOperation<Boolean> workbookOperation) {
        try (Workbook workbook = WorkbookFactory.create(new java.io.File(filePath))) {
            System.out.println("🔧 " + operation);
            
            boolean result = workbookOperation.execute(workbook, filePath);
            
            try (FileOutputStream outputStream = new FileOutputStream(filePath)) {
                workbook.write(outputStream);
                System.out.println("💾 Structure changes saved successfully to: " + filePath);
            }
            
            return result;
        } catch (Exception e) {
            System.err.println("❌ Error during structure operation: " + operation);
            System.err.println("   File: " + filePath);
            System.err.println("   Error: " + e.getMessage());
            return false;
        }
    }

    private void copyCellValue(Cell sourceCell, Cell targetCell) {
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
            default -> targetCell.setBlank();
        }
    }

    private Row getOrCreateRow(Sheet sheet, int rowIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        return row;
    }

    private void clearRange(Sheet sheet, ExcelRange range) {
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
    }
}
