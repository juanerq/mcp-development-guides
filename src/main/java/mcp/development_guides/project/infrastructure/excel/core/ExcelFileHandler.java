package mcp.development_guides.project.infrastructure.excel.core;

import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Component;

/**
 * Manejador principal de archivos Excel - Lógica común y fundamental
 */
@Component
public class ExcelFileHandler {

    /**
     * Interface funcional para operaciones con workbooks
     */
    @FunctionalInterface
    public interface WorkbookOperation<T> {
        T execute(Workbook workbook, String filePath) throws Exception;
    }

    /**
     * Ejecuta una operación con un workbook y maneja automáticamente el cierre de recursos
     */
    public <T> T executeWithWorkbook(String filePath, String operation, WorkbookOperation<T> workbookOperation) {
        try (Workbook workbook = WorkbookFactory.create(new java.io.File(filePath))) {
            System.out.println("📖 " + operation + " from: " + filePath);
            T result = workbookOperation.execute(workbook, filePath);
            System.out.println("✅ Operation completed successfully");
            return result;
        } catch (Exception e) {
            System.err.println("❌ Error during operation: " + operation);
            System.err.println("   File: " + filePath);
            System.err.println("   Error: " + e.getMessage());
            throw new RuntimeException("Failed to execute Excel operation: " + operation, e);
        }
    }

    /**
     * Obtiene una hoja por nombre con validación
     */
    public Sheet getSheetByName(Workbook workbook, String sheetName) {
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            throw new IllegalArgumentException("Sheet '" + sheetName + "' not found in workbook");
        }
        return sheet;
    }

    /**
     * Obtiene una hoja por índice con validación
     */
    public Sheet getSheetByIndex(Workbook workbook, int sheetIndex) {
        if (sheetIndex < 0 || sheetIndex >= workbook.getNumberOfSheets()) {
            throw new IllegalArgumentException("Sheet index " + sheetIndex + " is out of range. Available sheets: " + workbook.getNumberOfSheets());
        }
        return workbook.getSheetAt(sheetIndex);
    }

    /**
     * Valida que un archivo Excel existe y es accesible
     */
    public boolean validateExcelFile(String filePath) {
        try {
            java.io.File file = new java.io.File(filePath);
            if (!file.exists()) {
                System.err.println("❌ File does not exist: " + filePath);
                return false;
            }
            if (!file.canRead()) {
                System.err.println("❌ Cannot read file: " + filePath);
                return false;
            }
            return true;
        } catch (Exception e) {
            System.err.println("❌ Error validating file: " + e.getMessage());
            return false;
        }
    }
}
