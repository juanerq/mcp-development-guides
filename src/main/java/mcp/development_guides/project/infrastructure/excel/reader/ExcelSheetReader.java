package mcp.development_guides.project.infrastructure.excel.reader;

import mcp.development_guides.project.domain.model.ExcelCellData;
import mcp.development_guides.project.domain.model.ExcelSheetData;
import mcp.development_guides.project.domain.model.ExcelSheetInfo;
import mcp.development_guides.project.infrastructure.excel.core.ExcelDataConverter;
import mcp.development_guides.project.infrastructure.excel.core.ExcelFileHandler;
import org.apache.poi.ss.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import java.util.*;

/**
 * Lector especializado en hojas completas de Excel
 */
@Component
public class ExcelSheetReader {

    @Autowired
    private ExcelFileHandler fileHandler;

    @Autowired
    private ExcelDataConverter dataConverter;

    /**
     * Lee una hoja específica por nombre (compatibilidad con versión anterior)
     */
    public Map<String, Object> readSheet(String filePath, String sheetName) {
        return fileHandler.executeWithWorkbook(filePath, "Reading sheet '" + sheetName + "'", (workbook, path) -> {
            Sheet sheet = fileHandler.getSheetByName(workbook, sheetName);
            return loadSheetData(sheet, sheet.getWorkbook().getSheetIndex(sheet));
        });
    }

    /**
     * Lee una hoja específica por índice (compatibilidad con versión anterior)
     */
    public Map<String, Object> readSheetByIndex(String filePath, int sheetIndex) {
        return fileHandler.executeWithWorkbook(filePath, "Reading sheet at index " + sheetIndex, (workbook, path) -> {
            Sheet sheet = fileHandler.getSheetByIndex(workbook, sheetIndex);
            return loadSheetData(sheet, sheetIndex);
        });
    }

    /**
     * Lee una hoja específica por nombre retornando un ExcelSheetData (RECOMENDADO)
     */
    public ExcelSheetData readSheetData(String filePath, String sheetName) {
        return fileHandler.executeWithWorkbook(filePath, "Reading sheet '" + sheetName + "'", (workbook, path) -> {
            Sheet sheet = fileHandler.getSheetByName(workbook, sheetName);
            return loadSheetDataAsRecord(sheet, sheet.getWorkbook().getSheetIndex(sheet));
        });
    }

    /**
     * Lee una hoja específica por índice retornando un ExcelSheetData (RECOMENDADO)
     */
    public ExcelSheetData readSheetDataByIndex(String filePath, int sheetIndex) {
        return fileHandler.executeWithWorkbook(filePath, "Reading sheet at index " + sheetIndex, (workbook, path) -> {
            Sheet sheet = fileHandler.getSheetByIndex(workbook, sheetIndex);
            return loadSheetDataAsRecord(sheet, sheetIndex);
        });
    }

    /**
     * Obtiene información básica de todas las hojas (compatibilidad)
     */
    public List<Map<String, Object>> getSheetsSummary(String filePath) {
        return fileHandler.executeWithWorkbook(filePath, "Getting sheets summary", (workbook, path) -> {
            List<Map<String, Object>> sheets = new ArrayList<>();

            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                Map<String, Object> sheetInfo = new HashMap<>();
                sheetInfo.put("name", sheet.getSheetName());
                sheetInfo.put("index", i);
                sheetInfo.put("rowCount", sheet.getLastRowNum() + 1);
                sheetInfo.put("columnCount", getMaxColumnCount(sheet));
                sheets.add(sheetInfo);
            }

            return sheets;
        });
    }

    /**
     * Obtiene información básica de todas las hojas usando records (RECOMENDADO)
     */
    public List<ExcelSheetInfo> getSheetsSummaryAsRecords(String filePath) {
        return fileHandler.executeWithWorkbook(filePath, "Getting sheets summary", (workbook, path) -> {
            List<ExcelSheetInfo> sheets = new ArrayList<>();

            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                int rowCount = sheet.getLastRowNum() + 1;
                int columnCount = getMaxColumnCount(sheet);
                boolean hasData = rowCount > 0 && columnCount > 0;

                sheets.add(new ExcelSheetInfo(
                        sheet.getSheetName(),
                        i,
                        rowCount,
                        columnCount,
                        hasData
                ));
            }

            return sheets;
        });
    }

    /**
     * Obtiene solo los nombres de las hojas
     */
    public List<String> getSheetNames(String filePath) {
        return fileHandler.executeWithWorkbook(filePath, "Getting sheet names", (workbook, path) -> {
            List<String> sheetNames = new ArrayList<>();
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                sheetNames.add(workbook.getSheetAt(i).getSheetName());
            }
            return sheetNames;
        });
    }

    // MÉTODOS PRIVADOS DE SOPORTE

    /**
     * Carga todos los datos de una hoja (compatibilidad)
     */
    private Map<String, Object> loadSheetData(Sheet sheet, int index) {
        Map<String, Object> sheetData = new HashMap<>();
        sheetData.put("name", sheet.getSheetName());
        sheetData.put("index", index);

        List<ExcelCellData> cells = new ArrayList<>();
        int rowCount = 0;
        int columnCount = 0;

        for (Row row : sheet) {
            rowCount = Math.max(rowCount, row.getRowNum() + 1);
            for (Cell cell : row) {
                columnCount = Math.max(columnCount, cell.getColumnIndex() + 1);
                cells.add(dataConverter.getCellData(cell));
            }
        }

        sheetData.put("cells", cells);
        sheetData.put("rowCount", rowCount);
        sheetData.put("columnCount", columnCount);

        return sheetData;
    }

    /**
     * Carga todos los datos de una hoja como ExcelSheetData
     */
    private ExcelSheetData loadSheetDataAsRecord(Sheet sheet, int index) {
        List<List<ExcelCellData>> rows = new ArrayList<>();
        int rowCount = 0;
        int columnCount = 0;

        // Primero calculamos las dimensiones
        for (Row row : sheet) {
            rowCount = Math.max(rowCount, row.getRowNum() + 1);
            for (Cell cell : row) {
                columnCount = Math.max(columnCount, cell.getColumnIndex() + 1);
            }
        }

        // Inicializamos la estructura de filas
        for (int i = 0; i < rowCount; i++) {
            rows.add(new ArrayList<>());
        }

        // Llenamos con datos de celdas
        for (Row row : sheet) {
            List<ExcelCellData> rowData = rows.get(row.getRowNum());

            // Inicializamos celdas vacías hasta la columna necesaria
            while (rowData.size() < columnCount) {
                int currentColumn = rowData.size();
                rowData.add(new ExcelCellData(row.getRowNum(), currentColumn, "", "BLANK",
                        String.format("%c%d", 'A' + currentColumn, row.getRowNum() + 1)));
            }

            // Llenamos las celdas con datos
            for (Cell cell : row) {
                ExcelCellData cellData = dataConverter.getCellData(cell);
                rowData.set(cell.getColumnIndex(), cellData);
            }
        }

        boolean hasData = rowCount > 0 && columnCount > 0;
        ExcelSheetInfo info = new ExcelSheetInfo(
                sheet.getSheetName(),
                index,
                rowCount,
                columnCount,
                hasData
        );

        return new ExcelSheetData(info, rows);
    }

    /**
     * Obtiene el número máximo de columnas en una hoja
     */
    private int getMaxColumnCount(Sheet sheet) {
        int maxColumns = 0;
        for (Row row : sheet) {
            maxColumns = Math.max(maxColumns, row.getLastCellNum());
        }
        return maxColumns;
    }
}
