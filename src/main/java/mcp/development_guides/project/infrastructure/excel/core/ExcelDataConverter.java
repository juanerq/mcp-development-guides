package mcp.development_guides.project.infrastructure.excel.core;

import mcp.development_guides.project.domain.model.ExcelCellData;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Component;

/**
 * Convertidor de datos de celdas Excel - Lógica fundamental de conversión
 */
@Component
public class ExcelDataConverter {

    /**
     * Convierte el valor de una celda a String
     */
    public String getCellValueAsString(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    double numericValue = cell.getNumericCellValue();
                    // Si es un número entero, mostrarlo sin decimales
                    if (numericValue == Math.floor(numericValue)) {
                        return String.valueOf((long) numericValue);
                    } else {
                        return String.valueOf(numericValue);
                    }
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return evaluateFormula(cell);
            case BLANK:
            case _NONE:
            case ERROR:
            default:
                return "";
        }
    }

    /**
     * Convierte una celda a un ExcelCellData con información detallada
     */
    public ExcelCellData getCellData(Cell cell) {
        String value = getCellValueAsString(cell);
        String type = cell.getCellType().name();
        String address = cell.getAddress().formatAsString();
        int row = cell.getRowIndex();
        int column = cell.getColumnIndex();

        // Información adicional según el tipo
        switch (cell.getCellType()) {
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return new ExcelCellData(row, column, value, type, address, cell.getDateCellValue());
                } else {
                    return new ExcelCellData(row, column, value, type, address, cell.getNumericCellValue());
                }
            case FORMULA:
                return ExcelCellData.withFormula(row, column, value, type, address, cell.getCellFormula());
            default:
                return new ExcelCellData(row, column, value, type, address);
        }
    }

    /**
     * Evalúa una fórmula y retorna su resultado como String
     */
    private String evaluateFormula(Cell cell) {
        try {
            FormulaEvaluator evaluator = cell.getSheet().getWorkbook()
                .getCreationHelper().createFormulaEvaluator();
            CellValue cellValue = evaluator.evaluate(cell);

            switch (cellValue.getCellType()) {
                case NUMERIC:
                    double numValue = cellValue.getNumberValue();
                    if (numValue == Math.floor(numValue)) {
                        return String.valueOf((long) numValue);
                    } else {
                        return String.valueOf(numValue);
                    }
                case STRING:
                    return cellValue.getStringValue();
                case BOOLEAN:
                    return String.valueOf(cellValue.getBooleanValue());
                default:
                    return cell.getCellFormula();
            }
        } catch (Exception e) {
            // Si no se puede evaluar, retornar la fórmula
            return cell.getCellFormula();
        }
    }

    /**
     * Detecta automáticamente el tipo de datos de una celda
     */
    public String detectDataType(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return "TEXT";
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return "DATE";
                } else {
                    double value = cell.getNumericCellValue();
                    if (value == Math.floor(value)) {
                        return "INTEGER";
                    } else {
                        return "DECIMAL";
                    }
                }
            case BOOLEAN:
                return "BOOLEAN";
            case FORMULA:
                return "FORMULA";
            case BLANK:
                return "EMPTY";
            case ERROR:
                return "ERROR";
            default:
                return "UNKNOWN";
        }
    }
}
