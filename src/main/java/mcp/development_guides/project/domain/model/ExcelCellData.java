package mcp.development_guides.project.domain.model;

import java.util.Date;

/**
 * Representa los datos de una celda Excel de forma inmutable y tipada
 */
public record ExcelCellData(
        int row,
        int column,
        String value,
        String type,
        String address,
        Boolean isDate,
        Date dateValue,
        Double numericValue,
        String formula
) {
    /**
     * Constructor para celdas simples sin datos adicionales
     */
    public ExcelCellData(int row, int column, String value, String type, String address) {
        this(row, column, value, type, address, null, null, null, null);
    }

    /**
     * Constructor para celdas numéricas
     */
    public ExcelCellData(int row, int column, String value, String type, String address, Double numericValue) {
        this(row, column, value, type, address, false, null, numericValue, null);
    }

    /**
     * Constructor para celdas de fecha
     */
    public ExcelCellData(int row, int column, String value, String type, String address, Date dateValue) {
        this(row, column, value, type, address, true, dateValue, null, null);
    }

    /**
     * Constructor para celdas con fórmula
     */
    public static ExcelCellData withFormula(int row, int column, String value, String type, String address, String formula) {
        return new ExcelCellData(row, column, value, type, address, null, null, null, formula);
    }
}
