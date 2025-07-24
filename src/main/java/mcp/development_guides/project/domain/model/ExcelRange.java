package mcp.development_guides.project.domain.model;

/**
 * Representa un rango de celdas en Excel
 */
public record ExcelRange(
        CellPosition startPosition,
        CellPosition endPosition
) {
    /**
     * Constructor con validación para asegurar que el rango sea válido
     */
    public ExcelRange {
        if (startPosition.row() > endPosition.row() ||
            startPosition.column() > endPosition.column()) {
            throw new IllegalArgumentException("Start position must be before or equal to end position");
        }
    }

    /**
     * Constructor alternativo usando coordenadas directas
     */
    public ExcelRange(int startRow, int startColumn, int endRow, int endColumn) {
        this(new CellPosition(startRow, startColumn), new CellPosition(endRow, endColumn));
    }

    /**
     * Retorna el número de filas en el rango
     */
    public int getRowCount() {
        return endPosition.row() - startPosition.row() + 1;
    }

    /**
     * Retorna el número de columnas en el rango
     */
    public int getColumnCount() {
        return endPosition.column() - startPosition.column() + 1;
    }

    /**
     * Verifica si una posición está dentro del rango
     */
    public boolean contains(CellPosition position) {
        return position.row() >= startPosition.row() &&
               position.row() <= endPosition.row() &&
               position.column() >= startPosition.column() &&
               position.column() <= endPosition.column();
    }

    /**
     * Retorna la representación del rango en notación Excel (ej: "A1:C3")
     */
    public String toExcelNotation() {
        return startPosition.toExcelNotation() + ":" + endPosition.toExcelNotation();
    }
}
