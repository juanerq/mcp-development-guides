package mcp.development_guides.project.domain.model;

/**
 * Representa una posición de celda en Excel (fila, columna)
 */
public record CellPosition(
        int row,
        int column
) {
    /**
     * Crea una posición de celda validando que los valores sean no negativos
     */
    public CellPosition {
        if (row < 0 || column < 0) {
            throw new IllegalArgumentException("Row and column must be non-negative");
        }
    }

    /**
     * Retorna una representación textual de la posición (ej: "A1", "B2")
     */
    public String toExcelNotation() {
        StringBuilder columnName = new StringBuilder();
        int col = column;
        while (col >= 0) {
            columnName.insert(0, (char) ('A' + col % 26));
            col = col / 26 - 1;
        }
        return columnName.toString() + (row + 1);
    }
}
