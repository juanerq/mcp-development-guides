package mcp.development_guides.project.domain.model;

/**
 * Representa una modificación a realizar en una celda específica
 */
public record CellModification(
        CellPosition position,
        Object value,
        ModificationType type
) {
    public enum ModificationType {
        TEXT,
        NUMBER,
        FORMULA,
        BOOLEAN,
        CLEAR
    }

    /**
     * Constructor para modificación de texto
     */
    public static CellModification text(int row, int column, String value) {
        return new CellModification(new CellPosition(row, column), value, ModificationType.TEXT);
    }

    /**
     * Constructor para modificación numérica
     */
    public static CellModification number(int row, int column, double value) {
        return new CellModification(new CellPosition(row, column), value, ModificationType.NUMBER);
    }

    /**
     * Constructor para modificación de fórmula
     */
    public static CellModification formula(int row, int column, String formula) {
        return new CellModification(new CellPosition(row, column), formula, ModificationType.FORMULA);
    }

    /**
     * Constructor para modificación booleana
     */
    public static CellModification bool(int row, int column, boolean value) {
        return new CellModification(new CellPosition(row, column), value, ModificationType.BOOLEAN);
    }

    /**
     * Constructor para limpiar celda
     */
    public static CellModification clear(int row, int column) {
        return new CellModification(new CellPosition(row, column), null, ModificationType.CLEAR);
    }
}
