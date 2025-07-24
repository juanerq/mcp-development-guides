package mcp.development_guides.project.domain.model;

/**
 * Representa información de una hoja de Excel
 */
public record ExcelSheetInfo(
        String name,
        int index,
        int rowCount,
        int columnCount,
        boolean hasData
) {
    /**
     * Constructor que valida los datos de entrada
     */
    public ExcelSheetInfo {
        if (name == null || name.trim().isEmpty()) {
            throw new IllegalArgumentException("Sheet name cannot be null or empty");
        }
        if (index < 0) {
            throw new IllegalArgumentException("Sheet index must be non-negative");
        }
        if (rowCount < 0 || columnCount < 0) {
            throw new IllegalArgumentException("Row and column counts must be non-negative");
        }
    }
}
