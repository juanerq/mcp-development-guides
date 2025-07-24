package mcp.development_guides.project.domain.model;

import java.util.List;

/**
 * Representa los datos completos de una hoja de Excel
 */
public record ExcelSheetData(
        ExcelSheetInfo info,
        List<List<ExcelCellData>> rows
) {
    /**
     * Constructor que valida los datos
     */
    public ExcelSheetData {
        if (info == null) {
            throw new IllegalArgumentException("Sheet info cannot be null");
        }
        if (rows == null) {
            throw new IllegalArgumentException("Rows cannot be null");
        }
    }

    /**
     * Obtiene una celda específica por posición
     */
    public ExcelCellData getCell(int row, int column) {
        if (row >= 0 && row < rows.size() &&
            column >= 0 && column < rows.get(row).size()) {
            return rows.get(row).get(column);
        }
        return null;
    }

    /**
     * Obtiene una celda específica por CellPosition
     */
    public ExcelCellData getCell(CellPosition position) {
        return getCell(position.row(), position.column());
    }
}
