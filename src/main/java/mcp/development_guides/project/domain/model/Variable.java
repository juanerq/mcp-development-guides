package mcp.development_guides.project.domain.model;

/**
 * Represents a variable definition from the JSON configuration
 */
public record Variable(
    String name,
    String description,
    String type,
    int row,
    int column
) {
    
    /**
     * Creates a CellPosition from this variable's row and column
     */
    public CellPosition toCellPosition() {
        return new CellPosition(row, column);
    }
    
    /**
     * Validates if the variable has all required fields
     */
    public boolean isValid() {
        return name != null && !name.trim().isEmpty() &&
               description != null && !description.trim().isEmpty() &&
               type != null && !type.trim().isEmpty() &&
               row >= 0 && column >= 0;
    }
}
