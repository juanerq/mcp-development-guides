package mcp.development_guides.project.domain.model;

import lombok.Builder;

@Builder
public record Variable(
        String id,
        String name,
        String description,
        VariableType type,
        int row,
        int column
) {
    public CellPosition toCellPosition() {
        return new CellPosition(row, column);
    }

    public boolean isValid() {
        return id != null && !id.trim().isEmpty() &&
                name != null && !name.trim().isEmpty() &&
                description != null && !description.trim().isEmpty() &&
                type != null && isValidType(type) &&
                row >= 0 && column >= 0;
    }

    private boolean isValidType(VariableType type) {
        for (VariableType validType : VariableType.values()) {
            if (validType == type) {
                return true;
            }
        }
        return false;
    }
}
