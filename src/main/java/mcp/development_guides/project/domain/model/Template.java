package mcp.development_guides.project.domain.model;

import java.util.List;

public record Template(
        String name,
        String description,
        String path,
        String type,
        List<Variable> variables
) {

    /**
     * Validates if the template has all required fields
     */
    public boolean isValid() {
        return name != null && !name.trim().isEmpty() &&
                description != null && !description.trim().isEmpty() &&
                path != null && !path.trim().isEmpty() &&
                type != null && !type.trim().isEmpty() &&
                variables != null && !variables.isEmpty();
    }
}
