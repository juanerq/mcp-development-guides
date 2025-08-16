package mcp.development_guides.project.domain.model;

import lombok.Builder;

import java.util.List;

@Builder
public record Template(
        String id,
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
        return id != null && !id.trim().isEmpty() &&
                name != null && !name.trim().isEmpty() &&
                description != null && !description.trim().isEmpty() &&
                path != null && !path.trim().isEmpty() &&
                type != null && !type.trim().isEmpty() &&
                variables != null && !variables.isEmpty();
    }
}
