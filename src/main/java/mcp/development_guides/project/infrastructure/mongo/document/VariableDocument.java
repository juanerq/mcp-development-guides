package mcp.development_guides.project.infrastructure.mongo.document;

import lombok.AllArgsConstructor;
import lombok.Data;
import mcp.development_guides.project.domain.model.VariableType;
import org.springframework.data.annotation.Id;
import org.springframework.data.mongodb.core.mapping.Document;

@Document("variables")
@Data
@AllArgsConstructor
public class VariableDocument {
    @Id
    private String id;
    private String name;
    private String description;
    private VariableType type;
    private int row;
    private int column;
}
