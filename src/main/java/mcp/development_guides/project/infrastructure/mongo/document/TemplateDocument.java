package mcp.development_guides.project.infrastructure.mongo.document;

import lombok.AllArgsConstructor;
import lombok.Data;
import org.springframework.data.annotation.Id;
import org.springframework.data.mongodb.core.mapping.Document;

import java.util.List;

@Document("templates")
@Data
@AllArgsConstructor
public class TemplateDocument {
    @Id
    private String id;
    private String name;
    private String description;
    private String path;
    private String type;
    private List<VariableDocument> variables;
}
