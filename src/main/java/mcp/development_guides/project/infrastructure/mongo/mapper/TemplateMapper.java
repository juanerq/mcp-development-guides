package mcp.development_guides.project.infrastructure.mongo.mapper;

import mcp.development_guides.project.domain.model.Template;
import mcp.development_guides.project.domain.model.Variable;
import mcp.development_guides.project.infrastructure.mongo.document.TemplateDocument;
import mcp.development_guides.project.infrastructure.mongo.document.VariableDocument;
import org.springframework.stereotype.Component;

import java.util.List;

@Component
public class TemplateMapper {
    public Template toDomain(TemplateDocument document) {
        if (document == null) {
            return null;
        }

        List<Variable> variables = document.getVariables() == null ? List.of() : document.getVariables()
                .stream()
                .map(variable -> Variable.builder()
                        .id(variable.getId())
                        .name(variable.getName())
                        .description(variable.getDescription())
                        .type(variable.getType())
                        .row(variable.getRow())
                        .column(variable.getColumn())
                        .build()
                ).toList();

        return Template.builder()
                .id(document.getId())
                .name(document.getName())
                .description(document.getDescription())
                .path(document.getPath())
                .type(document.getType())
                .variables(variables)
                .build();
    }

    public TemplateDocument toDocument(Template template) {
        if (template == null) {
            return null;
        }

        List<VariableDocument> variableDocuments = template.variables() == null ? List.of() : template.variables()
                .stream()
                .map(variable -> new VariableDocument(
                        variable.id(),
                        variable.name(),
                        variable.description(),
                        variable.type(),
                        variable.row(),
                        variable.column()
                )).toList();

        return new TemplateDocument(
                template.id(),
                template.name(),
                template.description(),
                template.path(),
                template.type(),
                variableDocuments
        );
    }
}
