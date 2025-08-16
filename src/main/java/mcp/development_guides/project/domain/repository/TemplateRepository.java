package mcp.development_guides.project.domain.repository;

import mcp.development_guides.project.domain.model.Template;

import java.util.List;
import java.util.Optional;

public interface TemplateRepository {
    List<Template> findAll();
    Optional<Template> findById(String id);
    Optional<Template> findByName(String name);
    Template save(Template template);
}
