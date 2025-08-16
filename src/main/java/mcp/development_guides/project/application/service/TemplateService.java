package mcp.development_guides.project.application.service;

import mcp.development_guides.project.domain.model.Template;
import mcp.development_guides.project.domain.repository.TemplateRepository;
import org.springframework.stereotype.Service;

import java.util.List;

@Service
public class TemplateService {
    private final TemplateRepository templateRepository;

    public TemplateService(TemplateRepository templateRepository) {
        this.templateRepository = templateRepository;
    }

    public List<Template> getAllTemplates() {
        return templateRepository.findAll();
    }

    public Template getTemplateByName(String name) {
        return templateRepository.findByName(name)
                .orElseThrow(() -> new IllegalArgumentException("Template not found: " + name));
    }
}
