package mcp.development_guides.project.infrastructure.mongo.repository;

import mcp.development_guides.project.domain.model.Template;
import mcp.development_guides.project.domain.repository.TemplateRepository;
import mcp.development_guides.project.infrastructure.mongo.mapper.TemplateMapper;
import org.springframework.stereotype.Repository;

import java.util.List;
import java.util.Optional;

@Repository
public class TemplateRepositoryMongo implements TemplateRepository {
    private final DataTemplateRepository repository;
    private final TemplateMapper mapper;

    public TemplateRepositoryMongo(DataTemplateRepository repository, TemplateMapper mapper) {
        this.repository = repository;
        this.mapper = mapper;
    }

    @Override
    public List<Template> findAll() {
        return repository.findAll().stream()
                .map(mapper::toDomain)
                .toList();
    }

    @Override
    public Optional<Template> findById(String id) {
        return repository.findById(id).map(mapper::toDomain);
    }

    @Override
    public Optional<Template> findByName(String id) {
        return repository.findByName(id)
                .map(mapper::toDomain);
    }

    @Override
    public Template save(Template template) {
        return mapper.toDomain(repository.save(mapper.toDocument(template)));
    }
}
