package mcp.development_guides.project.infrastructure.mongo.repository;

import mcp.development_guides.project.infrastructure.mongo.document.TemplateDocument;
import org.springframework.data.mongodb.repository.MongoRepository;
import org.springframework.stereotype.Repository;

import java.util.Optional;

@Repository
public interface DataTemplateRepository extends MongoRepository<TemplateDocument, String> {
    Optional<TemplateDocument> findByName(String name);
}
