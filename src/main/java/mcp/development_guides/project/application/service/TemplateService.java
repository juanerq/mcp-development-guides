package mcp.development_guides.project.application.service;

import com.fasterxml.jackson.core.type.TypeReference;
import mcp.development_guides.project.domain.model.Template;
import mcp.development_guides.project.infrastructure.util.JsonReaderUtil;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import java.nio.file.Paths;
import java.util.List;

@Service
public class TemplateService {
    private final JsonReaderUtil jsonReaderUtil;

    @Value("{app.data.file.templates}")
    private String templatesFileName;

    @Value("${app.data.directory:data}")
    private String dataDirectory;

    public TemplateService(JsonReaderUtil jsonReaderUtil) {
        this.jsonReaderUtil = jsonReaderUtil;
    }

    public List<Template> getAllTemplates() {
        String filePath = Paths.get(dataDirectory, templatesFileName).toString();
        TypeReference<List<Template>> typeReference = new TypeReference<List<Template>>() {};
        return jsonReaderUtil.readFile(filePath, typeReference);
    }
}
