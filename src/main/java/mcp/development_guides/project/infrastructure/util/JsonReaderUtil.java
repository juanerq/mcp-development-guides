package mcp.development_guides.project.infrastructure.util;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import mcp.development_guides.project.domain.model.Variable;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;
import org.springframework.stereotype.Component;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Collections;
import java.util.List;

/**
 * Utility class for reading JSON configuration files
 */
@Component
public class JsonReaderUtil {

    private final ObjectMapper objectMapper;

    public JsonReaderUtil() {
        this.objectMapper = new ObjectMapper();
    }

    /**
     * Read variables from the JSON file located in the database/data directory
     */
    public List<Variable> readVariables() {
        try {
            // First try to read from classpath (when running as JAR)
            try {
                Resource resource = new ClassPathResource("mcp/development_guides/project/database/data/variables.json");
                if (resource.exists()) {
                    try (InputStream inputStream = resource.getInputStream()) {
                        return objectMapper.readValue(inputStream, new TypeReference<List<Variable>>() {});
                    }
                }
            } catch (Exception e) {
                // If classpath reading fails, try file system path
            }

            // Try to read from file system (during development)
            Path filePath = Paths.get("src/main/java/mcp/development_guides/project/database/data/variables.json");
            if (Files.exists(filePath)) {
                return objectMapper.readValue(filePath.toFile(), new TypeReference<List<Variable>>() {});
            }

            // If neither works, return empty list
            return Collections.emptyList();

        } catch (IOException e) {
            throw new RuntimeException("Error reading variables.json file", e);
        }
    }
}
