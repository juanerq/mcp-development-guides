package mcp.development_guides.project.infrastructure.util;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Component;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

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
     * Read JSON file from classpath or filesystem
     */
    public <T> T readFile(String path, TypeReference<T> typeReference) {
        try {
            try {
                ClassPathResource resource = new ClassPathResource(path);
                if (resource.exists()) {
                    try (InputStream inputStream = resource.getInputStream()) {
                        return objectMapper.readValue(inputStream, typeReference);
                    }
                }
            } catch (Exception ignore) {
            }

            Path filePath = Paths.get(path);
            if (Files.exists(filePath)) {
                return objectMapper.readValue(filePath.toFile(), typeReference);
            }

            throw new RuntimeException("Error file not found: " + filePath);

        } catch (IOException e) {
            throw new RuntimeException("Error reading JSON file: " + path, e);
        }
    }
}
