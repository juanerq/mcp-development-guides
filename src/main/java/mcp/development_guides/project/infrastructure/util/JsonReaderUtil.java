package mcp.development_guides.project.infrastructure.util;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

import java.io.IOException;
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
     * Read variables from the JSON file located in the database/data directory
     */
    public <T> T readFile(String path, TypeReference<T> typeReference) {
        try {
            Path filePath = Paths.get(path);
            if (Files.exists(filePath)) {
                return objectMapper.readValue(filePath.toFile(), typeReference);
            }

            return null;

        } catch (IOException e) {
            throw new RuntimeException("Error reading variables.json file", e);
        }
    }
}
