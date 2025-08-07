package mcp.development_guides.project.application.service;

import com.fasterxml.jackson.core.type.TypeReference;
import mcp.development_guides.project.domain.model.Variable;
import mcp.development_guides.project.infrastructure.util.JsonReaderUtil;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import java.nio.file.Paths;
import java.util.List;
import java.util.Optional;

/**
 * Service for managing variable configurations from JSON
 */
@Service
public class VariableService {

    private final JsonReaderUtil jsonReaderUtil;

    @Value("{app.data.file.variables}")
    private String variablesFileName;

    @Value("${app.data.directory:data}")
    private String dataDirectory;

    public VariableService(JsonReaderUtil jsonReaderUtil) {
        this.jsonReaderUtil = jsonReaderUtil;
    }

    /**
     * Get all variables from the configuration
     */
    public List<Variable> getAllVariables() {
        String filePath = Paths.get(dataDirectory, variablesFileName).toString();
        TypeReference<List<Variable>> typeReference = new TypeReference<>() {};
        return jsonReaderUtil.readFile(filePath, typeReference);
    }

    /**
     * Find a variable by name
     */
    public Optional<Variable> findVariableByName(String name) {
        return getAllVariables().stream()
                .filter(variable -> variable.name().equals(name))
                .findFirst();
    }

    /**
     * Get variables by type
     */
    public List<Variable> getVariablesByType(String type) {
        return getAllVariables().stream()
                .filter(variable -> variable.type().equals(type))
                .toList();
    }

    /**
     * Validate all variables in the configuration
     */
    public boolean validateAllVariables() {
        return getAllVariables().stream()
                .allMatch(Variable::isValid);
    }

    /**
     * Get invalid variables
     */
    public List<Variable> getInvalidVariables() {
        return getAllVariables().stream()
                .filter(variable -> !variable.isValid())
                .toList();
    }

    /**
     * Count total variables
     */
    public long countVariables() {
        return getAllVariables().size();
    }
}
