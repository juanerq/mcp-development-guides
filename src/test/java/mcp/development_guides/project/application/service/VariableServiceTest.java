package mcp.development_guides.project.application.service;


import mcp.development_guides.project.domain.model.Variable;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import java.util.List;

import static org.junit.jupiter.api.Assertions.*;
import static org.junit.jupiter.api.Assertions.assertEquals;

@SpringBootTest
public class VariableServiceTest {
    @Autowired
    private VariableService variableService;

    @Test
    void getAllVariables_ShouldReturnAllVariables() {
        List<Variable> variables = variableService.getAllVariables();

        // Then - Validaciones básicas
        assertNotNull(variables);
        assertFalse(variables.isEmpty());
        assertEquals(1, variables.size());

        // Validar el contenido específico de la variable
        Variable variable = variables.get(0);
        assertEquals("mcs-name", variable.name());
        assertEquals("Nombre de microservicio", variable.description());
        assertEquals("TEXT", variable.type());
        assertEquals(1, variable.row());
        assertEquals(1, variable.column());

        System.out.println("All Variables: " + variables);
    }
}
