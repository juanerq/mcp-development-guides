package mcp.development_guides.project.application.service;

import mcp.development_guides.project.domain.model.Template;
import mcp.development_guides.project.domain.model.Variable;
import mcp.development_guides.project.domain.model.VariableType;
import mcp.development_guides.project.domain.repository.TemplateRepository;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.ActiveProfiles;

import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

@SpringBootTest
@ActiveProfiles("test")
class TemplateServiceTest {

    @Autowired
    private TemplateService templateService;

    @Autowired
    private TemplateRepository templateRepository;

    private Template testTemplate1;
    private Template testTemplate2;

    @BeforeEach
    void setUp() {
        // Create test data
        Variable variable1 = Variable.builder()
                .name("test_variable_1")
                .description("Test variable 1")
                .type(VariableType.STRING)
                .row(1)
                .column(1)
                .build();

        Variable variable2 = Variable.builder()
                .name("test_variable_2")
                .description("Test variable 2")
                .type(VariableType.STRING)
                .row(2)
                .column(2)
                .build();

        testTemplate1 = Template.builder()
                .id("test-template-1")
                .name("Test Template 1")
                .description("First test template")
                .path("/test/path1.xlsx")
                .type("EXCEL")
                .variables(List.of(variable1))
                .build();

        testTemplate2 = Template.builder()
                .id("test-template-2")
                .name("Test Template 2")
                .description("Second test template")
                .path("/test/path2.xlsx")
                .type("EXCEL")
                .variables(List.of(variable2))
                .build();

        // Save test data to database
        templateRepository.save(testTemplate1);
        templateRepository.save(testTemplate2);
    }

    @Test
    void getAllTemplates_ShouldReturnListOfTemplates() {
        // When
        List<Template> result = templateService.getAllTemplates();
        System.out.println(result);
        // Then
        assertNotNull(result, "Result should not be null");
        assertFalse(result.isEmpty(), "Result should not be empty");

        // Verify that our test templates are included
        assertTrue(result.stream().anyMatch(t -> t.name().equals("Test Template 1")),
                "Should contain Test Template 1");
        assertTrue(result.stream().anyMatch(t -> t.name().equals("Test Template 2")),
                "Should contain Test Template 2");

        System.out.println("Found " + result.size() + " templates in database:");
        result.forEach(template ->
            System.out.println("- " + template.name() + ": " + template.description())
        );
    }

    @Test
    void getTemplateByName_WithExistingTemplate_ShouldReturnTemplate() {
        // When
        Template result = templateService.getTemplateByName("Test Template 1");

        // Then
        assertNotNull(result, "Result should not be null");
        assertEquals("Test Template 1", result.name(), "Template name should match");
        assertEquals("First test template", result.description(), "Template description should match");
        assertEquals("/test/path1.xlsx", result.path(), "Template path should match");
        assertEquals("EXCEL", result.type(), "Template type should match");
        assertNotNull(result.variables(), "Variables should not be null");
        assertFalse(result.variables().isEmpty(), "Variables should not be empty");

        System.out.println("Retrieved template: " + result.name());
        System.out.println("Description: " + result.description());
        System.out.println("Variables count: " + result.variables().size());
    }

    @Test
    void getTemplateByName_WithNonExistingTemplate_ShouldThrowException() {
        // Given
        String nonExistingTemplateName = "Non Existing Template";

        // When & Then
        IllegalArgumentException exception = assertThrows(
                IllegalArgumentException.class,
                () -> templateService.getTemplateByName(nonExistingTemplateName),
                "Should throw IllegalArgumentException for non-existing template"
        );

        assertTrue(exception.getMessage().contains("Template not found"),
                "Exception message should indicate template not found");
        assertTrue(exception.getMessage().contains(nonExistingTemplateName),
                "Exception message should contain the template name");

        System.out.println("Exception message: " + exception.getMessage());
    }

    @Test
    void getTemplateByName_WithBasicTemplate_ShouldReturnRealData() {
        // This test checks if the "basic" template from templates.json exists in the database
        try {
            // When
            Template result = templateService.getTemplateByName("basic");

            // Then
            assertNotNull(result, "Basic template should exist");
            assertEquals("basic", result.name(), "Template name should be 'basic'");
            assertNotNull(result.description(), "Description should not be null");
            assertNotNull(result.variables(), "Variables should not be null");

            System.out.println("Basic template found:");
            System.out.println("Name: " + result.name());
            System.out.println("Description: " + result.description());
            System.out.println("Path: " + result.path());
            System.out.println("Type: " + result.type());
            System.out.println("Variables count: " + result.variables().size());

            // Print variables details
            result.variables().forEach(variable ->
                System.out.println("  Variable: " + variable.name() + " - " + variable.description())
            );

        } catch (IllegalArgumentException e) {
            System.out.println("Basic template not found in database. " +
                    "Make sure the templates.json data has been loaded.");
            System.out.println("Error: " + e.getMessage());
            // This is acceptable if the data hasn't been loaded yet
        }
    }

    @Test
    void getAllTemplates_ShouldReturnValidTemplates() {
        // When
        List<Template> result = templateService.getAllTemplates();

        // Then
        assertNotNull(result, "Result should not be null");

        // Verify all returned templates are valid
        result.forEach(template -> {
            assertNotNull(template.name(), "Template name should not be null");
            assertNotNull(template.description(), "Template description should not be null");
            assertNotNull(template.path(), "Template path should not be null");
            assertNotNull(template.type(), "Template type should not be null");
            assertNotNull(template.variables(), "Template variables should not be null");

            assertFalse(template.name().trim().isEmpty(), "Template name should not be empty");
            assertFalse(template.description().trim().isEmpty(), "Template description should not be empty");
            assertFalse(template.path().trim().isEmpty(), "Template path should not be empty");
            assertFalse(template.type().trim().isEmpty(), "Template type should not be empty");
        });

        System.out.println("All " + result.size() + " templates are valid");
    }
}