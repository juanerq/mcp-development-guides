package mcp.development_guides.project;

import mcp.development_guides.project.infrastructure.excel.specialized.ExcelMCPService;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.File;
import java.util.List;
import java.util.Map;

/**
 * Prueba interactiva para demostrar la función getFileInfo del servicio Excel
 */
@SpringBootTest
public class ExcelTestRunner {

    @Autowired
    private ExcelMCPService excelMCPService;

    @Test
    public void runExcelDemo() {
        System.out.println("🚀 INICIANDO PRUEBA DE EXCEL getFileInfo");
        System.out.println("==========================================");

        // Archivo de prueba
        String testFilePath = "/home/juanerq/Downloads/test1.xlsx";

        // Verificar si el archivo existe
        File testFile = new File(testFilePath);
        if (!testFile.exists()) {
            System.out.println("⚠️  Archivo original no encontrado: " + testFilePath);
            System.out.println("📝 Creando archivo de prueba temporal...");

            // Crear archivo de prueba
            testFilePath = createTestFile();
            if (testFilePath == null) {
                System.out.println("❌ No se pudo crear archivo de prueba");
                return;
            }
        }

        System.out.println("📂 Usando archivo: " + testFilePath);
        System.out.println();

        try {
            // PRUEBA 1: Validar archivo
            System.out.println("🔍 PRUEBA 1: Validando archivo Excel");
            System.out.println("-----------------------------------");
            boolean isValid = excelMCPService.validateExcelFile(testFilePath);
            System.out.println("✅ Archivo válido: " + isValid);
            System.out.println();

            if (!isValid) {
                System.out.println("❌ El archivo no es válido, terminando pruebas");
                return;
            }

            // PRUEBA 2: Obtener información del archivo
            System.out.println("📊 PRUEBA 2: Obteniendo información del archivo");
            System.out.println("----------------------------------------------");

            Map<String, Object> fileInfo = excelMCPService.getFileInfo(testFilePath);

            System.out.println("📋 INFORMACIÓN OBTENIDA:");
            System.out.println("══════════════════════");

            // Mostrar cada campo de información
            System.out.println("📄 Nombre del archivo: " + fileInfo.get("fileName"));
            System.out.println("📁 Ruta completa: " + fileInfo.get("filePath"));
            System.out.println("💾 Tamaño del archivo: " + fileInfo.get("fileSize") + " bytes");
            System.out.println("📊 Número de hojas: " + fileInfo.get("numberOfSheets"));

            @SuppressWarnings("unchecked")
            List<String> sheetNames = (List<String>) fileInfo.get("sheetNames");
            System.out.println("📋 Nombres de hojas (" + sheetNames.size() + "):");

            for (int i = 0; i < sheetNames.size(); i++) {
                System.out.println("   [" + (i + 1) + "] " + sheetNames.get(i));
            }

            System.out.println();

            // PRUEBA 3: Obtener nombres de hojas usando método específico
            System.out.println("📝 PRUEBA 3: Obteniendo nombres de hojas (método alternativo)");
            System.out.println("----------------------------------------------------------");

            List<String> directSheetNames = excelMCPService.getSheetNames(testFilePath);
            System.out.println("✅ Hojas encontradas directamente: " + directSheetNames.size());
            directSheetNames.forEach(name -> System.out.println("   • " + name));

            System.out.println();
            System.out.println("🎉 ¡TODAS LAS PRUEBAS COMPLETADAS EXITOSAMENTE!");

        } catch (Exception e) {
            System.err.println("❌ ERROR DURANTE LAS PRUEBAS:");
            System.err.println("   Mensaje: " + e.getMessage());
            System.err.println("   Causa: " + (e.getCause() != null ? e.getCause().getMessage() : "No especificada"));
            e.printStackTrace();
        }

        System.out.println();
        System.out.println("🏁 Fin de las pruebas");
    }

    /**
     * Crea un archivo Excel de prueba con algunas hojas
     */
    private String createTestFile() {
        try {
            String tempPath = "/tmp/excel_test_demo.xlsx";

            System.out.println("📝 Creando archivo Excel de prueba...");

            // Crear archivo base
            boolean created = excelMCPService.createNewExcelFile(tempPath);
            if (!created) {
                System.out.println("❌ No se pudo crear el archivo base");
                return null;
            }

            // Crear hojas adicionales
            System.out.println("📋 Agregando hojas adicionales...");
            excelMCPService.createSheet(tempPath, "Datos");
            excelMCPService.createSheet(tempPath, "Resultados");
            excelMCPService.createSheet(tempPath, "Configuración");

            // Escribir algunos datos de ejemplo
            System.out.println("✏️  Escribiendo datos de ejemplo...");
            excelMCPService.writeCellValue(tempPath, "Sheet", 0, 0, "Ejemplo de datos");
            excelMCPService.writeCellValue(tempPath, "Datos", 0, 0, "Nombre");
            excelMCPService.writeCellValue(tempPath, "Datos", 0, 1, "Edad");
            excelMCPService.writeCellValue(tempPath, "Datos", 1, 0, "Juan");
            excelMCPService.writeCellNumber(tempPath, "Datos", 1, 1, 25);

            System.out.println("✅ Archivo de prueba creado exitosamente: " + tempPath);
            return tempPath;

        } catch (Exception e) {
            System.err.println("❌ Error creando archivo de prueba: " + e.getMessage());
            return null;
        }
    }
}
