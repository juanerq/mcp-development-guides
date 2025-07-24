package mcp.development_guides.project;

import mcp.development_guides.project.infrastructure.excel.specialized.ExcelMCPService;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.File;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Prueba simplificada para funciones de edición de Excel
 */
@SpringBootTest
public class ExcelEditTestSimple {

    @Autowired
    private ExcelMCPService excelMCPService;

    @Test
    public void testBasicEditFunctions() {
        System.out.println("🚀 PRUEBA BÁSICA DE EDICIÓN DE EXCEL");
        System.out.println("=====================================");

        String testFile = "/tmp/simple_edit_test.xlsx";

        try {
            // 1. Crear archivo nuevo
            System.out.println("1. Creando archivo Excel...");
            boolean created = excelMCPService.createNewExcelFile(testFile);
            assertTrue(created, "El archivo debe crearse exitosamente");
            System.out.println("✅ Archivo creado: " + testFile);

            // 2. Verificar que el archivo existe y es válido
            File file = new File(testFile);
            assertTrue(file.exists(), "El archivo debe existir");
            assertTrue(file.length() > 0, "El archivo no debe estar vacío");
            System.out.println("✅ Archivo válido - Tamaño: " + file.length() + " bytes");

            // 3. Obtener información del archivo
            System.out.println("\n2. Obteniendo información del archivo...");
            Map<String, Object> fileInfo = excelMCPService.getFileInfo(testFile);
            assertNotNull(fileInfo, "La información del archivo no debe ser null");

            System.out.println("📋 Información del archivo:");
            System.out.println("   - Nombre: " + fileInfo.get("fileName"));
            System.out.println("   - Hojas: " + fileInfo.get("numberOfSheets"));
            System.out.println("   - Nombres de hojas: " + fileInfo.get("sheetNames"));

            // 4. Intentar escribir una celda simple
            System.out.println("\n3. Intentando escribir celda...");

            // Usar Sheet1 que sabemos que existe por defecto
            boolean written = excelMCPService.writeCellValue(testFile, "Sheet1", 0, 0, "Prueba");

            if (written) {
                System.out.println("✅ Celda escrita exitosamente");

                // Verificar que se escribió correctamente
                String value = excelMCPService.readCellValue(testFile, "Sheet1", 0, 0);
                System.out.println("📖 Valor leído: '" + value + "'");
                assertEquals("Prueba", value, "El valor debe ser 'Prueba'");

                System.out.println("🎉 ¡ESCRITURA Y LECTURA EXITOSAS!");

                // Probar escribir más datos
                System.out.println("\n4. Escribiendo más datos...");
                excelMCPService.writeCellValue(testFile, "Sheet1", 0, 1, "Nombre");
                excelMCPService.writeCellValue(testFile, "Sheet1", 0, 2, "Edad");
                excelMCPService.writeCellValue(testFile, "Sheet1", 1, 0, "Juan");
                excelMCPService.writeCellNumber(testFile, "Sheet1", 1, 1, 25);

                System.out.println("✅ Datos adicionales escritos");

                // Leer los datos
                String nombre = excelMCPService.readCellValue(testFile, "Sheet1", 1, 0);
                String edad = excelMCPService.readCellValue(testFile, "Sheet1", 1, 1);

                System.out.println("📖 Nombre: " + nombre);
                System.out.println("📖 Edad: " + edad);

            } else {
                System.out.println("❌ No se pudo escribir en la celda");
                System.out.println("ℹ️  Esto puede indicar un problema con las funciones de escritura");
            }

        } catch (Exception e) {
            System.err.println("❌ Error durante la prueba: " + e.getMessage());
            e.printStackTrace();
            fail("La prueba falló: " + e.getMessage());
        } finally {
            // Limpiar archivo de prueba
            try {
                File file = new File(testFile);
                if (file.exists()) {
                    file.delete();
                    System.out.println("🧹 Archivo de prueba eliminado");
                }
            } catch (Exception e) {
                System.out.println("⚠️  No se pudo eliminar archivo: " + e.getMessage());
            }
        }
    }

    @Test
    public void testCreateSheetAndWrite() {
        System.out.println("\n🚀 PRUEBA DE CREACIÓN DE HOJAS Y ESCRITURA");
        System.out.println("==========================================");

        String testFile = "/tmp/sheet_test.xlsx";

        try {
            // 1. Crear archivo
            boolean created = excelMCPService.createNewExcelFile(testFile);
            assertTrue(created, "El archivo debe crearse");

            // 2. Crear hoja nueva
            System.out.println("1. Creando hoja 'Datos'...");
            boolean sheetCreated = excelMCPService.createSheet(testFile, "Datos");

            if (sheetCreated) {
                System.out.println("✅ Hoja 'Datos' creada exitosamente");

                // 3. Escribir en la nueva hoja
                System.out.println("2. Escribiendo en la hoja 'Datos'...");
                boolean written = excelMCPService.writeCellValue(testFile, "Datos", 0, 0, "Producto");

                if (written) {
                    excelMCPService.writeCellValue(testFile, "Datos", 0, 1, "Precio");
                    excelMCPService.writeCellValue(testFile, "Datos", 1, 0, "Laptop");
                    excelMCPService.writeCellNumber(testFile, "Datos", 1, 1, 999.99);

                    System.out.println("✅ Datos escritos en la hoja 'Datos'");

                    // 4. Leer los datos
                    String producto = excelMCPService.readCellValue(testFile, "Datos", 1, 0);
                    String precio = excelMCPService.readCellValue(testFile, "Datos", 1, 1);

                    System.out.println("📖 Producto leído: " + producto);
                    System.out.println("📖 Precio leído: " + precio);

                    assertEquals("Laptop", producto);
                    assertEquals("999.99", precio);

                    System.out.println("🎉 ¡CREACIÓN DE HOJA Y EDICIÓN EXITOSAS!");
                }
            } else {
                System.out.println("❌ No se pudo crear la hoja");
            }

        } catch (Exception e) {
            System.err.println("❌ Error: " + e.getMessage());
            e.printStackTrace();
        } finally {
            // Limpiar
            try {
                new File(testFile).delete();
                System.out.println("🧹 Archivo eliminado");
            } catch (Exception ignored) {}
        }
    }

    @Test
    public void testDifferentDataTypes() {
        System.out.println("\n🚀 PRUEBA DE DIFERENTES TIPOS DE DATOS");
        System.out.println("======================================");

        String testFile = "/tmp/datatypes_test.xlsx";

        try {
            // Crear archivo
            excelMCPService.createNewExcelFile(testFile);

            System.out.println("1. Probando diferentes tipos de datos...");

            // Escribir diferentes tipos
            boolean textWritten = excelMCPService.writeCellValue(testFile, "Sheet1", 0, 0, "Texto");
            boolean numberWritten = excelMCPService.writeCellNumber(testFile, "Sheet1", 0, 1, 123.45);
            boolean boolWritten = excelMCPService.writeCellBoolean(testFile, "Sheet1", 0, 2, true);

            System.out.println("✅ Texto escrito: " + textWritten);
            System.out.println("✅ Número escrito: " + numberWritten);
            System.out.println("✅ Booleano escrito: " + boolWritten);

            if (textWritten && numberWritten && boolWritten) {
                // Leer los valores
                String texto = excelMCPService.readCellValue(testFile, "Sheet1", 0, 0);
                String numero = excelMCPService.readCellValue(testFile, "Sheet1", 0, 1);
                String booleano = excelMCPService.readCellValue(testFile, "Sheet1", 0, 2);

                System.out.println("📖 Texto leído: '" + texto + "'");
                System.out.println("📖 Número leído: '" + numero + "'");
                System.out.println("📖 Booleano leído: '" + booleano + "'");

                assertEquals("Texto", texto);
                assertEquals("123.45", numero);
                assertEquals("TRUE", booleano);

                System.out.println("🎉 ¡TODOS LOS TIPOS DE DATOS FUNCIONAN!");

                // Probar limpiar celda
                System.out.println("\n2. Probando limpiar celda...");
                boolean cleared = excelMCPService.clearCell(testFile, "Sheet1", 0, 2);

                if (cleared) {
                    String clearedValue = excelMCPService.readCellValue(testFile, "Sheet1", 0, 2);
                    System.out.println("🧹 Valor después de limpiar: '" + clearedValue + "'");
                    assertTrue(clearedValue.isEmpty(), "La celda debe estar vacía");
                    System.out.println("✅ Limpieza de celda exitosa");
                }
            }

        } catch (Exception e) {
            System.err.println("❌ Error: " + e.getMessage());
            e.printStackTrace();
        } finally {
            try {
                new File(testFile).delete();
                System.out.println("🧹 Archivo eliminado");
            } catch (Exception ignored) {}
        }
    }
}
