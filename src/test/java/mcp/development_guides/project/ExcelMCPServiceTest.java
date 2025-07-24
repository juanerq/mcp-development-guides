package mcp.development_guides.project;

import mcp.development_guides.project.infrastructure.excel.specialized.ExcelMCPService;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.BeforeEach;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.File;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

@SpringBootTest
public class ExcelMCPServiceTest {

    @Autowired
    private ExcelMCPService excelMCPService;

    @Test
    public void testGetFileInfo() {
        // Ruta del archivo Excel de prueba
        String filePath = "/home/juanerq/Downloads/test1.xlsx";

        System.out.println("=== PRUEBA DE getFileInfo ===");
        System.out.println("Archivo: " + filePath);

        try {
            // Verificar si el archivo existe antes de la prueba
            File file = new File(filePath);
            if (!file.exists()) {
                System.out.println("❌ El archivo no existe: " + filePath);
                System.out.println("Creando un archivo de prueba alternativo...");

                // Crear un archivo Excel simple para la prueba
                createTestExcelFile();
                filePath = "/tmp/test_excel.xlsx";
            }

            // Ejecutar la función getFileInfo
            Map<String, Object> fileInfo = excelMCPService.getFileInfo(filePath);

            // Mostrar los resultados
            System.out.println("\n📊 INFORMACIÓN DEL ARCHIVO:");
            System.out.println("═══════════════════════════");

            fileInfo.forEach((key, value) -> {
                if (value instanceof List) {
                    System.out.println(key + ": " + value);
                } else {
                    System.out.println(key + ": " + value);
                }
            });

            // Validaciones
            assertNotNull(fileInfo, "La información del archivo no debe ser null");
            assertTrue(fileInfo.containsKey("fileName"), "Debe contener el nombre del archivo");
            assertTrue(fileInfo.containsKey("filePath"), "Debe contener la ruta del archivo");
            assertTrue(fileInfo.containsKey("fileSize"), "Debe contener el tamaño del archivo");
            assertTrue(fileInfo.containsKey("numberOfSheets"), "Debe contener el número de hojas");
            assertTrue(fileInfo.containsKey("sheetNames"), "Debe contener los nombres de las hojas");

            // Verificar tipos de datos
            assertTrue(fileInfo.get("fileName") instanceof String, "fileName debe ser String");
            assertTrue(fileInfo.get("filePath") instanceof String, "filePath debe ser String");
            assertTrue(fileInfo.get("fileSize") instanceof Long, "fileSize debe ser Long");
            assertTrue(fileInfo.get("numberOfSheets") instanceof Integer, "numberOfSheets debe ser Integer");
            assertTrue(fileInfo.get("sheetNames") instanceof List, "sheetNames debe ser List");

            // Mostrar información específica
            System.out.println("\n✅ VALIDACIONES EXITOSAS:");
            System.out.println("- Nombre del archivo: " + fileInfo.get("fileName"));
            System.out.println("- Número de hojas: " + fileInfo.get("numberOfSheets"));
            System.out.println("- Tamaño del archivo: " + fileInfo.get("fileSize") + " bytes");

            @SuppressWarnings("unchecked")
            List<String> sheetNames = (List<String>) fileInfo.get("sheetNames");
            System.out.println("- Hojas encontradas: " + sheetNames.size());

            for (int i = 0; i < sheetNames.size(); i++) {
                System.out.println("  [" + i + "] " + sheetNames.get(i));
            }

        } catch (Exception e) {
            System.err.println("❌ Error durante la prueba: " + e.getMessage());
            e.printStackTrace();
            fail("La prueba falló con excepción: " + e.getMessage());
        }
    }

    @Test
    public void testGetFileInfoWithInvalidFile() {
        System.out.println("\n=== PRUEBA CON ARCHIVO INVÁLIDO ===");

        String invalidFilePath = "/ruta/inexistente/archivo.xlsx";

        try {
            Map<String, Object> fileInfo = excelMCPService.getFileInfo(invalidFilePath);
            fail("Debería haber lanzado una excepción para archivo inexistente");
        } catch (Exception e) {
            System.out.println("✅ Excepción esperada capturada: " + e.getMessage());
            assertTrue(e.getMessage().contains("Failed to execute Excel operation"),
                       "El mensaje de error debe contener 'Failed to execute Excel operation'");
        }
    }

    @Test
    public void testValidateExcelFile() {
        System.out.println("\n=== PRUEBA DE VALIDACIÓN DE ARCHIVO ===");

        // Prueba con archivo inexistente
        String invalidPath = "/ruta/inexistente/archivo.xlsx";
        boolean isValid = excelMCPService.validateExcelFile(invalidPath);
        assertFalse(isValid, "Un archivo inexistente no debe ser válido");
        System.out.println("✅ Validación de archivo inexistente: " + isValid);

        // Si existe un archivo válido, probarlo
        String validPath = "/home/juanerq/Downloads/test1.xlsx";
        File file = new File(validPath);
        if (file.exists()) {
            boolean isValidFile = excelMCPService.validateExcelFile(validPath);
            System.out.println("✅ Validación de archivo existente: " + isValidFile);
        } else {
            System.out.println("ℹ️  Archivo de prueba no encontrado, creando uno temporal...");
            createTestExcelFile();
            boolean isValidFile = excelMCPService.validateExcelFile("/tmp/test_excel.xlsx");
            System.out.println("✅ Validación de archivo temporal: " + isValidFile);
        }
    }

    /**
     * Crea un archivo Excel simple para pruebas
     */
    private void createTestExcelFile() {
        try {
            String tempFilePath = "/tmp/test_excel.xlsx";
            boolean created = excelMCPService.createNewExcelFile(tempFilePath);
            if (created) {
                // Agregar datos de prueba
                excelMCPService.createSheet(tempFilePath, "Datos");
                excelMCPService.createSheet(tempFilePath, "Resultados");

                // Escribir datos en la hoja principal
                excelMCPService.writeCellValue(tempFilePath, "Sheet", 0, 0, "Nombre");
                excelMCPService.writeCellValue(tempFilePath, "Sheet", 0, 1, "Edad");
                excelMCPService.writeCellValue(tempFilePath, "Sheet", 0, 2, "Ciudad");

                excelMCPService.writeCellValue(tempFilePath, "Sheet", 1, 0, "Juan");
                excelMCPService.writeCellNumber(tempFilePath, "Sheet", 1, 1, 25);
                excelMCPService.writeCellValue(tempFilePath, "Sheet", 1, 2, "Madrid");

                excelMCPService.writeCellValue(tempFilePath, "Sheet", 2, 0, "María");
                excelMCPService.writeCellNumber(tempFilePath, "Sheet", 2, 1, 30);
                excelMCPService.writeCellValue(tempFilePath, "Sheet", 2, 2, "Barcelona");

                System.out.println("✅ Archivo de prueba creado: " + tempFilePath);
            } else {
                System.out.println("❌ No se pudo crear el archivo de prueba");
            }
        } catch (Exception e) {
            System.out.println("❌ Error creando archivo de prueba: " + e.getMessage());
        }
    }

    @Test
    public void testCompleteWorkflow() {
        System.out.println("\n=== PRUEBA DE FLUJO COMPLETO ===");

        String testFile = "/tmp/complete_test.xlsx";

        try {
            // 1. Crear archivo
            System.out.println("1. Creando archivo Excel...");
            boolean created = excelMCPService.createNewExcelFile(testFile);
            assertTrue(created, "El archivo debe crearse exitosamente");

            // 2. Validar archivo
            System.out.println("2. Validando archivo...");
            boolean isValid = excelMCPService.validateExcelFile(testFile);
            assertTrue(isValid, "El archivo debe ser válido");

            // 3. Obtener información del archivo
            System.out.println("3. Obteniendo información del archivo...");
            Map<String, Object> fileInfo = excelMCPService.getFileInfo(testFile);
            assertNotNull(fileInfo, "La información del archivo no debe ser null");

            // 4. Mostrar resultados
            System.out.println("4. Información obtenida:");
            fileInfo.forEach((key, value) ->
                System.out.println("   " + key + ": " + value));

            System.out.println("✅ Flujo completo exitoso");

        } catch (Exception e) {
            System.err.println("❌ Error en flujo completo: " + e.getMessage());
            fail("Flujo completo falló: " + e.getMessage());
        } finally {
            // Limpiar archivo de prueba
            try {
                File file = new File(testFile);
                if (file.exists()) {
                    file.delete();
                    System.out.println("🧹 Archivo de prueba eliminado");
                }
            } catch (Exception e) {
                System.out.println("⚠️  No se pudo eliminar el archivo de prueba: " + e.getMessage());
            }
        }
    }

    @Test
    public void testEditCellFunctions() {
        System.out.println("\n=== PRUEBA DE FUNCIONES DE EDICIÓN ===");

        String testFile = "/tmp/edit_test.xlsx";

        try {
            // 1. Crear archivo con datos iniciales
            System.out.println("1. Creando archivo con datos iniciales...");
            boolean created = excelMCPService.createNewExcelFile(testFile);
            assertTrue(created, "El archivo debe crearse exitosamente");

            // Escribir datos iniciales
            excelMCPService.writeCellValue(testFile, "Sheet", 0, 0, "Producto");
            excelMCPService.writeCellValue(testFile, "Sheet", 0, 1, "Precio");
            excelMCPService.writeCellValue(testFile, "Sheet", 0, 2, "Stock");

            excelMCPService.writeCellValue(testFile, "Sheet", 1, 0, "Laptop");
            excelMCPService.writeCellNumber(testFile, "Sheet", 1, 1, 999.99);
            excelMCPService.writeCellNumber(testFile, "Sheet", 1, 2, 10);

            System.out.println("✅ Datos iniciales escritos");

            // 2. Probar lectura de celdas
            System.out.println("\n2. Probando lectura de celdas...");
            String producto = excelMCPService.readCellValue(testFile, "Sheet", 1, 0);
            String precio = excelMCPService.readCellValue(testFile, "Sheet", 1, 1);
            String stock = excelMCPService.readCellValue(testFile, "Sheet", 1, 2);

            System.out.println("📖 Producto: " + producto);
            System.out.println("📖 Precio: " + precio);
            System.out.println("📖 Stock: " + stock);

            assertEquals("Laptop", producto, "El producto debe ser 'Laptop'");
            assertEquals("999.99", precio, "El precio debe ser '999.99'");
            assertEquals("10.0", stock, "El stock debe ser '10.0'");

            // 3. Probar edición de celdas
            System.out.println("\n3. Probando edición de celdas...");
            boolean updated1 = excelMCPService.writeCellValue(testFile, "Sheet", 1, 0, "Laptop Gaming");
            boolean updated2 = excelMCPService.writeCellNumber(testFile, "Sheet", 1, 1, 1299.99);
            boolean updated3 = excelMCPService.writeCellNumber(testFile, "Sheet", 1, 2, 5);

            assertTrue(updated1, "Debe actualizar el producto exitosamente");
            assertTrue(updated2, "Debe actualizar el precio exitosamente");
            assertTrue(updated3, "Debe actualizar el stock exitosamente");

            // 4. Verificar los cambios
            System.out.println("\n4. Verificando cambios...");
            String nuevoProducto = excelMCPService.readCellValue(testFile, "Sheet", 1, 0);
            String nuevoPrecio = excelMCPService.readCellValue(testFile, "Sheet", 1, 1);
            String nuevoStock = excelMCPService.readCellValue(testFile, "Sheet", 1, 2);

            System.out.println("🔄 Nuevo Producto: " + nuevoProducto);
            System.out.println("🔄 Nuevo Precio: " + nuevoPrecio);
            System.out.println("🔄 Nuevo Stock: " + nuevoStock);

            assertEquals("Laptop Gaming", nuevoProducto, "El producto debe haberse actualizado");
            assertEquals("1299.99", nuevoPrecio, "El precio debe haberse actualizado");
            assertEquals("5.0", nuevoStock, "El stock debe haberse actualizado");

            // 5. Probar escritura de fórmulas
            System.out.println("\n5. Probando escritura de fórmulas...");
            boolean formulaWritten = excelMCPService.writeCellFormula(testFile, "Sheet", 1, 3, "B2*C2");
            assertTrue(formulaWritten, "Debe escribir la fórmula exitosamente");

            String formulaResult = excelMCPService.readCellValue(testFile, "Sheet", 1, 3);
            System.out.println("📊 Resultado de fórmula (Precio * Stock): " + formulaResult);

            // 6. Probar escritura de valores booleanos
            System.out.println("\n6. Probando valores booleanos...");
            boolean boolWritten = excelMCPService.writeCellBoolean(testFile, "Sheet", 1, 4, true);
            assertTrue(boolWritten, "Debe escribir el valor booleano exitosamente");

            String boolValue = excelMCPService.readCellValue(testFile, "Sheet", 1, 4);
            System.out.println("✅ Valor booleano: " + boolValue);
            assertEquals("TRUE", boolValue, "El valor booleano debe ser TRUE");

            // 7. Probar limpieza de celdas
            System.out.println("\n7. Probando limpieza de celdas...");
            boolean cleared = excelMCPService.clearCell(testFile, "Sheet", 1, 4);
            assertTrue(cleared, "Debe limpiar la celda exitosamente");

            String clearedValue = excelMCPService.readCellValue(testFile, "Sheet", 1, 4);
            System.out.println("🧹 Valor después de limpiar: '" + clearedValue + "'");
            assertTrue(clearedValue.isEmpty(), "La celda debe estar vacía después de limpiar");

            System.out.println("\n🎉 ¡TODAS LAS PRUEBAS DE EDICIÓN COMPLETADAS EXITOSAMENTE!");

        } catch (Exception e) {
            System.err.println("❌ Error durante las pruebas de edición: " + e.getMessage());
            e.printStackTrace();
            fail("Las pruebas de edición fallaron: " + e.getMessage());
        } finally {
            // Limpiar archivo de prueba
            try {
                File file = new File(testFile);
                if (file.exists()) {
                    file.delete();
                    System.out.println("🧹 Archivo de prueba eliminado");
                }
            } catch (Exception e) {
                System.out.println("⚠️  No se pudo eliminar el archivo de prueba: " + e.getMessage());
            }
        }
    }

    @Test
    public void testReadRangeAndRowColumn() {
        System.out.println("\n=== PRUEBA DE LECTURA DE RANGOS Y FILAS/COLUMNAS ===");

        String testFile = "/tmp/range_test.xlsx";

        try {
            // 1. Crear archivo con datos en tabla
            System.out.println("1. Creando tabla de datos...");
            boolean created = excelMCPService.createNewExcelFile(testFile);
            assertTrue(created, "El archivo debe crearse exitosamente");

            // Crear una tabla 3x3
            String[][] data = {
                {"A1", "B1", "C1"},
                {"A2", "B2", "C2"},
                {"A3", "B3", "C3"}
            };

            for (int row = 0; row < data.length; row++) {
                for (int col = 0; col < data[row].length; col++) {
                    excelMCPService.writeCellValue(testFile, "Sheet", row, col, data[row][col]);
                }
            }

            System.out.println("✅ Tabla de datos creada");

            // 2. Probar lectura de rango
            System.out.println("\n2. Probando lectura de rango...");
            Object[][] range = excelMCPService.readRange(testFile, "Sheet", 0, 0, 2, 2);

            System.out.println("📊 Rango leído (3x3):");
            for (int i = 0; i < range.length; i++) {
                for (int j = 0; j < range[i].length; j++) {
                    System.out.print(range[i][j] + "\t");
                }
                System.out.println();
            }

            assertEquals("A1", range[0][0].toString(), "Primera celda debe ser A1");
            assertEquals("C3", range[2][2].toString(), "Última celda debe ser C3");

            // 3. Probar lectura de fila
            System.out.println("\n3. Probando lectura de fila...");
            String[] row1 = excelMCPService.readRow(testFile, "Sheet", 1);

            System.out.println("📄 Fila 1 (índice 1): " + String.join(", ", row1));
            assertEquals("A2", row1[0], "Primera celda de la fila debe ser A2");
            assertEquals("B2", row1[1], "Segunda celda de la fila debe ser B2");
            assertEquals("C2", row1[2], "Tercera celda de la fila debe ser C2");

            // 4. Probar lectura de columna
            System.out.println("\n4. Probando lectura de columna...");
            String[] col1 = excelMCPService.readColumn(testFile, "Sheet", 1);

            System.out.println("📄 Columna 1 (índice 1): " + String.join(", ", col1));
            assertEquals("B1", col1[0], "Primera celda de la columna debe ser B1");
            assertEquals("B2", col1[1], "Segunda celda de la columna debe ser B2");
            assertEquals("B3", col1[2], "Tercera celda de la columna debe ser B3");

            System.out.println("\n🎉 ¡PRUEBAS DE LECTURA COMPLETADAS EXITOSAMENTE!");

        } catch (Exception e) {
            System.err.println("❌ Error durante las pruebas de lectura: " + e.getMessage());
            e.printStackTrace();
            fail("Las pruebas de lectura fallaron: " + e.getMessage());
        } finally {
            // Limpiar archivo de prueba
            try {
                File file = new File(testFile);
                if (file.exists()) {
                    file.delete();
                    System.out.println("🧹 Archivo de prueba eliminado");
                }
            } catch (Exception e) {
                System.out.println("⚠️  No se pudo eliminar el archivo de prueba: " + e.getMessage());
            }
        }
    }
}
