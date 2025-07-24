package mcp.development_guides.project;

import mcp.development_guides.project.infrastructure.excel.specialized.ExcelMCPService;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.File;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Prueba específica para editar el archivo Excel real del usuario
 */
@SpringBootTest
public class ExcelRealFileEditTest {

    @Autowired
    private ExcelMCPService excelMCPService;

    @Test
    public void testEditRealExcelFile() {
        System.out.println("🚀 PRUEBA DE EDICIÓN EN ARCHIVO EXCEL REAL");
        System.out.println("==========================================");

        String targetFile = "/home/juanerq/Downloads/test1.xlsx";
        String backupFile = "/home/juanerq/Downloads/test1_backup.xlsx";
        String workingFile = "/home/juanerq/Downloads/test1_working.xlsx";

        try {
            // 1. Verificar si el archivo original existe
            File originalFile = new File(targetFile);

            if (!originalFile.exists() || originalFile.length() == 0) {
                System.out.println("⚠️  Archivo original no encontrado o vacío: " + targetFile);
                System.out.println("📝 Creando archivo de demostración con datos reales...");

                // Crear un archivo de demostración con estructura real
                createRealDemoFile(targetFile);
                System.out.println("✅ Archivo de demostración creado: " + targetFile);
            }

            // 2. Crear una copia de trabajo para no dañar el original
            System.out.println("\n1. Creando copia de seguridad...");
            boolean backupCreated = excelMCPService.copyExcelFile(targetFile, backupFile);

            if (backupCreated) {
                System.out.println("✅ Backup creado: " + backupFile);
            }

            // Trabajar con una copia
            boolean workingCopyCreated = excelMCPService.copyExcelFile(targetFile, workingFile);
            assertTrue(workingCopyCreated, "Debe crear la copia de trabajo");
            System.out.println("✅ Copia de trabajo creada: " + workingFile);

            // 3. Analizar el archivo actual
            System.out.println("\n2. Analizando archivo actual...");
            Map<String, Object> fileInfo = excelMCPService.getFileInfo(workingFile);

            System.out.println("📊 INFORMACIÓN DEL ARCHIVO:");
            System.out.println("   📄 Nombre: " + fileInfo.get("fileName"));
            System.out.println("   💾 Tamaño: " + fileInfo.get("fileSize") + " bytes");
            System.out.println("   📋 Hojas: " + fileInfo.get("numberOfSheets"));

            @SuppressWarnings("unchecked")
            List<String> sheetNames = (List<String>) fileInfo.get("sheetNames");
            System.out.println("   📝 Nombres de hojas: " + sheetNames);

            // 4. Trabajar con la primera hoja disponible
            String sheetToEdit = sheetNames.get(0);
            System.out.println("\n3. Editando hoja: '" + sheetToEdit + "'");

            // 5. Leer algunos datos existentes (si los hay)
            System.out.println("\n4. Leyendo datos existentes...");
            try {
                String cellA1 = excelMCPService.readCellValue(workingFile, sheetToEdit, 0, 0);
                String cellB1 = excelMCPService.readCellValue(workingFile, sheetToEdit, 0, 1);
                String cellA2 = excelMCPService.readCellValue(workingFile, sheetToEdit, 1, 0);

                System.out.println("   📖 A1: '" + cellA1 + "'");
                System.out.println("   📖 B1: '" + cellB1 + "'");
                System.out.println("   📖 A2: '" + cellA2 + "'");
            } catch (Exception e) {
                System.out.println("   ℹ️  No se pudieron leer celdas existentes (archivo puede estar vacío)");
            }

            // 6. Realizar ediciones de demostración
            System.out.println("\n5. Realizando ediciones de demostración...");

            // Agregar encabezados en la fila 0
            boolean edit1 = excelMCPService.writeCellValue(workingFile, sheetToEdit, 0, 0, "ID");
            boolean edit2 = excelMCPService.writeCellValue(workingFile, sheetToEdit, 0, 1, "Nombre");
            boolean edit3 = excelMCPService.writeCellValue(workingFile, sheetToEdit, 0, 2, "Edad");
            boolean edit4 = excelMCPService.writeCellValue(workingFile, sheetToEdit, 0, 3, "Fecha Edición");

            System.out.println("   ✏️  Encabezados agregados: " + (edit1 && edit2 && edit3 && edit4));

            // Agregar datos de ejemplo
            boolean data1 = excelMCPService.writeCellNumber(workingFile, sheetToEdit, 1, 0, 1);
            boolean data2 = excelMCPService.writeCellValue(workingFile, sheetToEdit, 1, 1, "Juan Pérez");
            boolean data3 = excelMCPService.writeCellNumber(workingFile, sheetToEdit, 1, 2, 28);
            boolean data4 = excelMCPService.writeCellValue(workingFile, sheetToEdit, 1, 3, "2025-07-23");

            boolean data5 = excelMCPService.writeCellNumber(workingFile, sheetToEdit, 2, 0, 2);
            boolean data6 = excelMCPService.writeCellValue(workingFile, sheetToEdit, 2, 1, "María García");
            boolean data7 = excelMCPService.writeCellNumber(workingFile, sheetToEdit, 2, 2, 32);
            boolean data8 = excelMCPService.writeCellValue(workingFile, sheetToEdit, 2, 3, "2025-07-23");

            System.out.println("   ✏️  Datos fila 1 agregados: " + (data1 && data2 && data3 && data4));
            System.out.println("   ✏️  Datos fila 2 agregados: " + (data5 && data6 && data7 && data8));

            // 7. Agregar una fórmula
            System.out.println("\n6. Agregando fórmula...");
            boolean formula = excelMCPService.writeCellFormula(workingFile, sheetToEdit, 3, 2, "AVERAGE(C2:C3)");
            System.out.println("   📊 Fórmula de promedio agregada: " + formula);

            // 8. Verificar las ediciones
            System.out.println("\n7. Verificando ediciones realizadas...");

            String headerID = excelMCPService.readCellValue(workingFile, sheetToEdit, 0, 0);
            String headerNombre = excelMCPService.readCellValue(workingFile, sheetToEdit, 0, 1);
            String persona1 = excelMCPService.readCellValue(workingFile, sheetToEdit, 1, 1);
            String edad1 = excelMCPService.readCellValue(workingFile, sheetToEdit, 1, 2);
            String persona2 = excelMCPService.readCellValue(workingFile, sheetToEdit, 2, 1);
            String formulaResult = excelMCPService.readCellValue(workingFile, sheetToEdit, 3, 2);

            System.out.println("   📖 Header ID: '" + headerID + "'");
            System.out.println("   📖 Header Nombre: '" + headerNombre + "'");
            System.out.println("   📖 Persona 1: '" + persona1 + "'");
            System.out.println("   📖 Edad 1: '" + edad1 + "'");
            System.out.println("   📖 Persona 2: '" + persona2 + "'");
            System.out.println("   📖 Resultado fórmula: '" + formulaResult + "'");

            // 9. Validaciones
            assertEquals("ID", headerID, "El encabezado ID debe escribirse correctamente");
            assertEquals("Nombre", headerNombre, "El encabezado Nombre debe escribirse correctamente");
            assertEquals("Juan Pérez", persona1, "El nombre de la persona 1 debe escribirse correctamente");
            assertEquals("28.0", edad1, "La edad debe escribirse correctamente");
            assertEquals("María García", persona2, "El nombre de la persona 2 debe escribirse correctamente");

            // 10. Mostrar información final del archivo editado
            System.out.println("\n8. Información del archivo editado:");
            Map<String, Object> finalInfo = excelMCPService.getFileInfo(workingFile);
            System.out.println("   💾 Tamaño final: " + finalInfo.get("fileSize") + " bytes");

            System.out.println("\n🎉 ¡EDICIÓN DEL ARCHIVO EXCEL COMPLETADA EXITOSAMENTE!");
            System.out.println("📁 Archivos generados:");
            System.out.println("   🔹 Original/Demo: " + targetFile);
            System.out.println("   🔹 Backup: " + backupFile);
            System.out.println("   🔹 Editado: " + workingFile);

        } catch (Exception e) {
            System.err.println("❌ Error durante la edición del archivo: " + e.getMessage());
            e.printStackTrace();
            fail("La edición del archivo falló: " + e.getMessage());
        } finally {
            // Nota: Mantenemos los archivos para que puedas inspeccionarlos
            System.out.println("\nℹ️  Los archivos se mantienen para inspección manual.");
        }
    }

    /**
     * Crea un archivo Excel de demostración con estructura realista
     */
    private void createRealDemoFile(String filePath) {
        try {
            // Crear archivo base
            boolean created = excelMCPService.createNewExcelFile(filePath);
            if (!created) {
                throw new RuntimeException("No se pudo crear el archivo base");
            }

            // Crear hojas adicionales
            excelMCPService.createSheet(filePath, "Empleados");
            excelMCPService.createSheet(filePath, "Ventas");
            excelMCPService.createSheet(filePath, "Resumen");

            // Agregar datos iniciales en la hoja principal
            excelMCPService.writeCellValue(filePath, "Sheet1", 0, 0, "Demo");
            excelMCPService.writeCellValue(filePath, "Sheet1", 0, 1, "Archivo");
            excelMCPService.writeCellValue(filePath, "Sheet1", 1, 0, "Creado");
            excelMCPService.writeCellValue(filePath, "Sheet1", 1, 1, "Para");
            excelMCPService.writeCellValue(filePath, "Sheet1", 2, 0, "Pruebas");
            excelMCPService.writeCellValue(filePath, "Sheet1", 2, 1, "Edición");

            // Agregar datos en hoja Empleados
            excelMCPService.writeCellValue(filePath, "Empleados", 0, 0, "ID");
            excelMCPService.writeCellValue(filePath, "Empleados", 0, 1, "Nombre");
            excelMCPService.writeCellValue(filePath, "Empleados", 0, 2, "Departamento");

            excelMCPService.writeCellNumber(filePath, "Empleados", 1, 0, 101);
            excelMCPService.writeCellValue(filePath, "Empleados", 1, 1, "Ana López");
            excelMCPService.writeCellValue(filePath, "Empleados", 1, 2, "Ventas");

        } catch (Exception e) {
            System.err.println("Error creando archivo de demostración: " + e.getMessage());
            throw new RuntimeException(e);
        }
    }
}
