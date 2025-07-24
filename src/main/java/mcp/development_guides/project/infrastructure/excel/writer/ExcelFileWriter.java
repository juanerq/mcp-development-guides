package mcp.development_guides.project.infrastructure.excel.writer;

import mcp.development_guides.project.infrastructure.excel.core.ExcelFileHandler;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;

/**
 * Editor especializado en operaciones de archivo Excel completo
 */
@Component
public class ExcelFileWriter {

    @Autowired
    private ExcelFileHandler fileHandler;

    /**
     * Crea un nuevo archivo Excel vacío
     */
    public boolean createNewFile(String filePath) {
        try {
            System.out.println("📄 Creating new Excel file: " + filePath);

            try (Workbook workbook = new XSSFWorkbook()) {
                // Crear una hoja por defecto
                workbook.createSheet("Sheet1");

                // Guardar el archivo
                try (FileOutputStream outputStream = new FileOutputStream(filePath)) {
                    workbook.write(outputStream);
                    System.out.println("✅ New Excel file created successfully: " + filePath);
                    return true;
                }
            }
        } catch (Exception e) {
            System.err.println("❌ Error creating new Excel file: " + e.getMessage());
            return false;
        }
    }

    /**
     * Crea un nuevo archivo Excel con hojas específicas
     */
    public boolean createNewFileWithSheets(String filePath, String... sheetNames) {
        try {
            System.out.println("📄 Creating new Excel file with " + sheetNames.length + " sheets: " + filePath);

            try (Workbook workbook = new XSSFWorkbook()) {
                // Crear las hojas especificadas
                for (String sheetName : sheetNames) {
                    workbook.createSheet(sheetName);
                }

                // Si no se especificaron hojas, crear una por defecto
                if (sheetNames.length == 0) {
                    workbook.createSheet("Sheet1");
                }

                // Guardar el archivo
                try (FileOutputStream outputStream = new FileOutputStream(filePath)) {
                    workbook.write(outputStream);
                    System.out.println("✅ New Excel file with sheets created successfully: " + filePath);
                    return true;
                }
            }
        } catch (Exception e) {
            System.err.println("❌ Error creating new Excel file with sheets: " + e.getMessage());
            return false;
        }
    }

    /**
     * Copia un archivo Excel completo
     */
    public boolean copyFile(String sourceFilePath, String targetFilePath) {
        try {
            System.out.println("📋 Copying Excel file from " + sourceFilePath + " to " + targetFilePath);

            Path sourcePath = Path.of(sourceFilePath);
            Path targetPath = Path.of(targetFilePath);

            Files.copy(sourcePath, targetPath, StandardCopyOption.REPLACE_EXISTING);

            System.out.println("✅ Excel file copied successfully");
            return true;
        } catch (Exception e) {
            System.err.println("❌ Error copying Excel file: " + e.getMessage());
            return false;
        }
    }

    /**
     * Combina múltiples archivos Excel en uno solo
     */
    public boolean mergeFiles(String[] sourceFilePaths, String targetFilePath) {
        try {
            System.out.println("🔗 Merging " + sourceFilePaths.length + " Excel files into: " + targetFilePath);

            try (Workbook targetWorkbook = new XSSFWorkbook()) {
                int sheetCounter = 1;

                for (String sourceFilePath : sourceFilePaths) {
                    try (Workbook sourceWorkbook = WorkbookFactory.create(new java.io.File(sourceFilePath))) {

                        for (int i = 0; i < sourceWorkbook.getNumberOfSheets(); i++) {
                            Sheet sourceSheet = sourceWorkbook.getSheetAt(i);
                            String sheetName = "Merged_" + sheetCounter + "_" + sourceSheet.getSheetName();

                            Sheet targetSheet = targetWorkbook.createSheet(sheetName);
                            copySheetData(sourceSheet, targetSheet);
                            sheetCounter++;
                        }
                    }
                }

                // Guardar el archivo combinado
                try (FileOutputStream outputStream = new FileOutputStream(targetFilePath)) {
                    targetWorkbook.write(outputStream);
                    System.out.println("✅ Files merged successfully into: " + targetFilePath);
                    return true;
                }
            }
        } catch (Exception e) {
            System.err.println("❌ Error merging Excel files: " + e.getMessage());
            return false;
        }
    }

    /**
     * Divide un archivo Excel en múltiples archivos (uno por hoja)
     */
    public boolean splitFileBySheets(String sourceFilePath, String outputDirectory) {
        try {
            System.out.println("✂️ Splitting Excel file by sheets: " + sourceFilePath);

            return fileHandler.executeWithWorkbook(sourceFilePath, "Splitting file by sheets", (workbook, path) -> {
                String baseFileName = Path.of(path).getFileName().toString().replaceFirst("[.][^.]+$", "");


                for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                    Sheet sourceSheet = workbook.getSheetAt(i);
                    String sheetFileName = outputDirectory + "/" + baseFileName + "_" + sourceSheet.getSheetName() + ".xlsx";

                    try (Workbook newWorkbook = new XSSFWorkbook()) {
                        Sheet newSheet = newWorkbook.createSheet(sourceSheet.getSheetName());
                        copySheetData(sourceSheet, newSheet);

                        try (FileOutputStream outputStream = new FileOutputStream(sheetFileName)) {
                            newWorkbook.write(outputStream);
                            System.out.println("📄 Created: " + sheetFileName);
                        }
                    }
                }

                System.out.println("✅ File split successfully into " + workbook.getNumberOfSheets() + " files");
                return true;
            });
        } catch (Exception e) {
            System.err.println("❌ Error splitting Excel file: " + e.getMessage());
            return false;
        }
    }

    /**
     * Convierte un archivo Excel a CSV
     */
    public boolean convertToCSV(String excelFilePath, String sheetName, String csvFilePath) {
        try {
            System.out.println("💱 Converting Excel sheet '" + sheetName + "' to CSV: " + csvFilePath);

            return fileHandler.executeWithWorkbook(excelFilePath, "Converting to CSV", (workbook, path) -> {
                Sheet sheet = fileHandler.getSheetByName(workbook, sheetName);

                try (java.io.PrintWriter writer = new java.io.PrintWriter(csvFilePath)) {
                    for (Row row : sheet) {
                        StringBuilder csvRow = new StringBuilder();
                        boolean firstCell = true;

                        for (Cell cell : row) {
                            if (!firstCell) {
                                csvRow.append(",");
                            }

                            String cellValue = getCellValueAsString(cell);
                            // Escapar comillas y agregar comillas si contiene comas
                            if (cellValue.contains(",") || cellValue.contains("\"")) {
                                cellValue = "\"" + cellValue.replace("\"", "\"\"") + "\"";
                            }

                            csvRow.append(cellValue);
                            firstCell = false;
                        }

                        writer.println(csvRow.toString());
                    }

                    System.out.println("✅ Excel converted to CSV successfully");
                    return true;
                }
            });
        } catch (Exception e) {
            System.err.println("❌ Error converting Excel to CSV: " + e.getMessage());
            return false;
        }
    }

    /**
     * Protege un archivo Excel con contraseña
     */
    public boolean protectFile(String filePath, String password) {
        try {
            System.out.println("🔒 Protecting Excel file with password");

            return fileHandler.executeWithWorkbook(filePath, "Protecting file", (workbook, path) -> {
                // Proteger cada hoja
                for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                    Sheet sheet = workbook.getSheetAt(i);
                    sheet.protectSheet(password);
                }

                // Guardar cambios
                try (FileOutputStream outputStream = new FileOutputStream(filePath)) {
                    workbook.write(outputStream);
                    System.out.println("✅ File protected successfully");
                    return true;
                }
            });
        } catch (Exception e) {
            System.err.println("❌ Error protecting Excel file: " + e.getMessage());
            return false;
        }
    }

    // MÉTODOS PRIVADOS DE SOPORTE

    /**
     * Copia datos completos de una hoja a otra
     */
    private void copySheetData(Sheet sourceSheet, Sheet targetSheet) {
        for (Row sourceRow : sourceSheet) {
            Row targetRow = targetSheet.createRow(sourceRow.getRowNum());

            for (Cell sourceCell : sourceRow) {
                Cell targetCell = targetRow.createCell(sourceCell.getColumnIndex());

                // Copiar valor
                switch (sourceCell.getCellType()) {
                    case STRING -> targetCell.setCellValue(sourceCell.getStringCellValue());
                    case NUMERIC -> {
                        if (DateUtil.isCellDateFormatted(sourceCell)) {
                            targetCell.setCellValue(sourceCell.getDateCellValue());
                        } else {
                            targetCell.setCellValue(sourceCell.getNumericCellValue());
                        }
                    }
                    case BOOLEAN -> targetCell.setCellValue(sourceCell.getBooleanCellValue());
                    case FORMULA -> targetCell.setCellFormula(sourceCell.getCellFormula());
                    default -> targetCell.setBlank();
                }

                // Copiar estilo si existe
                if (sourceCell.getCellStyle() != null) {
                    targetCell.setCellStyle(sourceCell.getCellStyle());
                }
            }
        }
    }

    /**
     * Obtiene el valor de una celda como String
     */
    private String getCellValueAsString(Cell cell) {
        switch (cell.getCellType()) {
            case STRING -> { return cell.getStringCellValue(); }
            case NUMERIC -> {
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    double numericValue = cell.getNumericCellValue();
                    if (numericValue == Math.floor(numericValue)) {
                        return String.valueOf((long) numericValue);
                    } else {
                        return String.valueOf(numericValue);
                    }
                }
            }
            case BOOLEAN -> { return String.valueOf(cell.getBooleanCellValue()); }
            case FORMULA -> { return cell.getCellFormula(); }
            default -> { return ""; }
        }
    }
}
