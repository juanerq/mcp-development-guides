package mcp.development_guides.project.infrastructure.excel.writer;

import mcp.development_guides.project.domain.model.ExcelRange;
import mcp.development_guides.project.infrastructure.excel.core.ExcelFileHandler;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import java.awt.Color;
import java.io.FileOutputStream;

/**
 * Editor especializado en formato y estilo de celdas Excel
 */
@Component
public class ExcelFormatWriter {

    @Autowired
    private ExcelFileHandler fileHandler;

    /**
     * Aplica formato de texto (negrita, cursiva, color)
     */
    public boolean formatText(String filePath, String sheetName, int row, int column,
                             boolean bold, boolean italic, Color textColor) {
        return modifyWorkbook(filePath,
            String.format("Formatting text in cell [%d,%d] of sheet '%s'", row, column, sheetName),
            (workbook, path) -> {
                Sheet sheet = fileHandler.getSheetByName(workbook, sheetName);
                Cell cell = getOrCreateCell(sheet, row, column);

                CellStyle style = workbook.createCellStyle();
                Font font = workbook.createFont();

                if (bold) font.setBold(true);
                if (italic) font.setItalic(true);

                // Corregir el manejo de colores para XSSFWorkbook
                if (textColor != null && workbook instanceof XSSFWorkbook) {
                    XSSFFont xssfFont = (XSSFFont) font;
                    xssfFont.setColor(new XSSFColor(textColor, null));
                }

                style.setFont(font);
                cell.setCellStyle(style);
                return true;
            });
    }

    /**
     * Aplica color de fondo a una celda
     */
    public boolean setBackgroundColor(String filePath, String sheetName, int row, int column, Color backgroundColor) {
        return modifyWorkbook(filePath,
            String.format("Setting background color in cell [%d,%d] of sheet '%s'", row, column, sheetName),
            (workbook, path) -> {
                Sheet sheet = fileHandler.getSheetByName(workbook, sheetName);
                Cell cell = getOrCreateCell(sheet, row, column);

                CellStyle style = workbook.createCellStyle();
                if (backgroundColor != null && workbook instanceof XSSFWorkbook) {
                    XSSFCellStyle xssfStyle = (XSSFCellStyle) style;
                    xssfStyle.setFillForegroundColor(new XSSFColor(backgroundColor, null));
                    xssfStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                }

                cell.setCellStyle(style);
                return true;
            });
    }

    /**
     * Aplica bordes a una celda
     */
    public boolean setBorders(String filePath, String sheetName, int row, int column,
                             BorderStyle borderStyle, Color borderColor) {
        return modifyWorkbook(filePath,
            String.format("Setting borders in cell [%d,%d] of sheet '%s'", row, column, sheetName),
            (workbook, path) -> {
                Sheet sheet = fileHandler.getSheetByName(workbook, sheetName);
                Cell cell = getOrCreateCell(sheet, row, column);

                CellStyle style = workbook.createCellStyle();
                style.setBorderTop(borderStyle);
                style.setBorderBottom(borderStyle);
                style.setBorderLeft(borderStyle);
                style.setBorderRight(borderStyle);

                // Corregir el manejo de colores de borde para XSSFWorkbook
                if (borderColor != null && workbook instanceof XSSFWorkbook) {
                    XSSFCellStyle xssfStyle = (XSSFCellStyle) style;
                    XSSFColor color = new XSSFColor(borderColor, null);
                    xssfStyle.setTopBorderColor(color);
                    xssfStyle.setBottomBorderColor(color);
                    xssfStyle.setLeftBorderColor(color);
                    xssfStyle.setRightBorderColor(color);
                }

                cell.setCellStyle(style);
                return true;
            });
    }

    /**
     * Aplica formato numérico (moneda, porcentaje, fecha, etc.)
     */
    public boolean setNumberFormat(String filePath, String sheetName, int row, int column, String formatPattern) {
        return modifyWorkbook(filePath,
            String.format("Setting number format '%s' in cell [%d,%d] of sheet '%s'", formatPattern, row, column, sheetName),
            (workbook, path) -> {
                Sheet sheet = fileHandler.getSheetByName(workbook, sheetName);
                Cell cell = getOrCreateCell(sheet, row, column);

                CellStyle style = workbook.createCellStyle();
                DataFormat format = workbook.createDataFormat();
                style.setDataFormat(format.getFormat(formatPattern));

                cell.setCellStyle(style);
                return true;
            });
    }

    /**
     * Aplica alineación a una celda
     */
    public boolean setAlignment(String filePath, String sheetName, int row, int column,
                               HorizontalAlignment horizontal, VerticalAlignment vertical) {
        return modifyWorkbook(filePath,
            String.format("Setting alignment in cell [%d,%d] of sheet '%s'", row, column, sheetName),
            (workbook, path) -> {
                Sheet sheet = fileHandler.getSheetByName(workbook, sheetName);
                Cell cell = getOrCreateCell(sheet, row, column);

                CellStyle style = workbook.createCellStyle();
                if (horizontal != null) style.setAlignment(horizontal);
                if (vertical != null) style.setVerticalAlignment(vertical);

                cell.setCellStyle(style);
                return true;
            });
    }

    /**
     * Aplica formato a un rango completo de celdas
     */
    public boolean formatRange(String filePath, String sheetName, ExcelRange range, CellStyle templateStyle) {
        return modifyWorkbook(filePath,
            String.format("Formatting range %s in sheet '%s'", range.toExcelNotation(), sheetName),
            (workbook, path) -> {
                Sheet sheet = fileHandler.getSheetByName(workbook, sheetName);

                for (int rowIndex = range.startPosition().row(); rowIndex <= range.endPosition().row(); rowIndex++) {
                    for (int colIndex = range.startPosition().column(); colIndex <= range.endPosition().column(); colIndex++) {
                        Cell cell = getOrCreateCell(sheet, rowIndex, colIndex);
                        cell.setCellStyle(templateStyle);
                    }
                }
                return true;
            });
    }

    // MÉTODOS PRIVADOS DE SOPORTE

    /**
     * Ejecuta una operación de modificación en un workbook y guarda los cambios
     */
    private boolean modifyWorkbook(String filePath, String operation, ExcelFileHandler.WorkbookOperation<Boolean> workbookOperation) {
        try (Workbook workbook = WorkbookFactory.create(new java.io.File(filePath))) {
            System.out.println("🎨 " + operation);

            boolean result = workbookOperation.execute(workbook, filePath);

            // Guardar los cambios
            try (FileOutputStream outputStream = new FileOutputStream(filePath)) {
                workbook.write(outputStream);
                System.out.println("💾 Format changes saved successfully to: " + filePath);
            }

            return result;
        } catch (Exception e) {
            System.err.println("❌ Error during formatting operation: " + operation);
            System.err.println("   File: " + filePath);
            System.err.println("   Error: " + e.getMessage());
            return false;
        }
    }

    /**
     * Obtiene una celda existente o la crea si no existe
     */
    private Cell getOrCreateCell(Sheet sheet, int rowIndex, int columnIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }

        Cell cell = row.getCell(columnIndex);
        if (cell == null) {
            cell = row.createCell(columnIndex);
        }

        return cell;
    }
}
