package com.davivienda.excelpdf.application;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Convertidor de archivos CSV a formato Excel (.xlsx)
 * Diseñado para procesar datos de composición accionaria
 */
public class CsvToExcelConverter {

    private static final Logger logger = LoggerFactory.getLogger(CsvToExcelConverter.class);

    /**
     * Convierte un archivo CSV a formato Excel y lo guarda junto al JAR
     * 
     * @param csvFilePath Ruta del archivo CSV origen
     * @return Ruta del archivo Excel generado
     * @throws IOException Si hay error en la lectura o escritura
     */
    public String convertCsvToExcel(String csvFilePath) throws IOException {
        logger.info("Iniciando conversión de CSV a Excel: {}", csvFilePath);

        // Leer el archivo CSV
        List<String[]> csvData = readCsvFile(csvFilePath);
        
        if (csvData.isEmpty()) {
            throw new IOException("El archivo CSV está vacío");
        }

        // Obtener directorio del JAR para guardar el Excel
        String outputDir = getJarDirectory();
        String excelFileName = "data.xlsx";
        String excelFilePath = Paths.get(outputDir, excelFileName).toString();

        // Crear archivo Excel
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Datos");
            
            // Crear estilos
            CellStyle headerStyle = createHeaderStyle(workbook);
            CellStyle dataStyle = createDataStyle(workbook);
            CellStyle percentageStyle = createPercentageStyle(workbook);

            // Escribir datos
            for (int rowIndex = 0; rowIndex < csvData.size(); rowIndex++) {
                Row row = sheet.createRow(rowIndex);
                String[] rowData = csvData.get(rowIndex);
                
                for (int colIndex = 0; colIndex < rowData.length; colIndex++) {
                    Cell cell = row.createCell(colIndex);
                    String value = rowData[colIndex].trim();
                    
                    // Primera fila es encabezado
                    if (rowIndex == 0) {
                        cell.setCellValue(value);
                        cell.setCellStyle(headerStyle);
                    } else {
                        // Intentar convertir a número si es posible
                        if (colIndex == 2 && !value.isEmpty()) {
                            // Columna C: porcentaje
                            try {
                                double numValue = parsePercentage(value);
                                cell.setCellValue(numValue);
                                cell.setCellStyle(percentageStyle);
                            } catch (NumberFormatException e) {
                                cell.setCellValue(value);
                                cell.setCellStyle(dataStyle);
                            }
                        } else {
                            // Texto
                            cell.setCellValue(value);
                            cell.setCellStyle(dataStyle);
                        }
                    }
                }
            }

            // Ajustar ancho de columnas
            for (int i = 0; i < csvData.get(0).length; i++) {
                sheet.autoSizeColumn(i);
                // Añadir un poco más de espacio
                sheet.setColumnWidth(i, sheet.getColumnWidth(i) + 1000);
            }

            // Guardar archivo
            try (FileOutputStream fileOut = new FileOutputStream(excelFilePath)) {
                workbook.write(fileOut);
            }
        }

        logger.info("Conversión completada. Archivo guardado en: {}", excelFilePath);
        return excelFilePath;
    }

    /**
     * Lee un archivo CSV y retorna sus datos como lista de arrays de strings
     */
    private List<String[]> readCsvFile(String csvFilePath) throws IOException {
        List<String[]> data = new ArrayList<>();
        
        // Detectar el delimitador del CSV
        char delimiter = detectDelimiter(csvFilePath);
        logger.info("Delimitador detectado: '{}'", delimiter);

        try (BufferedReader reader = new BufferedReader(
                new InputStreamReader(new FileInputStream(csvFilePath), StandardCharsets.UTF_8))) {
            
            String line;
            while ((line = reader.readLine()) != null) {
                // Ignorar líneas vacías
                if (line.trim().isEmpty()) {
                    continue;
                }
                
                String[] values = parseCsvLine(line, delimiter);
                data.add(values);
            }
        }

        logger.info("Leídas {} líneas del CSV", data.size());
        return data;
    }

    /**
     * Detecta el delimitador usado en el CSV (coma o punto y coma)
     */
    private char detectDelimiter(String csvFilePath) throws IOException {
        try (BufferedReader reader = new BufferedReader(
                new InputStreamReader(new FileInputStream(csvFilePath), StandardCharsets.UTF_8))) {
            
            String firstLine = reader.readLine();
            if (firstLine == null) {
                return ',';
            }

            // Contar comas y punto y comas
            int commas = firstLine.length() - firstLine.replace(",", "").length();
            int semicolons = firstLine.length() - firstLine.replace(";", "").length();

            return semicolons > commas ? ';' : ',';
        }
    }

    /**
     * Parsea una línea CSV considerando valores entre comillas
     */
    private String[] parseCsvLine(String line, char delimiter) {
        List<String> values = new ArrayList<>();
        StringBuilder currentValue = new StringBuilder();
        boolean inQuotes = false;

        for (int i = 0; i < line.length(); i++) {
            char c = line.charAt(i);

            if (c == '"') {
                inQuotes = !inQuotes;
            } else if (c == delimiter && !inQuotes) {
                values.add(currentValue.toString().trim());
                currentValue.setLength(0);
            } else {
                currentValue.append(c);
            }
        }

        // Agregar el último valor
        values.add(currentValue.toString().trim());

        return values.toArray(new String[0]);
    }

    /**
     * Parsea un valor de porcentaje (ej: "50%", "0.5", "50")
     */
    private double parsePercentage(String value) {
        value = value.replace("%", "").replace(",", ".").trim();
        double numValue = Double.parseDouble(value);
        
        // Si el valor es mayor a 1, asumir que está en formato porcentaje (50 = 50%)
        if (numValue > 1) {
            numValue = numValue / 100.0;
        }
        
        return numValue;
    }

    /**
     * Obtiene el directorio donde está ubicado el JAR
     */
    private String getJarDirectory() {
        try {
            String jarPath = CsvToExcelConverter.class
                    .getProtectionDomain()
                    .getCodeSource()
                    .getLocation()
                    .toURI()
                    .getPath();
            
            File jarFile = new File(jarPath);
            String directory = jarFile.isDirectory() 
                    ? jarFile.getAbsolutePath() 
                    : jarFile.getParent();
            
            logger.info("Directorio del JAR: {}", directory);
            return directory;
        } catch (Exception e) {
            // Fallback al directorio actual
            logger.warn("No se pudo determinar directorio del JAR, usando directorio actual");
            return System.getProperty("user.dir");
        }
    }

    /**
     * Crea estilo para encabezados
     */
    private CellStyle createHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 11);
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        return style;
    }

    /**
     * Crea estilo para datos de texto
     */
    private CellStyle createDataStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        return style;
    }

    /**
     * Crea estilo para porcentajes
     */
    private CellStyle createPercentageStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        
        // Formato de porcentaje con 4 decimales
        DataFormat format = workbook.createDataFormat();
        style.setDataFormat(format.getFormat("0.0000%"));
        
        return style;
    }
}
