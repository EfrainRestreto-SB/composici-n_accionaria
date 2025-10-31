package com.davivienda.excelpdf.application;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.davivienda.excelpdf.infrastructure.PdfOwnershipReportGenerator;

/**
 * Procesador principal que coordina la lectura de Excel y generación de PDF
 * para el análisis de composición accionaria.
 * 
 * @author Davivienda
 * @version 1.0
 */
public class ExcelOwnershipProcessor {
    
    private static final Logger logger = LoggerFactory.getLogger(ExcelOwnershipProcessor.class);
    
    private final OwnershipCalculator calculator;
    private final PdfOwnershipReportGenerator pdfGenerator;
    
    /**
     * Constructor del procesador.
     */
    public ExcelOwnershipProcessor() {
        this.calculator = new OwnershipCalculator();
        this.pdfGenerator = new PdfOwnershipReportGenerator();
    }
    
    /**
     * Procesa un archivo Excel y genera un reporte PDF con la composición accionaria.
     * 
     * @param excelPath ruta del archivo Excel de entrada
     * @param rootEntity nombre de la entidad raíz para el cálculo
     * @param outputPdfPath ruta del archivo PDF de salida
     * @return resultado del procesamiento con estadísticas
     * @throws IOException si hay problemas de acceso a archivos
     * @throws IllegalArgumentException si los parámetros son inválidos
     */
    public ProcessingResult processOwnershipAnalysis(String excelPath, String rootEntity, String outputPdfPath) 
            throws IOException {
        
        logger.info("=== Iniciando análisis de composición accionaria ===");
        logger.info("Archivo Excel: {}", excelPath);
        logger.info("Entidad raíz: {}", rootEntity);
        logger.info("PDF salida: {}", outputPdfPath);
        
        // Validar parámetros de entrada
        validateInputParameters(excelPath, rootEntity, outputPdfPath);
        
        long startTime = System.currentTimeMillis();
        ProcessingResult.Builder resultBuilder = new ProcessingResult.Builder();
        
        try {
            // Paso 1: Cargar datos desde Excel
            logger.info(" Paso 1: Cargando datos desde Excel...");
            calculator.loadFromExcel(excelPath);
            resultBuilder.withGraphStatistics(calculator.getGraphStatistics());
            logger.info(" Datos cargados exitosamente");
            
            // Paso 2: Calcular participaciones finales
            logger.info(" Paso 2: Calculando participaciones finales...");
            calculator.calculateFinalOwnership(rootEntity);
            
            Map<String, Double> finalResults = calculator.getFinalResults();
            Map<String, String> beneficiaryPaths = calculator.getBeneficiaryPaths();
            Map<String, Map<String, Double>> originalData = calculator.getOriginalData();
            
            resultBuilder
                .withFinalResults(finalResults)
                .withBeneficiaryPaths(beneficiaryPaths)
                .withRootEntity(rootEntity);
                
            logger.info(" Cálculos completados. Beneficiarios encontrados: {}", finalResults.size());
            
            // Paso 3: Cargar datos originales para la tabla de desglose
            logger.info(" Paso 3: Cargando datos originales de data.xlsx...");
            Map<String, Map<String, Double>> originalDataFromFile = loadOriginalData("data.xlsx");
            java.util.List<String[]> dataXlsxRows = loadDataXlsxRows("data.xlsx");
            
            // Paso 4: Generar reporte PDF
            logger.info(" Paso 4: Generando reporte PDF...");
            pdfGenerator.generateOwnershipReport(
                finalResults, 
                beneficiaryPaths, 
                rootEntity, 
                outputPdfPath,
                originalDataFromFile,
                dataXlsxRows
            );
            
            // Verificar que el PDF se generó correctamente
            Path pdfPath = Paths.get(outputPdfPath);
            if (!Files.exists(pdfPath)) {
                throw new IOException("El archivo PDF no se generó correctamente: " + outputPdfPath);
            }
            
            long pdfSize = Files.size(pdfPath);
            if (pdfSize == 0) {
                throw new IOException("El archivo PDF generado está vacío");
            }
            
            resultBuilder
                .withOutputPdfPath(outputPdfPath)
                .withPdfSize(pdfSize);
                
            logger.info(" PDF generado exitosamente. Tamaño: {} bytes", pdfSize);
            
            // Calcular tiempo total
            long processingTime = System.currentTimeMillis() - startTime;
            resultBuilder.withProcessingTime(processingTime);
            
            ProcessingResult result = resultBuilder.build();
            
            logger.info("=== Procesamiento completado exitosamente ===");
            logger.info("Tiempo total: {} ms", processingTime);
            logger.info("Beneficiarios finales: {}", finalResults.size());
            logger.info("Archivo PDF: {} ({} bytes)", outputPdfPath, pdfSize);
            
            return result;
            
        } catch (Exception e) {
            logger.error(" Error durante el procesamiento: {}", e.getMessage(), e);
            throw e;
        }
    }
    
    /**
     * Valida los parámetros de entrada.
     */
    private void validateInputParameters(String excelPath, String rootEntity, String outputPdfPath) {
        if (excelPath == null || excelPath.trim().isEmpty()) {
            throw new IllegalArgumentException("La ruta del archivo Excel no puede estar vacía");
        }
        
        if (rootEntity == null || rootEntity.trim().isEmpty()) {
            throw new IllegalArgumentException("El nombre de la entidad raíz no puede estar vacío");
        }
        
        if (outputPdfPath == null || outputPdfPath.trim().isEmpty()) {
            throw new IllegalArgumentException("La ruta del archivo PDF de salida no puede estar vacía");
        }
        
        // Verificar que el archivo Excel existe
        Path excelFile = Paths.get(excelPath);
        if (!Files.exists(excelFile)) {
            throw new IllegalArgumentException("El archivo Excel no existe: " + excelPath);
        }
        
        if (!Files.isReadable(excelFile)) {
            throw new IllegalArgumentException("El archivo Excel no es legible: " + excelPath);
        }
        
        // Verificar extensión del archivo Excel
        String fileName = excelFile.getFileName().toString().toLowerCase();
        if (!fileName.endsWith(".xlsx")) {
            throw new IllegalArgumentException("El archivo debe tener extensión .xlsx: " + excelPath);
        }
        
        // Verificar que el directorio de salida existe o se puede crear
        Path outputDir = Paths.get(outputPdfPath).getParent();
        if (outputDir != null && !Files.exists(outputDir)) {
            try {
                Files.createDirectories(outputDir);
            } catch (IOException e) {
                throw new IllegalArgumentException("No se puede crear el directorio de salida: " + outputDir);
            }
        }
    }
    
    /**
     * Obtiene la calculadora de participaciones para acceso directo.
     * 
     * @return calculadora de participaciones
     */
    public OwnershipCalculator getCalculator() {
        return calculator;
    }
    
    /**
     * Clase que encapsula el resultado del procesamiento.
     */
    public static class ProcessingResult {
        private final Map<String, Double> finalResults;
        private final Map<String, String> beneficiaryPaths;
        private final String rootEntity;
        private final String outputPdfPath;
        private final long pdfSize;
        private final long processingTime;
        private final String graphStatistics;
        
        private ProcessingResult(Builder builder) {
            this.finalResults = builder.finalResults;
            this.beneficiaryPaths = builder.beneficiaryPaths;
            this.rootEntity = builder.rootEntity;
            this.outputPdfPath = builder.outputPdfPath;
            this.pdfSize = builder.pdfSize;
            this.processingTime = builder.processingTime;
            this.graphStatistics = builder.graphStatistics;
        }
        
        // Getters
        public Map<String, Double> getFinalResults() { return finalResults; }
        public Map<String, String> getBeneficiaryPaths() { return beneficiaryPaths; }
        public String getRootEntity() { return rootEntity; }
        public String getOutputPdfPath() { return outputPdfPath; }
        public long getPdfSize() { return pdfSize; }
        public long getProcessingTime() { return processingTime; }
        public String getGraphStatistics() { return graphStatistics; }
        
        /**
         * Builder para crear resultados de procesamiento.
         */
        public static class Builder {
            private Map<String, Double> finalResults;
            private Map<String, String> beneficiaryPaths;
            private String rootEntity;
            private String outputPdfPath;
            private long pdfSize;
            private long processingTime;
            private String graphStatistics;
            
            public Builder withFinalResults(Map<String, Double> finalResults) {
                this.finalResults = finalResults;
                return this;
            }
            
            public Builder withBeneficiaryPaths(Map<String, String> beneficiaryPaths) {
                this.beneficiaryPaths = beneficiaryPaths;
                return this;
            }
            
            public Builder withRootEntity(String rootEntity) {
                this.rootEntity = rootEntity;
                return this;
            }
            
            public Builder withOutputPdfPath(String outputPdfPath) {
                this.outputPdfPath = outputPdfPath;
                return this;
            }
            
            public Builder withPdfSize(long pdfSize) {
                this.pdfSize = pdfSize;
                return this;
            }
            
            public Builder withProcessingTime(long processingTime) {
                this.processingTime = processingTime;
                return this;
            }
            
            public Builder withGraphStatistics(String graphStatistics) {
                this.graphStatistics = graphStatistics;
                return this;
            }
            
            public ProcessingResult build() {
                return new ProcessingResult(this);
            }
        }
    }
    
    /**
     * Carga los datos originales del archivo Excel para generar la tabla de desglose.
     * Lee el archivo data.xlsx y construye la estructura jerárquica original.
     * 
     * @param excelPath ruta al archivo Excel original (data.xlsx)
     * @return Map con la estructura jerárquica del Excel original
     */
    private Map<String, Map<String, Double>> loadOriginalData(String excelPath) {
        Map<String, Map<String, Double>> originalDataMap = new HashMap<>();
        
        try {
            logger.info("Cargando datos originales desde: {}", excelPath);
            
            File excelFile = new File(excelPath);
            if (!excelFile.exists()) {
                logger.warn("Archivo original no encontrado: {}", excelPath);
                return originalDataMap;
            }
            
            try (FileInputStream fis = new FileInputStream(excelFile);
                 Workbook workbook = WorkbookFactory.create(fis)) {
                
                Sheet sheet = workbook.getSheetAt(0);
                
                // Leer todas las filas del Excel
                for (Row row : sheet) {
                    if (row.getRowNum() == 0) continue; // Saltar encabezado
                    
                    Cell entityCell = row.getCell(0);
                    Cell percentageCell = row.getCell(2);
                    
                    if (entityCell != null && percentageCell != null) {
                        String entityName = getCellValueAsString(entityCell);
                        
                        if (entityName != null && !entityName.trim().isEmpty()) {
                            try {
                                double percentage = getCellValueAsDouble(percentageCell);
                                
                                // Crear mapa interno para la entidad
                                Map<String, Double> entityData = new HashMap<>();
                                entityData.put("PERCENTAGE_TOTAL", percentage);
                                
                                originalDataMap.put(entityName.trim(), entityData);
                                
                            } catch (Exception e) {
                                logger.warn("Error procesando porcentaje para entidad {}: {}", 
                                           entityName, e.getMessage());
                            }
                        }
                    }
                }
                
                logger.info("Datos originales cargados exitosamente. Entidades procesadas: {}", 
                           originalDataMap.size());
                
            }
            
        } catch (Exception e) {
            logger.error("Error cargando datos originales desde {}: {}", excelPath, e.getMessage(), e);
        }
        
        return originalDataMap;
    }
    
    /**
     * Carga las filas 4-45 del archivo data.xlsx con las tres columnas exactas para el PDF.
     * 
     * @param excelPath ruta al archivo Excel original (data.xlsx)
     * @return Lista de datos estructurados para la tabla del PDF
     */
    private java.util.List<String[]> loadDataXlsxRows(String excelPath) {
        java.util.List<String[]> tableData = new java.util.ArrayList<>();
        
        try {
            logger.info("Cargando filas 4-45 de data.xlsx para tabla PDF");
            
            File excelFile = new File(excelPath);
            if (!excelFile.exists()) {
                logger.warn("Archivo data.xlsx no encontrado: {}", excelPath);
                return tableData;
            }
            
            try (FileInputStream fis = new FileInputStream(excelFile);
                 Workbook workbook = WorkbookFactory.create(fis)) {
                
                Sheet sheet = workbook.getSheetAt(0);
                
                // Leer específicamente las filas 4-45 (índices 3-44)
                for (int rowIndex = 3; rowIndex < 45 && rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                    Row row = sheet.getRow(rowIndex);
                    if (row == null) continue;
                    
                    Cell columnA = row.getCell(0); // Entidad/Nombre
                    Cell columnB = row.getCell(1); // Porcentaje directo
                    Cell columnC = row.getCell(2); // Porcentaje final
                    
                    String entityName = getCellValueAsString(columnA);
                    String directPercentage = getCellValueAsString(columnB);
                    String finalPercentage = getCellValueAsString(columnC);
                    
                    // Solo incluir filas que tienen al menos algo en una de las tres columnas
                    if ((entityName != null && !entityName.trim().isEmpty()) || 
                        (directPercentage != null && !directPercentage.trim().isEmpty()) ||
                        (finalPercentage != null && !finalPercentage.trim().isEmpty())) {
                        
                        // Limpiar valores nulos o vacíos
                        entityName = (entityName != null) ? entityName.trim() : "";
                        
                        // Convertir valores decimales a porcentajes (multiplicar por 100)
                        directPercentage = convertToPercentage(directPercentage);
                        finalPercentage = convertToPercentage(finalPercentage);
                        
                        // Agregar la fila como array de strings
                        tableData.add(new String[]{entityName, directPercentage, finalPercentage});
                    }
                }
                
                logger.info("Filas de data.xlsx cargadas para tabla PDF: {}", tableData.size());
                
            }
            
        } catch (Exception e) {
            logger.error("Error cargando filas de data.xlsx: {}", e.getMessage(), e);
        }
        
        return tableData;
    }
    
    /**
     * Convierte un valor decimal a porcentaje con formato apropiado.
     * Ejemplo: "0.504" -> "50,4%", "1" -> "100,0%", "" -> ""
     */
    private String convertToPercentage(String value) {
        if (value == null || value.trim().isEmpty() || value.equals("0")) {
            return "";
        }
        
        try {
            double decimal = Double.parseDouble(value.trim());
            double percentage = decimal * 100.0;
            
            // Formatear con una decimal y coma como separador decimal
            java.text.DecimalFormat df = new java.text.DecimalFormat("#,##0.0");
            java.text.DecimalFormatSymbols symbols = df.getDecimalFormatSymbols();
            symbols.setDecimalSeparator(',');
            df.setDecimalFormatSymbols(symbols);
            
            return df.format(percentage) + "%";
            
        } catch (NumberFormatException e) {
            // Si no es un número, devolver el valor original
            return value.trim();
        }
    }
    
    /**
     * Obtiene el valor de una celda como String.
     */
    private String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        
        CellType cellType = cell.getCellType();
        if (cellType == CellType.FORMULA) {
            cellType = cell.getCachedFormulaResultType();
        }
        
        switch (cellType) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                // Si es un número entero, no mostrar decimales
                double numValue = cell.getNumericCellValue();
                if (numValue == (long) numValue) {
                    return String.valueOf((long) numValue);
                }
                return String.valueOf(numValue);
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case BLANK:
                return "";
            default:
                return "";
        }
    }
    
    /**
     * Obtiene el valor de una celda como double.
     */
    private double getCellValueAsDouble(Cell cell) {
        if (cell == null) {
            return 0.0;
        }
        
        CellType cellType = cell.getCellType();
        if (cellType == CellType.FORMULA) {
            cellType = cell.getCachedFormulaResultType();
        }
        
        switch (cellType) {
            case NUMERIC:
                return cell.getNumericCellValue();
            case STRING:
                try {
                    String strValue = cell.getStringCellValue().trim();
                    if (strValue.isEmpty()) {
                        return 0.0;
                    }
                    // Intentar parsear como número
                    return Double.parseDouble(strValue.replace("%", "").replace(",", "."));
                } catch (NumberFormatException e) {
                    return 0.0;
                }
            default:
                return 0.0;
        }
    }
}