package com.davivienda.excelpdf.application;

import com.davivienda.excelpdf.infrastructure.PdfOwnershipReportGenerator;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Map;

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
            logger.info("📂 Paso 1: Cargando datos desde Excel...");
            calculator.loadFromExcel(excelPath);
            resultBuilder.withGraphStatistics(calculator.getGraphStatistics());
            logger.info("✅ Datos cargados exitosamente");
            
            // Paso 2: Calcular participaciones finales
            logger.info("🧮 Paso 2: Calculando participaciones finales...");
            calculator.calculateFinalOwnership(rootEntity);
            
            Map<String, Double> finalResults = calculator.getFinalResults();
            Map<String, String> beneficiaryPaths = calculator.getBeneficiaryPaths();
            
            resultBuilder
                .withFinalResults(finalResults)
                .withBeneficiaryPaths(beneficiaryPaths)
                .withRootEntity(rootEntity);
                
            logger.info("✅ Cálculos completados. Beneficiarios encontrados: {}", finalResults.size());
            
            // Paso 3: Generar reporte PDF
            logger.info("📄 Paso 3: Generando reporte PDF...");
            pdfGenerator.generateOwnershipReport(
                finalResults, 
                beneficiaryPaths, 
                rootEntity, 
                outputPdfPath
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
                
            logger.info("✅ PDF generado exitosamente. Tamaño: {} bytes", pdfSize);
            
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
            logger.error("❌ Error durante el procesamiento: {}", e.getMessage(), e);
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
}