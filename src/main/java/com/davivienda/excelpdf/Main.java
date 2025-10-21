package com.davivienda.excelpdf;

import java.io.File;
import java.util.Scanner;

import com.davivienda.excelpdf.application.ExcelOwnershipProcessor;

/**
 * Clase principal para el proyecto de Composición Accionaria
 * Procesa archivos Excel y genera reportes PDF
 */
public class Main {

    public static void main(String[] args) {
        printHeader();
        try {
            // Validar argumentos
            if (args.length == 0) {
                printUsageAndExit();
                return;
            }

            String excelPath = args[0];
            validateExcelFile(excelPath);

            // Obtener entidad raíz (del argumento o interactivamente)
            String rootEntity = getRootEntity(args);

            // Generar nombre del archivo PDF de salida
            String outputPdfPath = generateOutputPdfPath(excelPath);

            System.out.println(" Procesando: " + new File(excelPath).getAbsolutePath());
            System.out.println(" Entidad raíz: " + rootEntity);
            System.out.println(" PDF salida: " + outputPdfPath);

            // Procesar análisis de composición accionaria
            ExcelOwnershipProcessor processor = new ExcelOwnershipProcessor();
            ExcelOwnershipProcessor.ProcessingResult result = processor.processOwnershipAnalysis(
                    excelPath,
                    rootEntity,
                    outputPdfPath);

            // Mostrar resultados
            printResults(result);

        } catch (Exception e) {
            System.err.println(" Error durante el procesamiento");
            System.err.println("   " + e.getMessage());
            e.printStackTrace();
            System.exit(2);
        }
    }

    /**
     * Imprime el encabezado de la aplicación.
     */
    private static void printHeader() {
        System.out.println("====================================================");
        System.out.println("        ANALISIS DE COMPOSICION ACCIONARIA        ");
        System.out.println("                  Davivienda                      ");
        System.out.println("====================================================");
        System.out.println("Version: 1.0.0");
        System.out.println("Java: " + System.getProperty("java.version"));
        System.out.println("====================================================");
    }

    /**
     * Imprime instrucciones de uso y termina la aplicación.
     */
    private static void printUsageAndExit() {
        System.err.println(" Error: Falta el archivo Excel");
        System.err.println("");
        System.err.println(" USO:");
        System.err.println("   java -jar excel-pdf-processor-standalone.jar <archivo.xlsx> [entidad_raiz]");
        System.err.println("");
        System.err.println(" PARÁMETROS:");
        System.err.println("   archivo.xlsx  : Archivo Excel con las relaciones de propiedad");
        System.err.println("   entidad_raiz  : (Opcional) Entidad desde la cual calcular participaciones");
        System.err.println("");
        System.err.println(" EJEMPLOS:");
        System.err.println("   java -jar excel-pdf-processor-standalone.jar datos.xlsx");
        System.err.println("   java -jar excel-pdf-processor-standalone.jar datos.xlsx \"RED COW INC\"");
        System.err.println("");
        System.err.println(" FORMATO DEL EXCEL:");
        System.err.println("   Columna A: Entidad");
        System.err.println("   Columna B: Accionista");
        System.err.println("   Columna C: % Participación");
        System.exit(1);
    }

    /**
     * Valida que el archivo Excel existe y es válido.
     */
    private static void validateExcelFile(String excelPath) {
        File excelFile = new File(excelPath);

        if (!excelFile.exists()) {
            System.err.println(" Error: El archivo no existe: " + excelPath);
            System.err.println();
            System.err.println(" Verifique que:");
            System.err.println("   • La ruta del archivo sea correcta");
            System.err.println("   • El archivo tenga extensión .xlsx");
            System.err.println("   • Tenga permisos de lectura");
            System.err.println();
            printUsageExamples();
            System.exit(1);
        }

        if (!excelFile.canRead()) {
            System.err.println(" Error: No se puede leer el archivo: " + excelPath);
            System.err.println(" Verifique los permisos del archivo");
            System.exit(1);
        }

        if (!excelPath.toLowerCase().endsWith(".xlsx")) {
            System.err.println(" Error: El archivo debe tener extensión .xlsx: " + excelPath);
            System.err.println("  Nota: No se soportan archivos .xls (formato antiguo)");
            System.exit(1);
        }
    }

    /**
     * Imprime ejemplos de uso del programa.
     */
    private static void printUsageExamples() {
        System.err.println(" EJEMPLOS DE USO:");
        System.err.println();
        System.err.println("   Modo interactivo:");
        System.err.println("   java -jar excel-pdf-processor-standalone.jar estructura.xlsx");
        System.err.println();
        System.err.println("   Con entidad específica:");
        System.err.println("   java -jar excel-pdf-processor-standalone.jar estructura.xlsx \"Empresa A\"");
        System.err.println();
        System.err.println("   Con archivo de salida personalizado:");
        System.err.println(
                "   java -jar excel-pdf-processor-standalone.jar estructura.xlsx \"Empresa A\" mi_reporte.pdf");
        System.err.println();
    }

    /**
     * Obtiene la entidad raíz desde argumentos o de forma interactiva.
     */
    private static String getRootEntity(String[] args) {
        if (args.length >= 2 && !args[1].trim().isEmpty()) {
            return args[1].trim();
        }

        // Solicitar entidad raíz interactivamente
        System.out.println("🏢 Ingrese la entidad raíz para el análisis:");
        System.out.print("   > ");

        try (Scanner scanner = new Scanner(System.in)) {
            String rootEntity = scanner.nextLine().trim();

            if (rootEntity.isEmpty()) {
                System.err.println(" Error: La entidad raíz no puede estar vacía");
                System.exit(1);
            }

            return rootEntity;
        }
    }

    /**
     * Genera el nombre del archivo PDF de salida basado en el archivo Excel.
     */
    private static String generateOutputPdfPath(String excelPath) {
        File excelFile = new File(excelPath);
        String fileName = excelFile.getName();
        String nameWithoutExtension = fileName.substring(0, fileName.lastIndexOf('.'));
        String directory = excelFile.getParent();

        String pdfFileName = nameWithoutExtension + "_composicion_accionaria.pdf";

        if (directory != null) {
            return new File(directory, pdfFileName).getAbsolutePath();
        } else {
            return pdfFileName;
        }
    }
    /**
     * Imprime los resultados del procesamiento.
     */
    private static void printResults(ExcelOwnershipProcessor.ProcessingResult result) {
        System.out.println("");
        System.out.println("✅ PROCESAMIENTO COMPLETADO EXITOSAMENTE");
        System.out.println("═".repeat(50));
        
        System.out.println("📊 ESTADÍSTICAS:");
        System.out.println("   " + result.getGraphStatistics());
        System.out.println("   Beneficiarios finales: " + result.getFinalResults().size());
        System.out.println("   Tiempo de procesamiento: " + result.getProcessingTime() + " ms");
        
        System.out.println("");
        System.out.println("📄 ARCHIVO GENERADO:");
        System.out.println("   Ubicación: " + result.getOutputPdfPath());
        System.out.println("   Tamaño: " + formatFileSize(result.getPdfSize()));
        
        System.out.println("");
        System.out.println("🏆 TOP 5 BENEFICIARIOS:");
        result.getFinalResults().entrySet().stream()
            .sorted(java.util.Map.Entry.<String, Double>comparingByValue().reversed())
            .limit(5)
            .forEach(entry -> {
                String beneficiary = entry.getKey();
                double percentage = entry.getValue();
                System.out.printf("   • %-30s %8.4f%%\n", 
                    truncateString(beneficiary, 30), percentage * 100);
            });
        
        System.out.println("");
        System.out.println("🎉 ¡Análisis completado! Revise el archivo PDF para más detalles.");
    }

    /**
     * Formatea el tamaño de archivo en formato legible.
     */
    private static String formatFileSize(long bytes) {
        if (bytes < 1024) return bytes + " bytes";
        if (bytes < 1024 * 1024) return String.format("%.1f KB", bytes / 1024.0);
        return String.format("%.1f MB", bytes / (1024.0 * 1024.0));
    }
    
    /**
     * Trunca una cadena si es muy larga.
     */
    private static String truncateString(String str, int maxLength) {
        if (str.length() <= maxLength) return str;
        return str.substring(0, maxLength - 3) + "...";
    }
}