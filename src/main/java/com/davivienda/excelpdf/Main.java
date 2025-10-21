package com.davivienda.excelpdf;

import java.io.File;
import java.util.Scanner;

import com.davivienda.excelpdf.application.ExcelOwnershipProcessor;

/**
 * Clase principal del proyecto de Composición Accionaria.
 * 
 * Este programa procesa un archivo Excel con relaciones de propiedad entre entidades y accionistas,
 * calcula las participaciones finales y genera un reporte PDF con los resultados.
 * 
 * <p>Uso desde consola:
 * <pre>
 *   java -jar excel-pdf-processor-standalone.jar <archivo.xlsx> [entidad_raiz]
 * </pre>
 * 
 * Ejemplo:
 * <pre>
 *   java -jar excel-pdf-processor-standalone.jar datos.xlsx "RED COW INC"
 * </pre>
 */
public class Main {

    /**
     * Método principal (punto de entrada del programa).
     */
    public static void main(String[] args) {
        printHeader();

        try {
            // Validar argumentos
            if (args.length == 0) {
                printUsageAndExit();
                return;
            }

            // Obtener archivo Excel
            String excelPath = args[0];
            validateExcelFile(excelPath);

            // Obtener entidad raíz (por argumento o interactivamente)
            String rootEntity = getRootEntity(args);

            // Generar nombre del archivo PDF de salida
            String outputPdfPath = generateOutputPdfPath(excelPath);

            // Mostrar información previa al procesamiento
            System.out.println("\n INFORMACIÓN DEL PROCESAMIENTO:");
            System.out.println("   Archivo Excel : " + new File(excelPath).getAbsolutePath());
            System.out.println("   Tamaño archivo: " + formatFileSize(new File(excelPath).length()));
            System.out.println("   Entidad raíz  : " + rootEntity);
            System.out.println("   PDF salida    : " + outputPdfPath);
            System.out.println("─".repeat(50));

            // Ejecutar el procesamiento principal
            ExcelOwnershipProcessor processor = new ExcelOwnershipProcessor();
            ExcelOwnershipProcessor.ProcessingResult result =
                    processor.processOwnershipAnalysis(excelPath, rootEntity, outputPdfPath);

            // Mostrar resultados finales
            printResults(result);

        } catch (Exception e) {
            System.err.println("\n ERROR DURANTE EL PROCESAMIENTO");
            System.err.println("═".repeat(50));
            System.err.println("Tipo de error: " + e.getClass().getSimpleName());
            System.err.println("Mensaje: " + e.getMessage());
            System.err.println("\n Traza de error:");
            e.printStackTrace();
            System.err.println("\n Sugerencias:");
            System.err.println("   • Verifique que el archivo Excel tenga el formato correcto");
            System.err.println("   • Asegúrese de que la entidad raíz exista en el Excel");
            System.err.println("   • Revise que el archivo no esté abierto en otra aplicación");
            System.exit(2);
        }
    }

    // ============================================================
    // =============== SECCIÓN: MÉTODOS DE IMPRESIÓN ===============
    // ============================================================

    /**
     * Imprime el encabezado de la aplicación.
     */
    private static void printHeader() {
        System.out.println("════════════════════════════════════════════════════");
        System.out.println("        ANALISIS DE COMPOSICION ACCIONARIA          ");
        System.out.println("                    Davivienda                      ");
        System.out.println("════════════════════════════════════════════════════");
        System.out.println("Version: 1.0.0");
        System.out.println("Java: " + System.getProperty("java.version"));
        System.out.println("OS: " + System.getProperty("os.name") + " " + System.getProperty("os.arch"));
        System.out.println("════════════════════════════════════════════════════");
    }

    /**
     * Imprime las instrucciones de uso y termina la aplicación.
     */
    private static void printUsageAndExit() {
        System.err.println(" Error: Falta el archivo Excel\n");
        System.err.println(" USO:");
        System.err.println("   java -jar excel-pdf-processor-standalone.jar <archivo.xlsx> [entidad_raiz]\n");
        System.err.println(" PARÁMETROS:");
        System.err.println("   archivo.xlsx  : Archivo Excel con las relaciones de propiedad");
        System.err.println("   entidad_raiz  : (Opcional) Entidad desde la cual calcular participaciones\n");
        System.err.println(" EJEMPLOS:");
        System.err.println("   java -jar excel-pdf-processor-standalone.jar datos.xlsx");
        System.err.println("   java -jar excel-pdf-processor-standalone.jar datos.xlsx \"RED COW INC\"\n");
        System.err.println(" FORMATO DEL EXCEL:");
        System.err.println("   Columna A: Entidad");
        System.err.println("   Columna B: Accionista");
        System.err.println("   Columna C: % Participación\n");
        System.exit(1);
    }

    /**
     * Imprime ejemplos adicionales de uso del programa.
     */
    private static void printUsageExamples() {
        System.err.println(" EJEMPLOS DE USO:\n");
        System.err.println("   Modo interactivo:");
        System.err.println("   java -jar excel-pdf-processor-standalone.jar estructura.xlsx\n");
        System.err.println("   Con entidad específica:");
        System.err.println("   java -jar excel-pdf-processor-standalone.jar estructura.xlsx \"Empresa A\"\n");
        System.err.println("   Con archivo de salida personalizado:");
        System.err.println("   java -jar excel-pdf-processor-standalone.jar estructura.xlsx \"Empresa A\" mi_reporte.pdf\n");
    }

    // ============================================================
    // ============== SECCIÓN: VALIDACIONES Y UTILIDADES ==========
    // ============================================================

    /**
     * Valida que el archivo Excel exista, sea legible y tenga extensión .xlsx.
     */
    private static void validateExcelFile(String excelPath) {
        File excelFile = new File(excelPath);

        if (!excelFile.exists()) {
            System.err.println(" Error: El archivo no existe: " + excelPath + "\n");
            System.err.println(" Verifique que:");
            System.err.println("   • La ruta del archivo sea correcta");
            System.err.println("   • El archivo tenga extensión .xlsx");
            System.err.println("   • Tenga permisos de lectura\n");
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
     * Obtiene la entidad raíz desde los argumentos o la solicita al usuario.
     */
    private static String getRootEntity(String[] args) {
        if (args.length >= 2 && !args[1].trim().isEmpty()) {
            return args[1].trim();
        }

        System.out.println(" Ingrese la entidad raíz para el análisis:");
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
     * Genera el nombre del archivo PDF de salida basado en el nombre del archivo Excel.
     */
    private static String generateOutputPdfPath(String excelPath) {
        File excelFile = new File(excelPath);
        String fileName = excelFile.getName();
        String nameWithoutExtension = fileName.substring(0, fileName.lastIndexOf('.'));
        String directory = excelFile.getParent();

        String pdfFileName = nameWithoutExtension + "_composicion_accionaria.pdf";
        return (directory != null)
                ? new File(directory, pdfFileName).getAbsolutePath()
                : pdfFileName;
    }

    // ============================================================
    // =============== SECCIÓN: RESULTADOS Y FORMATO ==============
    // ============================================================

    /**
     * Imprime los resultados del procesamiento en consola.
     */
    private static void printResults(ExcelOwnershipProcessor.ProcessingResult result) {
        System.out.println("\n PROCESAMIENTO COMPLETADO EXITOSAMENTE");
        System.out.println("═".repeat(50));

        System.out.println("\n ESTADÍSTICAS DEL ANÁLISIS:");
        System.out.println("   " + result.getGraphStatistics());
        System.out.println("   Beneficiarios finales: " + result.getFinalResults().size());
        System.out.println("   Tiempo de procesamiento: " + result.getProcessingTime() + " ms");
        System.out.println("   Velocidad: " + String.format("%.2f", 1000.0 / result.getProcessingTime()) + " análisis/segundo");

        System.out.println("\n ARCHIVO GENERADO:");
        System.out.println("   Ubicación: " + result.getOutputPdfPath());
        System.out.println("   Tamaño: " + formatFileSize(result.getPdfSize()));

        System.out.println("\n TOP 5 BENEFICIARIOS:");
        if (result.getFinalResults().isEmpty()) {
            System.out.println("   No se encontraron beneficiarios finales");
        } else {
            result.getFinalResults().entrySet().stream()
                    .sorted(java.util.Map.Entry.<String, Double>comparingByValue().reversed())
                    .limit(5)
                    .forEach(entry -> {
                        String beneficiary = entry.getKey();
                        double percentage = entry.getValue();
                        String bar = generateProgressBar(percentage, 20);
                        System.out.printf("   • %-30s %8.4f%% %s%n",
                                truncateString(beneficiary, 30), percentage * 100, bar);
                    });
        }

        System.out.println("\n ¡Análisis completado! Revise el archivo PDF para más detalles.");
        System.out.println("═".repeat(50));
    }

    /**
     * Formatea un tamaño de archivo (bytes) en una cadena legible (KB o MB).
     */
    private static String formatFileSize(long bytes) {
        if (bytes < 1024) return bytes + " bytes";
        if (bytes < 1024 * 1024) return String.format("%.1f KB", bytes / 1024.0);
        return String.format("%.1f MB", bytes / (1024.0 * 1024.0));
    }

    /**
     * Trunca una cadena si supera la longitud máxima especificada.
     */
    private static String truncateString(String str, int maxLength) {
        if (str.length() <= maxLength) return str;
        return str.substring(0, maxLength - 3) + "...";
    }

    /**
     * Genera una barra de progreso visual para representar un porcentaje.
     */
    private static String generateProgressBar(double percentage, int width) {
        int filled = (int) Math.round(percentage * width);
        if (filled > width) filled = width;
        if (filled < 0) filled = 0;
        
        StringBuilder bar = new StringBuilder("[");
        for (int i = 0; i < width; i++) {
            bar.append(i < filled ? "█" : "░");
        }
        bar.append("]");
        return bar.toString();
    }
}
