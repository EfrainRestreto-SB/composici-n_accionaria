package com.davivienda.excelpdf;

import java.io.BufferedReader;
import java.io.File;
import java.io.InputStreamReader;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Scanner;

import javax.swing.SwingUtilities;

import com.davivienda.excelpdf.application.ExcelOwnershipProcessor;
import com.davivienda.excelpdf.ui.ComposicionAccionariaGUI;

/**
 * Clase principal del proyecto de Composición Accionaria.
 * 
 * Este programa procesa un archivo Excel con relaciones de propiedad entre entidades y accionistas,
 * calcula las participaciones finales y genera un reporte PDF con los resultados.
 * 
 * <p>Uso desde consola:
 * <pre>
 *   java -jar composicion-accionaria-optimizado.jar &lt;archivo.xlsx&gt; [entidad_raiz]
 * </pre>
 * 
 * Ejemplo:
 * <pre>
 *   java -jar composicion-accionaria-optimizado.jar datos.xlsx "RED COW INC"
 * </pre>
 * 
 * @author Davivienda
 * @version 1.0.0
 */
public class Main {
    
    // ============================================================
    // ====================== CONSTANTES =========================
    // ============================================================
    
    private static final String APP_VERSION = "1.0.0";
    private static final String XLSX_EXTENSION = ".xlsx";
    private static final String CLEANED_SUFFIX = "_cleaned_fixed.xlsx";
    private static final String PDF_PREFIX = "composicion_accionaria_";
    private static final String TIME_PATTERN = "HHmmss";
    private static final int PROGRESS_BAR_WIDTH = 20;
    private static final int MAX_DISPLAY_LENGTH = 30;
    private static final int TOP_BENEFICIARIES_LIMIT = 5;

    /**
     * Método principal (punto de entrada del programa).
     */
    public static void main(String[] args) {
        // Si no hay argumentos, lanzar interfaz gráfica
        if (args.length == 0) {
            launchGUI();
            return;
        }

        // Modo consola
        printHeader();

        try {
            // Validar argumentos
            if (args.length < 1) {
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

            // Aplicar correcciones automáticas al Excel si es necesario
            String correctedExcelPath = applyExcelCorrections(excelPath);

            // Ejecutar el procesamiento principal
            ExcelOwnershipProcessor processor = new ExcelOwnershipProcessor();
            ExcelOwnershipProcessor.ProcessingResult result =
                    processor.processOwnershipAnalysis(correctedExcelPath, rootEntity, outputPdfPath);

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
        System.out.println("Version: " + APP_VERSION);
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
        System.err.println("   java -jar composicion-accionaria-optimizado.jar <archivo.xlsx> [entidad_raiz]\n");
        System.err.println(" PARÁMETROS:");
        System.err.println("   archivo.xlsx  : Archivo Excel con las relaciones de propiedad");
        System.err.println("   entidad_raiz  : (Opcional) Entidad desde la cual calcular participaciones\n");
        System.err.println(" EJEMPLOS:");
        System.err.println("   java -jar composicion-accionaria-optimizado.jar datos.xlsx");
        System.err.println("   java -jar composicion-accionaria-optimizado.jar datos.xlsx \"RED COW INC\"\n");
        System.err.println(" FORMATO DEL EXCEL:");
        System.err.println("   Columna A: Entidad");
        System.err.println("   Columna B: Accionista");
        System.err.println("   Columna C: % Participación\n");
        System.exit(1);
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
            System.exit(1);
        }

        if (!excelFile.canRead()) {
            System.err.println(" Error: No se puede leer el archivo: " + excelPath);
            System.err.println(" Verifique los permisos del archivo");
            System.exit(1);
        }

        if (!excelPath.toLowerCase().endsWith(XLSX_EXTENSION)) {
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
     * Genera el nombre del archivo PDF de salida con timestamp HHMMSS.
     */
    private static String generateOutputPdfPath(String excelPath) {
        File excelFile = new File(excelPath);
        String directory = excelFile.getParent();
        
        // Generar timestamp HHMMSS
        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern(TIME_PATTERN));
        String pdfFileName = PDF_PREFIX + timestamp + ".pdf";
        
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

        System.out.println("\n TOP " + TOP_BENEFICIARIES_LIMIT + " BENEFICIARIOS:");
        if (result.getFinalResults().isEmpty()) {
            System.out.println("   No se encontraron beneficiarios finales");
        } else {
            result.getFinalResults().entrySet().stream()
                    .sorted(java.util.Map.Entry.<String, Double>comparingByValue().reversed())
                    .limit(TOP_BENEFICIARIES_LIMIT)
                    .forEach(entry -> {
                        String beneficiary = entry.getKey();
                        double percentage = entry.getValue();
                        String bar = generateProgressBar(percentage, PROGRESS_BAR_WIDTH);
                        System.out.printf("   • %-30s %8.4f%% %s%n",
                                truncateString(beneficiary, MAX_DISPLAY_LENGTH), percentage * 100, bar);
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

    /**
     * Lanza la interfaz gráfica (GUI)
     */
    private static void launchGUI() {
        System.out.println("Iniciando interfaz gráfica...");
        
        // Configurar Look and Feel del sistema
        try {
            javax.swing.UIManager.setLookAndFeel(
                javax.swing.UIManager.getSystemLookAndFeelClassName()
            );
        } catch (ClassNotFoundException | InstantiationException | 
                 IllegalAccessException | javax.swing.UnsupportedLookAndFeelException e) {
            // Si falla, usar el Look and Feel por defecto
            System.err.println("No se pudo configurar el Look and Feel del sistema: " + e.getMessage());
        }

        // Lanzar GUI en el Event Dispatch Thread
        SwingUtilities.invokeLater(() -> {
            ComposicionAccionariaGUI gui = new ComposicionAccionariaGUI();
            gui.setVisible(true);
        });
    }
    
    // ============================================================
    // ============ SECCIÓN: CORRECCIONES AUTOMÁTICAS =============
    // ============================================================
    
    /**
     * Aplica correcciones automáticas al archivo Excel si es necesario.
     * Por ejemplo, unifica entidades DRA BLUE con nombres similares.
     */
    private static String applyExcelCorrections(String originalExcelPath) {
        try {
            // Verificar si el archivo necesita correcciones
            if (!needsCorrections(originalExcelPath)) {
                System.out.println("   ✓ Excel no requiere correcciones automáticas");
                return originalExcelPath;
            }
            
            System.out.println("   🔧 Aplicando correcciones automáticas al Excel...");
            
            // Generar ruta del archivo corregido
            String correctedPath = originalExcelPath.replace(XLSX_EXTENSION, CLEANED_SUFFIX);
            
            // Ejecutar el script de corrección usando ProcessBuilder
            ProcessBuilder pb = new ProcessBuilder("python", "fix_dra_blue_dynamic.py", originalExcelPath);
            pb.directory(new File("."));
            pb.redirectErrorStream(true);
            
            Process process = pb.start();
            
            // Leer la salida del proceso
            try (BufferedReader reader = new BufferedReader(new InputStreamReader(process.getInputStream()))) {
                String line;
                while ((line = reader.readLine()) != null) {
                    System.out.println("   [Python] " + line);
                }
            }
            
            int exitCode = process.waitFor();
            
            if (exitCode == 0) {
                // Verificar que el archivo corregido existe
                File correctedFile = new File(correctedPath);
                if (correctedFile.exists()) {
                    System.out.println("   ✓ Correcciones aplicadas exitosamente: " + correctedFile.getName());
                    return correctedPath;
                } else {
                    System.out.println("   ⚠ Archivo corregido no encontrado, usando original");
                    return originalExcelPath;
                }
            } else {
                System.out.println("   ⚠ Error en corrección automática, usando archivo original");
                return originalExcelPath;
            }
            
        } catch (java.io.IOException | InterruptedException e) {
            System.out.println("   ⚠ Error aplicando correcciones: " + e.getMessage());
            System.out.println("   📋 Usando archivo original sin correcciones");
            return originalExcelPath;
        }
    }
    
    /**
     * Verifica si el archivo Excel necesita correcciones automáticas.
     * Detecta casos conocidos como entidades DRA BLUE duplicadas.
     */
    private static boolean needsCorrections(String excelPath) {
        // Verificar si ya existe una versión corregida
        String correctedPath = excelPath.replace(XLSX_EXTENSION, CLEANED_SUFFIX);
        File correctedFile = new File(correctedPath);
        
        // Si ya existe el archivo corregido y es más reciente, no necesita correcciones
        File originalFile = new File(excelPath);
        if (correctedFile.exists() && 
            correctedFile.lastModified() >= originalFile.lastModified()) {
            return false;
        }
        
        // Para simplificar, siempre aplicamos correcciones para data.xlsx
        return excelPath.contains("data.xlsx");
    }
}
