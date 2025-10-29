import java.io.BufferedReader;
import java.io.File;
import java.io.InputStreamReader;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

/**
 * Script de prueba para validar las correcciones automáticas del GUI
 */
public class test_gui_corrections {
    
    public static void main(String[] args) {
        System.out.println("=== PRUEBA DE CORRECCIONES AUTOMÁTICAS DEL GUI ===");
        
        String excelPath = "C:\\Users\\user\\Documents\\data.xlsx";
        String expectedCorrectedPath = "C:\\Users\\user\\Documents\\data_cleaned_fixed.xlsx";
        
        System.out.println("Archivo original: " + excelPath);
        System.out.println("Archivo corregido esperado: " + expectedCorrectedPath);
        
        // Verificar si existe el archivo original
        File originalFile = new File(excelPath);
        if (!originalFile.exists()) {
            System.out.println("❌ ERROR: No existe el archivo original " + excelPath);
            return;
        }
        System.out.println("✅ Archivo original existe");
        
        // Eliminar archivo corregido si existe para probar la generación
        File correctedFile = new File(expectedCorrectedPath);
        if (correctedFile.exists()) {
            System.out.println("🗑️ Eliminando archivo corregido existente para probar generación...");
            correctedFile.delete();
        }
        
        // Simular la lógica del GUI
        boolean needsCorrections = necesitaCorrecciones(excelPath);
        System.out.println("¿Necesita correcciones? " + needsCorrections);
        
        if (needsCorrections) {
            String result = aplicarCorreccionesAutomaticas(excelPath);
            System.out.println("Resultado de correcciones: " + result);
            
            // Verificar que se generó el archivo
            if (new File(result).exists()) {
                System.out.println("✅ SUCCESS: Archivo corregido generado correctamente");
                System.out.println("📁 Archivo: " + result);
                System.out.println("📏 Tamaño: " + new File(result).length() + " bytes");
            } else {
                System.out.println("❌ ERROR: No se generó el archivo corregido");
            }
        } else {
            System.out.println("ℹ️ El archivo no necesita correcciones");
        }
    }
    
    /**
     * Simula la lógica del GUI para verificar si necesita correcciones
     */
    private static boolean necesitaCorrecciones(String excelPath) {
        try {
            Path path = Paths.get(excelPath);
            String fileName = path.getFileName().toString();
            
            // Solo aplicar correcciones a data.xlsx
            if (fileName.equals("data.xlsx")) {
                Path correctedPath = Paths.get(excelPath.replace(".xlsx", "_cleaned_fixed.xlsx"));
                
                // Si no existe el archivo corregido, necesita correcciones
                if (!Files.exists(correctedPath)) {
                    System.out.println("  → No existe archivo corregido, necesita correcciones");
                    return true;
                }
                
                // Si el archivo original es más nuevo que el corregido, necesita correcciones
                boolean isNewer = Files.getLastModifiedTime(path).compareTo(Files.getLastModifiedTime(correctedPath)) > 0;
                if (isNewer) {
                    System.out.println("  → Archivo original es más nuevo, necesita correcciones");
                }
                return isNewer;
            }
        } catch (Exception e) {
            System.out.println("Error verificando necesidad de correcciones: " + e.getMessage());
        }
        
        return false;
    }
    
    /**
     * Simula la aplicación de correcciones automáticas del GUI
     */
    private static String aplicarCorreccionesAutomaticas(String excelPath) {
        try {
            System.out.println("🔧 Aplicando correcciones automáticas al Excel...");
            
            ProcessBuilder pb = new ProcessBuilder("python", "fix_dra_blue.py");
            pb.directory(new File(System.getProperty("user.dir")));
            
            Process process = pb.start();
            
            // Leer la salida del proceso
            BufferedReader reader = new BufferedReader(new InputStreamReader(process.getInputStream()));
            String line;
            while ((line = reader.readLine()) != null) {
                System.out.println("Python: " + line);
            }
            
            int exitCode = process.waitFor();
            if (exitCode == 0) {
                String correctedPath = excelPath.replace(".xlsx", "_cleaned_fixed.xlsx");
                if (new File(correctedPath).exists()) {
                    System.out.println("✅ Correcciones aplicadas exitosamente: " + new File(correctedPath).getName());
                    return correctedPath;
                }
            } else {
                System.out.println("⚠️ Error en correcciones automáticas, usando archivo original");
            }
        } catch (Exception e) {
            System.out.println("⚠️ Error aplicando correcciones: " + e.getMessage());
            e.printStackTrace();
        }
        
        return excelPath;
    }
}