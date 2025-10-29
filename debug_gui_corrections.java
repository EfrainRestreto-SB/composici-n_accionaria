import java.io.BufferedReader;
import java.io.File;
import java.io.InputStreamReader;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

/**
 * Debug del sistema de correcciones automáticas del GUI
 */
public class debug_gui_corrections {
    
    public static void main(String[] args) {
        System.out.println("=== DEBUG: CORRECCIONES AUTOMÁTICAS DEL GUI ===");
        
        String excelPath = "C:\\Users\\user\\Documents\\data.xlsx";
        
        System.out.println("1. Verificando archivo original:");
        File originalFile = new File(excelPath);
        System.out.println("   Existe: " + originalFile.exists());
        System.out.println("   Tamaño: " + originalFile.length() + " bytes");
        
        System.out.println("\n2. Ejecutando lógica de necesitaCorrecciones():");
        boolean needsCorrections = necesitaCorrecciones(excelPath);
        System.out.println("   ¿Necesita correcciones? " + needsCorrections);
        
        System.out.println("\n3. Ejecutando aplicarCorreccionesAutomaticas():");
        String result = aplicarCorreccionesAutomaticas(excelPath);
        System.out.println("   Archivo resultado: " + result);
        
        System.out.println("\n4. Verificando archivo corregido:");
        String correctedPath = excelPath.replace(".xlsx", "_cleaned_fixed.xlsx");
        File correctedFile = new File(correctedPath);
        System.out.println("   Archivo esperado: " + correctedPath);
        System.out.println("   Existe: " + correctedFile.exists());
        if (correctedFile.exists()) {
            System.out.println("   Tamaño: " + correctedFile.length() + " bytes");
        }
        
        System.out.println("\n5. Resultado final:");
        if (result.equals(correctedPath) && correctedFile.exists()) {
            System.out.println("   ✅ SUCCESS: Las correcciones funcionan correctamente");
        } else {
            System.out.println("   ❌ ERROR: Las correcciones NO funcionan");
            System.out.println("   Esperado: " + correctedPath);
            System.out.println("   Obtenido: " + result);
        }
    }
    
    private static boolean necesitaCorrecciones(String excelPath) {
        try {
            Path path = Paths.get(excelPath);
            String fileName = path.getFileName().toString();
            
            System.out.println("   - Verificando archivo: " + fileName);
            
            // Solo aplicar correcciones a data.xlsx
            if (fileName.equals("data.xlsx")) {
                Path correctedPath = Paths.get(excelPath.replace(".xlsx", "_cleaned_fixed.xlsx"));
                
                System.out.println("   - Buscando archivo corregido: " + correctedPath.getFileName());
                
                // Si no existe el archivo corregido, necesita correcciones
                if (!Files.exists(correctedPath)) {
                    System.out.println("   - No existe archivo corregido → SE REQUIEREN correcciones");
                    return true;
                }
                
                // Si el archivo original es más nuevo que el corregido, necesita correcciones
                boolean isNewer = Files.getLastModifiedTime(path).compareTo(Files.getLastModifiedTime(correctedPath)) > 0;
                if (isNewer) {
                    System.out.println("   - Archivo original es más nuevo → SE REQUIEREN correcciones");
                } else {
                    System.out.println("   - Archivo corregido está actualizado → NO se requieren correcciones");
                }
                return isNewer;
            } else {
                System.out.println("   - No es data.xlsx → NO se requieren correcciones");
            }
        } catch (Exception e) {
            System.out.println("   - ERROR: " + e.getMessage());
            e.printStackTrace();
        }
        
        return false;
    }
    
    private static String aplicarCorreccionesAutomaticas(String excelPath) {
        try {
            System.out.println("   - Evaluando necesidad de correcciones para: " + excelPath);
            
            boolean needsCorrections = necesitaCorrecciones(excelPath);
            System.out.println("   - ¿Necesita correcciones? " + needsCorrections);
            
            if (needsCorrections) {
                System.out.println("   - 🔧 Aplicando correcciones automáticas al Excel...");
                
                ProcessBuilder pb = new ProcessBuilder("python", "fix_dra_blue.py");
                pb.directory(new File(System.getProperty("user.dir")));
                
                Process process = pb.start();
                
                // Leer la salida del proceso
                BufferedReader reader = new BufferedReader(new InputStreamReader(process.getInputStream()));
                String line;
                while ((line = reader.readLine()) != null) {
                    System.out.println("   Python: " + line);
                }
                
                int exitCode = process.waitFor();
                if (exitCode == 0) {
                    String correctedPath = excelPath.replace(".xlsx", "_cleaned_fixed.xlsx");
                    if (new File(correctedPath).exists()) {
                        System.out.println("   - ✅ Correcciones aplicadas exitosamente: " + new File(correctedPath).getName());
                        System.out.println("   - 📁 Archivo corregido en: " + correctedPath);
                        return correctedPath;
                    } else {
                        System.out.println("   - ❌ Error: El archivo corregido no se generó");
                    }
                } else {
                    System.out.println("   - ⚠️ Error en correcciones automáticas (código: " + exitCode + "), usando archivo original");
                }
            } else {
                System.out.println("   - ℹ️ El archivo no necesita correcciones automáticas");
            }
        } catch (Exception e) {
            System.out.println("   - ⚠️ Error aplicando correcciones: " + e.getMessage());
            e.printStackTrace();
        }
        
        return excelPath;
    }
}