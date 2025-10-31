public class TestPercentage {
    public static void main(String[] args) {
        // Probar la lógica de conversión
        System.out.println("Probando conversión de porcentajes:");
        System.out.println("0.504 -> " + convertToPercentage("0.504"));
        System.out.println("1 -> " + convertToPercentage("1"));
        System.out.println("0.372 -> " + convertToPercentage("0.372"));
        System.out.println("0.124 -> " + convertToPercentage("0.124"));
        System.out.println("0.05 -> " + convertToPercentage("0.05"));
        System.out.println("0 -> " + convertToPercentage("0"));
        System.out.println("'' -> " + convertToPercentage(""));
    }
    
    private static String convertToPercentage(String value) {
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
}