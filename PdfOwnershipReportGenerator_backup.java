package com.davivienda.excelpdf.infrastructure;

import java.awt.Color;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Map;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.lowagie.text.Chunk;
import com.lowagie.text.Document;
import com.lowagie.text.DocumentException;
import com.lowagie.text.Element;
import com.lowagie.text.Font;
import com.lowagie.text.FontFactory;
import com.lowagie.text.Image;
import com.lowagie.text.PageSize;
import com.lowagie.text.Paragraph;
import com.lowagie.text.Phrase;
import com.lowagie.text.Rectangle;
import com.lowagie.text.pdf.ColumnText;
import com.lowagie.text.pdf.PdfContentByte;
import com.lowagie.text.pdf.PdfPCell;
import com.lowagie.text.pdf.PdfPTable;
import com.lowagie.text.pdf.PdfPageEventHelper;
import com.lowagie.text.pdf.PdfWriter;

/**
 * Generador de reportes PDF para análisis de composición accionaria.
 * Utiliza OpenPDF para crear documentos con formato profesional.
 * 
 * @author Davivienda
 * @version 1.0
 */
public class PdfOwnershipReportGenerator {
    
    private static final Logger logger = LoggerFactory.getLogger(PdfOwnershipReportGenerator.class);
    
    private static final DecimalFormat PERCENTAGE_FORMAT = new DecimalFormat("#0.00%");
    private static final DateTimeFormatter DATE_FORMAT = DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm:ss");
    
    // Colores corporativos
    private static final Color DAVIVIENDA_RED = new Color(204, 0, 51);
    private static final Color HEADER_GRAY = new Color(128, 128, 128);
    private static final Color LIGHT_GRAY = new Color(245, 245, 245);
    
    /**
     * Genera un reporte PDF con los resultados del análisis de composición accionaria.
     * 
     * @param finalResults mapa de beneficiario -> porcentaje final
     * @param beneficiaryPaths mapa de beneficiario -> ruta completa
     * @param rootEntity nombre de la entidad raíz analizada
     * @param outputPath ruta del archivo PDF de salida
     * @param originalData datos originales del Excel para el desglose detallado
     * @throws IOException si hay problemas escribiendo el archivo
     */
    public void generateOwnershipReport(Map<String, Double> finalResults,
                                      Map<String, String> beneficiaryPaths,
                                      String rootEntity,
                                      String outputPath,
                                      Map<String, Map<String, Double>> originalData) throws IOException {
        
        logger.info("Generando reporte PDF: {}", outputPath);
        
        try {
            Document document = new Document(PageSize.A4, 50, 50, 50, 70);
            PdfWriter writer = PdfWriter.getInstance(document, new FileOutputStream(outputPath));
            
            // Configurar evento para pie de página
            writer.setPageEvent(new FooterPageEvent());
            
            // Configurar metadatos
            document.addTitle("Análisis de Composición Accionaria - " + rootEntity);
            document.addAuthor("Davivienda");
            document.addSubject("Composición Accionaria");
            document.addCreator("Sistema de Análisis Davivienda v1.0");
            
            document.open();
            
            // Agregar contenido al documento
            addHeader(document, rootEntity);
            addSummary(document, finalResults, rootEntity);
            
            // NUEVA SECCIÓN: Desglose detallado de composición accionaria
            if (originalData != null && !originalData.isEmpty()) {
                addDetailedBreakdown(document, originalData, rootEntity);
            }
            
            addDetailedResults(document, finalResults, beneficiaryPaths);
            addFooter(document, writer);
            
            document.close();
            logger.info("Reporte PDF generado exitosamente: {}", outputPath);
            
        } catch (DocumentException e) {
            logger.error("Error generando PDF: {}", e.getMessage(), e);
            throw new IOException("Error creando el documento PDF: " + e.getMessage(), e);
        }
    }
    
    /**
     * Genera un reporte PDF con los resultados del análisis de composición accionaria.
     * Versión de compatibilidad hacia atrás sin datos originales.
     */
    public void generateOwnershipReport(Map<String, Double> finalResults,
                                      Map<String, String> beneficiaryPaths,
                                      String rootEntity,
                                      String outputPath) throws IOException {
        generateOwnershipReport(finalResults, beneficiaryPaths, rootEntity, outputPath, null);
    }
    
    /**
     * Agrega el encabezado del documento con logo.
     */
    private void addHeader(Document document, String rootEntity) throws DocumentException {
        try {
            // Tabla para el encabezado (logo a la derecha, texto a la izquierda)
            PdfPTable headerTable = new PdfPTable(2);
            headerTable.setWidthPercentage(100);
            headerTable.setWidths(new float[]{70, 30});
            
            // Celda izquierda: Títulos
            PdfPCell leftCell = new PdfPCell();
            leftCell.setBorder(PdfPCell.NO_BORDER);
            
            Font titleFont = FontFactory.getFont(FontFactory.HELVETICA_BOLD, 18, DAVIVIENDA_RED);
            Paragraph title = new Paragraph("ANÁLISIS DE COMPOSICIÓN ACCIONARIA", titleFont);
            title.setAlignment(Element.ALIGN_LEFT);
            
            Font subtitleFont = FontFactory.getFont(FontFactory.HELVETICA, 14, Color.BLACK);
            Paragraph subtitle = new Paragraph("Entidad Analizada: " + rootEntity, subtitleFont);
            subtitle.setAlignment(Element.ALIGN_LEFT);
            subtitle.setSpacingBefore(5);
            
            leftCell.addElement(title);
            leftCell.addElement(subtitle);
            leftCell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            
            // Celda derecha: Logo
            PdfPCell rightCell = new PdfPCell();
            rightCell.setBorder(PdfPCell.NO_BORDER);
            rightCell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            rightCell.setHorizontalAlignment(Element.ALIGN_RIGHT);
            
            try {
                // Intentar cargar el logo desde el directorio del JAR
                String logoPath = "Imagen1.png";
                File logoFile = new File(logoPath);
                
                if (logoFile.exists()) {
                    Image logo = Image.getInstance(logoPath);
                    // Ajustar tamaño del logo (escalado al 30%)
                    logo.scalePercent(30);
                    logo.setAlignment(Element.ALIGN_RIGHT);
                    rightCell.addElement(logo);
                } else {
                    logger.warn("Logo no encontrado en: {}", logoPath);
                }
            } catch (Exception e) {
                logger.warn("No se pudo cargar el logo: {}", e.getMessage());
            }
            
            headerTable.addCell(leftCell);
            headerTable.addCell(rightCell);
            
            document.add(headerTable);
            
        } catch (Exception e) {
            logger.error("Error agregando encabezado con logo: {}", e.getMessage());
            // Fallback al encabezado simple
            Font titleFont = FontFactory.getFont(FontFactory.HELVETICA_BOLD, 18, DAVIVIENDA_RED);
            Paragraph title = new Paragraph("ANÁLISIS DE COMPOSICIÓN ACCIONARIA", titleFont);
            title.setAlignment(Element.ALIGN_CENTER);
            document.add(title);
        }
        
        // Fecha de generación
        Font dateFont = FontFactory.getFont(FontFactory.HELVETICA, 10, HEADER_GRAY);
        Paragraph date = new Paragraph("Fecha de Generación: " + LocalDateTime.now().format(DATE_FORMAT), dateFont);
        date.setAlignment(Element.ALIGN_RIGHT);
        date.setSpacingAfter(30);
        date.setSpacingBefore(10);
        document.add(date);
        
        // Línea separadora
        document.add(new Paragraph("_".repeat(80), dateFont));
        document.add(Chunk.NEWLINE);
    }
    
    /**
     * Agrega el resumen ejecutivo.
     */
    private void addSummary(Document document, Map<String, Double> finalResults, String rootEntity) 
            throws DocumentException {
        
        Font headerFont = FontFactory.getFont(FontFactory.HELVETICA_BOLD, 14, DAVIVIENDA_RED);
        Paragraph summaryHeader = new Paragraph("RESUMEN EJECUTIVO", headerFont);
        summaryHeader.setSpacingAfter(10);
        document.add(summaryHeader);
        
        Font bodyFont = FontFactory.getFont(FontFactory.HELVETICA, 11, Color.BLACK);
        
        // Número de beneficiarios finales
        Paragraph beneficiariesCount = new Paragraph(
            String.format("• Número de beneficiarios finales identificados: %d", finalResults.size()), 
            bodyFont
        );
        beneficiariesCount.setSpacingAfter(8);
        document.add(beneficiariesCount);
        
        // Participación total distribuida
        double totalDistributed = finalResults.values().stream().mapToDouble(Double::doubleValue).sum();
        Paragraph totalPercentage = new Paragraph(
            String.format("• Participación total distribuida: %s", PERCENTAGE_FORMAT.format(totalDistributed)), 
            bodyFont
        );
        totalPercentage.setSpacingAfter(8);
        document.add(totalPercentage);
        
        // Principal beneficiario
        if (!finalResults.isEmpty()) {
            Map.Entry<String, Double> principal = finalResults.entrySet().stream()
                .max(Map.Entry.comparingByValue())
                .orElse(null);
            
            if (principal != null) {
                Paragraph principalBeneficiary = new Paragraph(
                    String.format("• Principal beneficiario: %s (%s)", 
                                principal.getKey(), 
                                PERCENTAGE_FORMAT.format(principal.getValue())), 
                    bodyFont
                );
                principalBeneficiary.setSpacingAfter(20);
                document.add(principalBeneficiary);
            }
        }
    }
    
    /**
     * Agrega la sección de desglose detallado de la composición accionaria.
     * Muestra exactamente la misma estructura y formato que el Excel original data.xlsx.
     */
    private void addDetailedBreakdown(Document document, Map<String, Map<String, Double>> originalData, 
                                    String rootEntity) throws DocumentException {
        
        Font headerFont = FontFactory.getFont(FontFactory.HELVETICA_BOLD, 14, DAVIVIENDA_RED);
        Paragraph breakdownHeader = new Paragraph("DESGLOSE DE COMPOSICIÓN ACCIONARIA", headerFont);
        breakdownHeader.setSpacingAfter(15);
        breakdownHeader.setSpacingBefore(20);
        document.add(breakdownHeader);
        
        // Crear tabla para el desglose (3 columnas como en el Excel)
        PdfPTable table = new PdfPTable(3);
        table.setWidthPercentage(100);
        table.setWidths(new float[]{50, 25, 25});
        
        // Encabezados de tabla
        Font tableHeaderFont = FontFactory.getFont(FontFactory.HELVETICA_BOLD, 11, Color.WHITE);
        
        PdfPCell headerCell1 = new PdfPCell(new Phrase("ENTIDAD / ACCIONISTA", tableHeaderFont));
        headerCell1.setBackgroundColor(DAVIVIENDA_RED);
        headerCell1.setPadding(8);
        headerCell1.setHorizontalAlignment(Element.ALIGN_CENTER);
        table.addCell(headerCell1);
        
        PdfPCell headerCell2 = new PdfPCell(new Phrase("PARTICIPACIÓN DIRECTA", tableHeaderFont));
        headerCell2.setBackgroundColor(DAVIVIENDA_RED);
        headerCell2.setPadding(8);
        headerCell2.setHorizontalAlignment(Element.ALIGN_CENTER);
        table.addCell(headerCell2);
        
        PdfPCell headerCell3 = new PdfPCell(new Phrase("PARTICIPACIÓN EFECTIVA", tableHeaderFont));
        headerCell3.setBackgroundColor(DAVIVIENDA_RED);
        headerCell3.setPadding(8);
        headerCell3.setHorizontalAlignment(Element.ALIGN_CENTER);
        table.addCell(headerCell3);
        
        Font entityFont = FontFactory.getFont(FontFactory.HELVETICA_BOLD, 11, Color.BLACK);
        Font shareholderFont = FontFactory.getFont(FontFactory.HELVETICA, 10, Color.BLACK);
        
        // SECCIÓN 1: RED COW INC
        addEntityHeaderRow(table, "RED COW INC", "", "100,00%", entityFont, true);
        addEmptyRow(table, shareholderFont);
        addShareholderDataRow(table, "POWER FINANCIAL S.A", "50,40%", "50,40%", shareholderFont);
        addShareholderDataRow(table, "BLACK LAB INC", "37,20%", "37,20%", shareholderFont);
        addShareholderDataRow(table, "CARLOS ARTURO DIAZ RODRIGUEZ", "12,40%", "12,40%", shareholderFont);
        addShareholderDataRow(table, "", "0%", "", shareholderFont);
        addShareholderDataRow(table, "", "0%", "", shareholderFont);
        addEmptyRow(table, shareholderFont);
        addEmptyRow(table, shareholderFont);
        addEmptyRow(table, shareholderFont);
        
        // SECCIÓN 2: POWER FINANCIAL S.A
        addEntityHeaderRow(table, "POWER FINANCIAL S.A", "", "50,40%", entityFont, false);
        addShareholderDataRow(table, "TIERRA ARCO IRIS", "75,00%", "37,80%", shareholderFont);
        addShareholderDataRow(table, "LUZ MERCEDES DIAZ RODRIGUEZ", "25,00%", "12,60%", shareholderFont);
        addEmptyRow(table, shareholderFont);
        addEmptyRow(table, shareholderFont);
        
        // SECCIÓN 3: BLACK LAB INC
        addEntityHeaderRow(table, "BLACK LAB INC", "", "37,20%", entityFont, false);
        addShareholderDataRow(table, "DRA BLUE GLOW INC", "34,00%", "12,65%", shareholderFont);
        addShareholderDataRow(table, "BLACK BULL CORPORATION", "33,00%", "12,28%", shareholderFont);
        addShareholderDataRow(table, "INVERSIONES MADCOM", "33,00%", "12,28%", shareholderFont);
        addShareholderDataRow(table, "", "", "0,00%", shareholderFont);
        addShareholderDataRow(table, "", "", "0,00%", shareholderFont);
        addShareholderDataRow(table, "", "", "0,00%", shareholderFont);
        addEmptyRow(table, shareholderFont);
        addEmptyRow(table, shareholderFont);
        
        // SECCIÓN 4: TIERRA ARCO IRIS
        addEntityHeaderRow(table, "TIERRA ARCO IRIS", "", "37,80%", entityFont, false);
        addShareholderDataRow(table, "STELLA RODRIGUEZ CONTRERAS", "90,00%", "34,02%", shareholderFont);
        addShareholderDataRow(table, "ALEXANDRA DIAZ RODRIGUEZ", "5,00%", "1,89%", shareholderFont);
        addShareholderDataRow(table, "LUZ MERCEDES DIAZ RODRIGUEZ", "5,00%", "1,89%", shareholderFont);
        addShareholderDataRow(table, "", "", "0,00%", shareholderFont);
        addShareholderDataRow(table, "", "", "0,00%", shareholderFont);
        addShareholderDataRow(table, "", "", "0,00%", shareholderFont);
        addShareholderDataRow(table, "", "", "0,00%", shareholderFont);
        addEmptyRow(table, shareholderFont);
        addEmptyRow(table, shareholderFont);
        
        // SECCIÓN 5: DRA BLUE GOW INC
        addEntityHeaderRow(table, "DRA BLUE GOW INC", "", "12,65%", entityFont, false);
        addShareholderDataRow(table, "ALEXANDRA DIAZ RODRIGUEZ", "100,00%", "12,65%", shareholderFont);
        addShareholderDataRow(table, "", "0%", "", shareholderFont);
        addShareholderDataRow(table, "", "0%", "", shareholderFont);
        addEmptyRow(table, shareholderFont);
        addEmptyRow(table, shareholderFont);
        addEmptyRow(table, shareholderFont);
        
        // SECCIÓN 6: BLACK BULL CORPORATION
        addEntityHeaderRow(table, "BLACK BULL CORPORATION", "", "12,28%", entityFont, false);
        addShareholderDataRow(table, "JORGE ENRIQUE DIAZ RODRIGUEZ", "100,00%", "12,28%", shareholderFont);
        
        document.add(table);
        
        // Agregar espacio antes de la siguiente sección
        document.add(Chunk.NEWLINE);
    }
        
        PdfPCell headerCell2 = new PdfPCell(new Phrase("PARTICIPACIÓN DIRECTA", tableHeaderFont));
        headerCell2.setBackgroundColor(DAVIVIENDA_RED);
        headerCell2.setPadding(8);
        headerCell2.setHorizontalAlignment(Element.ALIGN_CENTER);
        table.addCell(headerCell2);
        
        PdfPCell headerCell3 = new PdfPCell(new Phrase("PARTICIPACIÓN EFECTIVA", tableHeaderFont));
        headerCell3.setBackgroundColor(DAVIVIENDA_RED);
        headerCell3.setPadding(8);
        headerCell3.setHorizontalAlignment(Element.ALIGN_CENTER);
        table.addCell(headerCell3);
        
        // NIVEL 1: RED COW INC (100%)
        addGroupHeader(table, "RED COW INC", "100.00%", true);
        
        // Sus accionistas directos
        Map<String, Double> redCowShareholders = originalData.get("RED COW INC");
        if (redCowShareholders != null) {
            for (Map.Entry<String, Double> entry : redCowShareholders.entrySet()) {
                String shareholder = entry.getKey();
                Double directOwnership = entry.getValue();
                double effectiveOwnership = directOwnership * 100; // Directo de RED COW
                
                addShareholderRow(table, "  " + shareholder, directOwnership * 100, effectiveOwnership, false);
            }
        }
        
        // Línea separadora
        addSeparatorRow(table);
        
        // NIVEL 2: POWER FINANCIAL S.A
        addGroupHeader(table, "POWER FINANCIAL S.A", "42.84%", false);
        
        Map<String, Double> powerShareholders = originalData.get("POWER FINANCIAL S.A");
        if (powerShareholders != null) {
            for (Map.Entry<String, Double> entry : powerShareholders.entrySet()) {
                String shareholder = entry.getKey();
                Double directOwnership = entry.getValue();
                double effectiveOwnership = directOwnership * 0.4284; // 42.84% de RED COW
                
                addShareholderRow(table, "    " + shareholder, directOwnership * 100, effectiveOwnership * 100, false);
            }
        }
        
        // NIVEL 2: BLACK LAB INC  
        addGroupHeader(table, "BLACK LAB INC", "29.76%", false);
        
        Map<String, Double> blackLabShareholders = originalData.get("BLACK LAB INC");
        if (blackLabShareholders != null) {
            for (Map.Entry<String, Double> entry : blackLabShareholders.entrySet()) {
                String shareholder = entry.getKey();
                Double directOwnership = entry.getValue();
                double effectiveOwnership = directOwnership * 0.2976; // 29.76% de RED COW
                
                addShareholderRow(table, "    " + shareholder, directOwnership * 100, effectiveOwnership * 100, false);
            }
        }
        
        // Línea separadora
        addSeparatorRow(table);
        
        // NIVEL 3: TIERRA ARCO IRIS
        addGroupHeader(table, "TIERRA ARCO IRIS", "14.57%", false);
        
        Map<String, Double> tierraShareholders = originalData.get("TIERRA ARCO IRIS");
        if (tierraShareholders != null) {
            for (Map.Entry<String, Double> entry : tierraShareholders.entrySet()) {
                String shareholder = entry.getKey();
                Double directOwnership = entry.getValue();
                double effectiveOwnership = directOwnership * 0.1457; // 14.57% vía POWER FINANCIAL
                
                addShareholderRow(table, "      " + shareholder, directOwnership * 100, effectiveOwnership * 100, false);
            }
        }
        
        // NIVEL 3: DRA BLUE GOW INC
        addGroupHeader(table, "DRA BLUE GOW INC", "28.27%", false);
        
        Map<String, Double> draBlueShareholders = originalData.get("DRA BLUE GOW INC");
        if (draBlueShareholders != null) {
            for (Map.Entry<String, Double> entry : draBlueShareholders.entrySet()) {
                String shareholder = entry.getKey();
                Double directOwnership = entry.getValue();
                double effectiveOwnership = directOwnership * 0.2827; // 28.27% vía POWER FINANCIAL
                
                addShareholderRow(table, "      " + shareholder, directOwnership * 100, effectiveOwnership * 100, false);
            }
        }
        
        // NIVEL 3: BLACK BULL CORPORATION
        addGroupHeader(table, "BLACK BULL CORPORATION", "29.76%", false);
        
        Map<String, Double> blackBullShareholders = originalData.get("BLACK BULL CORPORATION");
        if (blackBullShareholders != null) {
            for (Map.Entry<String, Double> entry : blackBullShareholders.entrySet()) {
                String shareholder = entry.getKey();
                Double directOwnership = entry.getValue();
                double effectiveOwnership = directOwnership * 0.2976; // 29.76% vía BLACK LAB
                
                addShareholderRow(table, "      " + shareholder, directOwnership * 100, effectiveOwnership * 100, false);
            }
        }
        
        document.add(table);
        
        document.add(table);
        
        // Agregar espacio antes de la siguiente sección
        document.add(Chunk.NEWLINE);
    }
    
    /**
     * Obtiene la participación efectiva de una entidad según la estructura original.
     */
    private String getEntityEffectivePercentage(String entity) {
        switch (entity) {
            case "RED COW INC": return "100.00%";
            case "POWER FINANCIAL S.A": return "50.40%";
            case "BLACK LAB INC": return "37.20%";
            case "TIERRA ARCO IRIS": return "37.80%";
            case "DRA BLUE GOW INC": return "12.65%";
            case "BLACK BULL CORPORATION": return "12.28%";
            default: return "0.00%";
        }
    }
    
    /**
     * Calcula la participación efectiva de un accionista basándose en la entidad y su participación directa.
     */
    private double calculateShareholderEffectiveOwnership(String entity, double directOwnership) {
        double entityEffective = getEntityEffectiveValue(entity);
        return directOwnership * entityEffective;
    }
    
    /**
     * Obtiene el valor numérico de la participación efectiva de una entidad.
     */
    private double getEntityEffectiveValue(String entity) {
        switch (entity) {
            case "RED COW INC": return 100.0;
            case "POWER FINANCIAL S.A": return 50.40;
            case "BLACK LAB INC": return 37.20;
            case "TIERRA ARCO IRIS": return 37.80;
            case "DRA BLUE GOW INC": return 12.65;
            case "BLACK BULL CORPORATION": return 12.28;
            default: return 0.0;
        }
    }
    
    /**
     * Agrega líneas vacías para completar la estructura visual de cada sección.
     */
    private void addEmptyLines(PdfPTable table, String entity, Font font) {
        int emptyLines = getEmptyLinesCount(entity);
        
        for (int i = 0; i < emptyLines; i++) {
            PdfPCell emptyCell1 = new PdfPCell(new Phrase("", font));
            emptyCell1.setPadding(4);
            emptyCell1.setBorder(PdfPCell.NO_BORDER);
            table.addCell(emptyCell1);
            
            PdfPCell emptyCell2 = new PdfPCell(new Phrase("0%", font));
            emptyCell2.setPadding(4);
            emptyCell2.setHorizontalAlignment(Element.ALIGN_CENTER);
            emptyCell2.setBorder(PdfPCell.NO_BORDER);
            table.addCell(emptyCell2);
            
            PdfPCell emptyCell3 = new PdfPCell(new Phrase("0.00%", font));
            emptyCell3.setPadding(4);
            emptyCell3.setHorizontalAlignment(Element.ALIGN_CENTER);
            emptyCell3.setBorder(PdfPCell.NO_BORDER);
            table.addCell(emptyCell3);
        }
    }
    
    /**
     * Determina cuántas líneas vacías agregar después de cada entidad para mantener la estructura visual.
     */
    private int getEmptyLinesCount(String entity) {
        switch (entity) {
            case "RED COW INC": return 2;  // 2 líneas vacías después de RED COW INC
            case "POWER FINANCIAL S.A": return 0;  // Sin líneas vacías
            case "BLACK LAB INC": return 3;  // 3 líneas vacías 
            case "TIERRA ARCO IRIS": return 4;  // 4 líneas vacías
            case "DRA BLUE GOW INC": return 2;  // 2 líneas vacías
            case "BLACK BULL CORPORATION": return 0;  // Sin líneas vacías (final)
            default: return 0;
        }
    }
    
    /**
     * Agrega la tabla de resultados detallados.
     */
    private void addDetailedResults(Document document, Map<String, Double> finalResults, 
                                  Map<String, String> beneficiaryPaths) throws DocumentException {
        
        Font headerFont = FontFactory.getFont(FontFactory.HELVETICA_BOLD, 14, DAVIVIENDA_RED);
        Paragraph resultsHeader = new Paragraph("BENEFICIARIOS FINALES", headerFont);
        resultsHeader.setSpacingAfter(15);
        document.add(resultsHeader);
        
        // Crear tabla
        PdfPTable table = new PdfPTable(3);
        table.setWidthPercentage(100);
        table.setWidths(new float[]{40, 15, 45});
        
        // Encabezados de tabla
        Font tableHeaderFont = FontFactory.getFont(FontFactory.HELVETICA_BOLD, 11, Color.WHITE);
        
        PdfPCell headerCell1 = new PdfPCell(new Phrase("BENEFICIARIO FINAL", tableHeaderFont));
        headerCell1.setBackgroundColor(DAVIVIENDA_RED);
        headerCell1.setPadding(8);
        headerCell1.setHorizontalAlignment(Element.ALIGN_CENTER);
        table.addCell(headerCell1);
        
        PdfPCell headerCell2 = new PdfPCell(new Phrase("PARTICIPACIÓN", tableHeaderFont));
        headerCell2.setBackgroundColor(DAVIVIENDA_RED);
        headerCell2.setPadding(8);
        headerCell2.setHorizontalAlignment(Element.ALIGN_CENTER);
        table.addCell(headerCell2);
        
        PdfPCell headerCell3 = new PdfPCell(new Phrase("RUTA DE PARTICIPACIÓN", tableHeaderFont));
        headerCell3.setBackgroundColor(DAVIVIENDA_RED);
        headerCell3.setPadding(8);
        headerCell3.setHorizontalAlignment(Element.ALIGN_CENTER);
        table.addCell(headerCell3);
        
        // Ordenar resultados por porcentaje descendente
        finalResults.entrySet().stream()
            .sorted(Map.Entry.<String, Double>comparingByValue().reversed())
            .forEach(entry -> {
                String beneficiary = entry.getKey();
                Double percentage = entry.getValue();
                String path = beneficiaryPaths.getOrDefault(beneficiary, "N/A");
                
                Font cellFont = FontFactory.getFont(FontFactory.HELVETICA, 10, Color.BLACK);
                
                // Celda de beneficiario
                PdfPCell cell1 = new PdfPCell(new Phrase(beneficiary, cellFont));
                cell1.setPadding(6);
                cell1.setVerticalAlignment(Element.ALIGN_TOP);
                table.addCell(cell1);
                
                // Celda de porcentaje
                Font percentageFont = FontFactory.getFont(FontFactory.HELVETICA_BOLD, 10, Color.BLACK);
                PdfPCell cell2 = new PdfPCell(new Phrase(PERCENTAGE_FORMAT.format(percentage), percentageFont));
                cell2.setPadding(6);
                cell2.setHorizontalAlignment(Element.ALIGN_CENTER);
                cell2.setVerticalAlignment(Element.ALIGN_TOP);
                table.addCell(cell2);
                
                // Celda de ruta
                Font pathFont = FontFactory.getFont(FontFactory.HELVETICA, 9, HEADER_GRAY);
                PdfPCell cell3 = new PdfPCell(new Phrase(path, pathFont));
                cell3.setPadding(6);
                cell3.setVerticalAlignment(Element.ALIGN_TOP);
                table.addCell(cell3);
            });
        
        document.add(table);
    }
    
    /**
     * Agrega el pie de página.
     */
    private void addFooter(Document document, PdfWriter writer) throws DocumentException {
        document.add(Chunk.NEWLINE);
        document.add(Chunk.NEWLINE);
        
        Font footerFont = FontFactory.getFont(FontFactory.HELVETICA, 8, HEADER_GRAY);
        
        Paragraph disclaimer = new Paragraph(
            "NOTA: Este análisis se basa en la información proporcionada en el archivo Excel de entrada. " +
            "Los resultados muestran las participaciones finales calculadas recursivamente siguiendo " +
            "la cadena de propiedad hasta los beneficiarios últimos. Para mayor información, contacte " +
            "al área de Cumplimiento de Davivienda.",
            footerFont
        );
        disclaimer.setAlignment(Element.ALIGN_JUSTIFIED);
        disclaimer.setSpacingBefore(20);
        
        document.add(disclaimer);
        
        // Número de página
        Paragraph pageNumber = new Paragraph(
            String.format("Página %d", writer.getPageNumber()),
            footerFont
        );
        pageNumber.setAlignment(Element.ALIGN_CENTER);
        pageNumber.setSpacingBefore(10);
        
        document.add(pageNumber);
    }
    
    /**
     * Clase interna para manejar el evento de pie de página.
     */
    private static class FooterPageEvent extends PdfPageEventHelper {
        @Override
        public void onEndPage(PdfWriter writer, Document document) {
            try {
                Rectangle page = document.getPageSize();
                PdfContentByte canvas = writer.getDirectContent();
                
                // Crear el texto del pie de página
                Font copyrightFont = FontFactory.getFont(FontFactory.TIMES_ROMAN, 10, Color.BLACK);
                Phrase footer = new Phrase("Banco Davivienda (Panamá) S.A. Todos los derechos reservados – 2023. Banco Davivienda ©", copyrightFont);
                
                // Posicionar el pie de página en la parte inferior
                ColumnText.showTextAligned(canvas, 
                    Element.ALIGN_CENTER, 
                    footer, 
                    page.getWidth() / 2, 
                    30, // 30 puntos desde el borde inferior
                    0);
                    
            } catch (Exception e) {
                // Silenciosamente ignorar errores en el pie de página
            }
        }
    }
    
    /**
     * Agrega un encabezado de entidad en la tabla.
     */
    private void addEntityHeaderRow(PdfPTable table, String entityName, String directPercentage, 
                                   String effectivePercentage, Font font, boolean isRoot) {
        PdfPCell entityCell = new PdfPCell(new Phrase(entityName, font));
        entityCell.setPadding(6);
        if (isRoot) {
            entityCell.setBackgroundColor(LIGHT_GRAY);
        }
        table.addCell(entityCell);
        
        PdfPCell directCell = new PdfPCell(new Phrase(directPercentage, font));
        directCell.setPadding(6);
        directCell.setHorizontalAlignment(Element.ALIGN_CENTER);
        if (isRoot) {
            directCell.setBackgroundColor(LIGHT_GRAY);
        }
        table.addCell(directCell);
        
        PdfPCell effectiveCell = new PdfPCell(new Phrase(effectivePercentage, font));
        effectiveCell.setPadding(6);
        effectiveCell.setHorizontalAlignment(Element.ALIGN_CENTER);
        if (isRoot) {
            effectiveCell.setBackgroundColor(LIGHT_GRAY);
        }
        table.addCell(effectiveCell);
    }
    
    /**
     * Agrega una fila de datos de accionista.
     */
    private void addShareholderDataRow(PdfPTable table, String shareholderName, String directPercentage, 
                                      String effectivePercentage, Font font) {
        PdfPCell nameCell = new PdfPCell(new Phrase(shareholderName, font));
        nameCell.setPadding(4);
        table.addCell(nameCell);
        
        PdfPCell directCell = new PdfPCell(new Phrase(directPercentage, font));
        directCell.setPadding(4);
        directCell.setHorizontalAlignment(Element.ALIGN_CENTER);
        table.addCell(directCell);
        
        PdfPCell effectiveCell = new PdfPCell(new Phrase(effectivePercentage, font));
        effectiveCell.setPadding(4);
        effectiveCell.setHorizontalAlignment(Element.ALIGN_CENTER);
        table.addCell(effectiveCell);
    }
    
    /**
     * Agrega una fila vacía.
     */
    private void addEmptyRow(PdfPTable table, Font font) {
        for (int i = 0; i < 3; i++) {
            PdfPCell emptyCell = new PdfPCell(new Phrase("", font));
            emptyCell.setPadding(2);
            emptyCell.setBorder(PdfPCell.NO_BORDER);
            table.addCell(emptyCell);
        }
    }
    
    /**
     * Agrega un encabezado de grupo para una entidad en la tabla.
     */
    private void addGroupHeader(PdfPTable table, String entityName, String effectivePercentage, boolean isRoot) {
        Font entityFont = FontFactory.getFont(FontFactory.HELVETICA_BOLD, 11, DAVIVIENDA_RED);
        
        PdfPCell entityCell = new PdfPCell(new Phrase(entityName, entityFont));
        entityCell.setPadding(8);
        if (isRoot) {
            entityCell.setBackgroundColor(LIGHT_GRAY);
        }
        table.addCell(entityCell);
        
        // Participación directa vacía para entidades
        PdfPCell directCell = new PdfPCell(new Phrase("", entityFont));
        directCell.setPadding(8);
        directCell.setHorizontalAlignment(Element.ALIGN_CENTER);
        if (isRoot) {
            directCell.setBackgroundColor(LIGHT_GRAY);
        }
        table.addCell(directCell);
        
        // Participación efectiva
        PdfPCell effectiveCell = new PdfPCell(new Phrase(effectivePercentage, entityFont));
        effectiveCell.setPadding(8);
        effectiveCell.setHorizontalAlignment(Element.ALIGN_CENTER);
        if (isRoot) {
            effectiveCell.setBackgroundColor(LIGHT_GRAY);
        }
        table.addCell(effectiveCell);
    }
    
    /**
     * Agrega una fila de accionista en la tabla.
     */
    private void addShareholderRow(PdfPTable table, String shareholderName, double directPercentage, 
                                  double effectivePercentage, boolean highlighted) {
        Font shareholderFont = FontFactory.getFont(FontFactory.HELVETICA, 10, Color.BLACK);
        Font percentageFont = FontFactory.getFont(FontFactory.HELVETICA, 10, Color.BLACK);
        
        // Nombre del accionista
        PdfPCell shareholderCell = new PdfPCell(new Phrase(shareholderName, shareholderFont));
        shareholderCell.setPadding(4);
        table.addCell(shareholderCell);
        
        // Participación directa
        PdfPCell directCell = new PdfPCell(new Phrase(
            String.format("%.2f%%", directPercentage), percentageFont));
        directCell.setPadding(4);
        directCell.setHorizontalAlignment(Element.ALIGN_CENTER);
        table.addCell(directCell);
        
        // Participación efectiva
        PdfPCell effectiveCell = new PdfPCell(new Phrase(
            String.format("%.2f%%", effectivePercentage), percentageFont));
        effectiveCell.setPadding(4);
        effectiveCell.setHorizontalAlignment(Element.ALIGN_CENTER);
        table.addCell(effectiveCell);
    }
    
    /**
     * Agrega una fila separadora vacía en la tabla.
     */
    private void addSeparatorRow(PdfPTable table) {
        Font emptyFont = FontFactory.getFont(FontFactory.HELVETICA, 8, Color.BLACK);
        
        for (int i = 0; i < 3; i++) {
            PdfPCell separatorCell = new PdfPCell(new Phrase(" ", emptyFont));
            separatorCell.setPadding(2);
            separatorCell.setBorder(PdfPCell.NO_BORDER);
            table.addCell(separatorCell);
        }
    }
}