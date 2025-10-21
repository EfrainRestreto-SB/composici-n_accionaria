package com.davivienda.excelpdf.infrastructure;

import com.lowagie.text.*;
import com.lowagie.text.pdf.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.Color;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Map;

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
     * @throws IOException si hay problemas escribiendo el archivo
     */
    public void generateOwnershipReport(Map<String, Double> finalResults,
                                      Map<String, String> beneficiaryPaths,
                                      String rootEntity,
                                      String outputPath) throws IOException {
        
        logger.info("Generando reporte PDF: {}", outputPath);
        
        try {
            Document document = new Document(PageSize.A4, 50, 50, 50, 50);
            PdfWriter writer = PdfWriter.getInstance(document, new FileOutputStream(outputPath));
            
            // Configurar metadatos
            document.addTitle("Análisis de Composición Accionaria - " + rootEntity);
            document.addAuthor("Davivienda");
            document.addSubject("Composición Accionaria");
            document.addCreator("Sistema de Análisis Davivienda v1.0");
            
            document.open();
            
            // Agregar contenido al documento
            addHeader(document, rootEntity);
            addSummary(document, finalResults, rootEntity);
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
     * Agrega el encabezado del documento.
     */
    private void addHeader(Document document, String rootEntity) throws DocumentException {
        // Título principal
        Font titleFont = FontFactory.getFont(FontFactory.HELVETICA_BOLD, 18, DAVIVIENDA_RED);
        Paragraph title = new Paragraph("ANÁLISIS DE COMPOSICIÓN ACCIONARIA", titleFont);
        title.setAlignment(Element.ALIGN_CENTER);
        title.setSpacingAfter(10);
        document.add(title);
        
        // Subtítulo con entidad analizada
        Font subtitleFont = FontFactory.getFont(FontFactory.HELVETICA, 14, Color.BLACK);
        Paragraph subtitle = new Paragraph("Entidad Analizada: " + rootEntity, subtitleFont);
        subtitle.setAlignment(Element.ALIGN_CENTER);
        subtitle.setSpacingAfter(20);
        document.add(subtitle);
        
        // Fecha de generación
        Font dateFont = FontFactory.getFont(FontFactory.HELVETICA, 10, HEADER_GRAY);
        Paragraph date = new Paragraph("Fecha de Generación: " + LocalDateTime.now().format(DATE_FORMAT), dateFont);
        date.setAlignment(Element.ALIGN_RIGHT);
        date.setSpacingAfter(30);
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
     * Agrega la tabla de resultados detallados.
     */
    private void addDetailedResults(Document document, Map<String, Double> finalResults, 
                                  Map<String, String> beneficiaryPaths) throws DocumentException {
        
        Font headerFont = FontFactory.getFont(FontFactory.HELVETICA_BOLD, 14, DAVIVIENDA_RED);
        Paragraph resultsHeader = new Paragraph("RESULTADOS DETALLADOS", headerFont);
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
}