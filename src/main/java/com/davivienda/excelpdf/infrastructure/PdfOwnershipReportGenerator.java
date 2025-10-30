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
     * @throws IOException si hay problemas escribiendo el archivo
     */
    public void generateOwnershipReport(Map<String, Double> finalResults,
                                      Map<String, String> beneficiaryPaths,
                                      String rootEntity,
                                      String outputPath,
                                      Map<String, Map<String, Double>> originalData,
                                      java.util.List<String[]> dataXlsxRows) throws IOException {
        
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
            addDetailedBreakdown(document, originalData, rootEntity, dataXlsxRows);
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
     * Añade la tabla detallada de desglose de composición accionaria replicando exactamente las filas 4-45 de data.xlsx
     */
    private void addDetailedBreakdown(Document document, Map<String, Map<String, Double>> originalData, String rootEntity, java.util.List<String[]> dataXlsxRows) 
            throws DocumentException {
        
        logger.info("Añadiendo tabla de desglose de composición accionaria...");
        
        // Espacio antes de la tabla
        document.add(new Paragraph(" "));
        document.add(new Paragraph(" "));
        
        // Título de la tabla
        Font titleFont = FontFactory.getFont(FontFactory.HELVETICA_BOLD, 12, Color.BLACK);
        Paragraph title = new Paragraph("DESGLOSE DE COMPOSICIÓN ACCIONARIA", titleFont);
        title.setAlignment(Element.ALIGN_CENTER);
        document.add(title);
        document.add(new Paragraph(" "));
        
        // Crear tabla con 3 columnas (A, B, C) como en Excel
        PdfPTable table = new PdfPTable(3);
        table.setWidthPercentage(100);
        
        // Configurar anchos de columnas: 50% para entidad, 25% para porcentaje directo, 25% para porcentaje final
        float[] columnWidths = {50f, 25f, 25f};
        table.setWidths(columnWidths);
        
        // Fuentes para la tabla
        Font headerFont = FontFactory.getFont(FontFactory.HELVETICA_BOLD, 9, Color.BLACK);
        Font contentFont = FontFactory.getFont(FontFactory.HELVETICA, 8, Color.BLACK);
        Font boldFont = FontFactory.getFont(FontFactory.HELVETICA_BOLD, 8, Color.BLACK);
        
        // Encabezados
        PdfPCell headerCell1 = new PdfPCell(new Phrase("ENTIDAD", headerFont));
        headerCell1.setBackgroundColor(new Color(220, 220, 220));
        headerCell1.setPadding(6);
        headerCell1.setHorizontalAlignment(Element.ALIGN_CENTER);
        table.addCell(headerCell1);
        
        PdfPCell headerCell2 = new PdfPCell(new Phrase("% DIRECTO", headerFont));
        headerCell2.setBackgroundColor(new Color(220, 220, 220));
        headerCell2.setPadding(6);
        headerCell2.setHorizontalAlignment(Element.ALIGN_CENTER);
        table.addCell(headerCell2);
        
        PdfPCell headerCell3 = new PdfPCell(new Phrase("% FINAL", headerFont));
        headerCell3.setBackgroundColor(new Color(220, 220, 220));
        headerCell3.setPadding(6);
        headerCell3.setHorizontalAlignment(Element.ALIGN_CENTER);
        table.addCell(headerCell3);
        
        // Añadir todas las filas de data.xlsx (4-45) exactamente como están
        for (String[] rowData : dataXlsxRows) {
            String entityName = rowData[0];
            String directPercentage = rowData[1];
            String finalPercentage = rowData[2];
            
            // Columna A - Entidad
            Font cellFont = (!entityName.isEmpty() && !entityName.equals("0")) ? boldFont : contentFont;
            PdfPCell entityCell = new PdfPCell(new Phrase(entityName, cellFont));
            entityCell.setPadding(4);
            entityCell.setHorizontalAlignment(Element.ALIGN_LEFT);
            table.addCell(entityCell);
            
            // Columna B - Porcentaje directo
            PdfPCell directCell = new PdfPCell(new Phrase(directPercentage, contentFont));
            directCell.setPadding(4);
            directCell.setHorizontalAlignment(Element.ALIGN_CENTER);
            table.addCell(directCell);
            
            // Columna C - Porcentaje final
            PdfPCell finalCell = new PdfPCell(new Phrase(finalPercentage, contentFont));
            finalCell.setPadding(4);
            finalCell.setHorizontalAlignment(Element.ALIGN_CENTER);
            table.addCell(finalCell);
        }
        
        // Añadir la tabla al documento
        document.add(table);
        logger.info("Tabla de desglose añadida exitosamente");
    }
}