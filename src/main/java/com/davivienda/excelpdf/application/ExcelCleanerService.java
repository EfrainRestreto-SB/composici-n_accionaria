package com.davivienda.excelpdf.application;package com.davivienda.excelpdf.application;package com.davivienda.excelpdf.application;



import org.apache.poi.ss.usermodel.*;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.slf4j.Logger;import org.apache.poi.ss.usermodel.*;import java.io.FileInputStream;

import org.slf4j.LoggerFactory;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;import java.io.FileOutputStream;

import java.io.FileInputStream;

import java.io.FileOutputStream;import org.slf4j.Logger;import java.io.IOException;

import java.io.IOException;

import java.util.ArrayList;import org.slf4j.LoggerFactory;import java.util.ArrayList;

import java.util.List;

import java.util.stream.Collectors;import java.util.List;



/**import java.io.FileInputStream;

 * Servicio para limpiar y reestructurar archivos Excel con formato jerárquico.

 * Incluye lógica especializada para unificar entidades DRA BLUE con nombres similares.import java.io.FileOutputStream;import org.apache.poi.ss.usermodel.Cell;

 */

public class ExcelCleanerService {import java.io.IOException;import org.apache.poi.ss.usermodel.CellType;

    

    private static final Logger logger = LoggerFactory.getLogger(ExcelCleanerService.class);import java.util.ArrayList;import org.apache.poi.ss.usermodel.Row;

    

    /**import java.util.List;import org.apache.poi.ss.usermodel.Sheet;

     * Limpia y reestructura un archivo Excel con formato jerárquico,

     * aplicando correcciones automáticas para entidades conocidas.import java.util.stream.Collectors;import org.apache.poi.ss.usermodel.Workbook;

     */

    public String cleanAndRestructureExcel(String originalFilePath) throws IOException {import org.apache.poi.xssf.usermodel.XSSFWorkbook;

        logger.info("Iniciando limpieza del Excel con formato jerárquico: {}", originalFilePath);

        /**import org.slf4j.Logger;

        List<OwnershipRow> validRows = new ArrayList<>();

         * Servicio para limpiar y reestructurar archivos Excel con formato jerárquicoimport org.slf4j.LoggerFactory;

        try (FileInputStream fis = new FileInputStream(originalFilePath);

             XSSFWorkbook workbook = new XSSFWorkbook(fis)) { */

            

            Sheet sheet = workbook.getSheetAt(0);public class ExcelCleanerService {/**

            int rowCount = sheet.getPhysicalNumberOfRows();

            logger.info("Leyendo {} filas del Excel original.", rowCount);     * Servicio para limpiar y reestructurar archivos Excel con composición accionaria.

            

            String currentEntity = null;    private static final Logger logger = LoggerFactory.getLogger(ExcelCleanerService.class); * Elimina filas vacías, datos inválidos y reorganiza el formato para el procesamiento.

            

            for (int i = 0; i < rowCount; i++) {     * 

                Row row = sheet.getRow(i);

                if (row == null) continue;    /** * @author Davivienda

                

                Cell cellA = row.getCell(0);     * Limpia y reestructura un archivo Excel con formato jerárquico * @version 1.0

                Cell cellB = row.getCell(1);

                     */ */

                // Saltar filas completamente vacías

                if (cellA == null || isEmptyCell(cellA)) {    public String cleanAndRestructureExcel(String originalFilePath) throws IOException {public class ExcelCleanerService {

                    continue;

                }        logger.info("Iniciando limpieza del Excel con formato jerárquico: {}", originalFilePath);    

                

                String entityName = getCellValueAsString(cellA);            private static final Logger logger = LoggerFactory.getLogger(ExcelCleanerService.class);

                if (entityName == null || entityName.trim().isEmpty()) {

                    continue;        List<OwnershipRow> validRows = new ArrayList<>();    

                }

                            /**

                // Filtrar filas que contienen texto descriptivo

                if (entityName.toUpperCase().contains("COMPOSICION") ||        try (FileInputStream fis = new FileInputStream(originalFilePath);     * Limpia y reestructura un archivo Excel de composición accionaria.

                    entityName.toUpperCase().contains("ACCIONARIA") ||

                    entityName.toUpperCase().contains("TOTAL")) {             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {     * Soporta formato jerárquico en cascada donde:

                    continue;

                }                 * - Columna A: Entidad

                

                // Lógica mejorada para detectar entidades raíz vs accionistas            Sheet sheet = workbook.getSheetAt(0);     * - Columna B: Porcentaje de participación del accionista

                double percentageInB = getCellValueAsDouble(cellB);

                            int rowCount = sheet.getPhysicalNumberOfRows();     * - Columna C: Valor acumulado

                // Una entidad raíz se identifica por:

                // 1. B está vacío/nulo O            logger.info("Leyendo {} filas del Excel original.", rowCount);     * 

                // 2. B = 0 O  

                // 3. B = 1 Y no hay entidad actual (primera fila)                 * @param inputPath ruta del archivo Excel a limpiar

                boolean isEmptyOrZero = (cellB == null || isEmptyCell(cellB) || percentageInB == 0.0);

                boolean isFirstEntity = (percentageInB == 1.0 && currentEntity == null && i < 10);            String currentEntity = null;     * @param outputPath ruta del archivo Excel limpio de salida

                

                if (isEmptyOrZero || isFirstEntity) {                 * @return número de filas válidas procesadas

                    // Esta es la entidad raíz del grupo actual

                    currentEntity = entityName;            for (int i = 0; i < rowCount; i++) {     * @throws IOException si hay problemas de acceso a archivos

                    logger.debug("Entidad raíz detectada: {} (B={}, fila={})", currentEntity, percentageInB, i + 1);

                    continue;                Row row = sheet.getRow(i);     */

                }

                                if (row == null) continue;    public int cleanAndRestructureExcel(String inputPath, String outputPath) throws IOException {

                // Si llegamos aquí, esta fila es un accionista de la entidad actual

                if (currentEntity != null) {                        logger.info("Iniciando limpieza del Excel con formato jerárquico: {}", inputPath);

                    String shareholder = entityName;

                    double percentage = percentageInB * 100; // Convertir a porcentaje                Cell cellA = row.getCell(0);        

                    

                    // Validar que el porcentaje esté en rango válido                Cell cellB = row.getCell(1);        List<OwnershipRow> validRows = new ArrayList<>();

                    if (percentage > 0 && percentage <= 100) {

                        OwnershipRow ownershipRow = new OwnershipRow(currentEntity, shareholder, percentage);                        

                        validRows.add(ownershipRow);

                        logger.debug("Fila {} válida: {} -> {} ({}%)",                 // Saltar filas completamente vacías        // Leer el archivo Excel original

                                   i + 1, currentEntity, shareholder, percentage);

                    }                if (cellA == null || isEmptyCell(cellA)) {        try (FileInputStream fis = new FileInputStream(inputPath);

                }

            }                    continue;             Workbook workbook = new XSSFWorkbook(fis)) {

        }

                        }            

        logger.info("Se encontraron {} filas válidas antes de correcciones", validRows.size());

                                    Sheet sheet = workbook.getSheetAt(0);

        // Aplicar correcciones automáticas para casos conocidos

        validRows = applyAutomaticCorrections(validRows);                String entityName = getCellValueAsString(cellA);            int totalRows = sheet.getLastRowNum() + 1;

        

        logger.info("Se encontraron {} filas válidas después de correcciones", validRows.size());                if (entityName == null || entityName.trim().isEmpty()) {            

        

        // Mostrar entidades únicas encontradas                    continue;            logger.info("Leyendo {} filas del Excel original...", totalRows);

        List<String> uniqueEntities = validRows.stream()

            .map(row -> row.entity)                }            

            .distinct()

            .collect(Collectors.toList());                            // Variables para rastrear la jerarquía

        logger.info("Entidades únicas encontradas: {}", uniqueEntities);

                        // Filtrar filas que contienen texto descriptivo            String currentEntity = null;

        // Generar archivo Excel limpio

        String cleanFilePath = originalFilePath.replace(".xlsx", "_cleaned_fixed.xlsx");                if (entityName.toUpperCase().contains("COMPOSICION") ||            

        generateCleanExcel(validRows, cleanFilePath);

                            entityName.toUpperCase().contains("ACCIONARIA") ||            // Procesar cada fila

        logger.info("Excel limpio generado exitosamente: {}", cleanFilePath);

        return cleanFilePath;                    entityName.toUpperCase().contains("TOTAL")) {            for (int i = 0; i <= sheet.getLastRowNum(); i++) {

    }

                        continue;                Row row = sheet.getRow(i);

    /**

     * Aplica correcciones automáticas para casos conocidos en los datos                }                

     */

    private List<OwnershipRow> applyAutomaticCorrections(List<OwnershipRow> originalRows) {                                if (row == null) {

        List<OwnershipRow> correctedRows = new ArrayList<>(originalRows);

                        // Lógica mejorada para detectar entidades raíz vs accionistas                    continue;

        // Corrección 1: Unificar entidades DRA BLUE con nombres similares

        correctedRows = unifyDraBluEntities(correctedRows);                double percentageInB = getCellValueAsDouble(cellB);                }

        

        // Aquí se pueden agregar más correcciones automáticas en el futuro                                

        // Corrección 2: Otras entidades con nombres similares

        // Corrección 3: Validaciones de porcentajes                // Una entidad raíz se identifica por:                Cell cellA = row.getCell(0);

        

        return correctedRows;                // 1. B está vacío/nulo O                Cell cellB = row.getCell(1);

    }

                    // 2. B = 0 O                  

    /**

     * Unifica entidades DRA BLUE GLOW INC y DRA BLUE GOW INC que representan la misma entidad                // 3. B = 1 Y no hay entidad actual (primera fila)                // Saltar filas vacías o de encabezado

     */

    private List<OwnershipRow> unifyDraBluEntities(List<OwnershipRow> originalRows) {                boolean isEmptyOrZero = (cellB == null || isEmptyCell(cellB) || percentageInB == 0.0);                if (cellA == null || isEmptyCell(cellA)) {

        // Buscar entidades DRA BLUE con nombres similares

        boolean hasGlow = originalRows.stream().anyMatch(row ->                 boolean isFirstEntity = (percentageInB == 1.0 && currentEntity == null && i < 10);                    continue;

            "DRA BLUE GLOW INC".equals(row.shareholder));

        boolean hasGow = originalRows.stream().anyMatch(row ->                                 }

            "DRA BLUE GOW INC".equals(row.entity));

                            if (isEmptyOrZero || isFirstEntity) {                

        if (!hasGlow || !hasGow) {

            logger.debug("No se encontraron entidades DRA BLUE para unificar");                    // Esta es la entidad raíz del grupo actual                String entityName = getCellValueAsString(cellA);

            return originalRows;

        }                    currentEntity = entityName;                

        

        logger.info("Detectadas entidades DRA BLUE similares - aplicando unificación automática...");                    logger.debug("Entidad raíz detectada: {} (B={}, fila={})", currentEntity, percentageInB, i + 1);                // Saltar encabezados y valores inválidos

        

        List<OwnershipRow> unifiedRows = new ArrayList<>(originalRows);                    continue;                if (entityName.isEmpty() || 

        

        // Encontrar la relación padre -> DRA BLUE GLOW INC                }                    entityName.equalsIgnoreCase("X") ||

        OwnershipRow glowRelation = originalRows.stream()

            .filter(row -> "DRA BLUE GLOW INC".equals(row.shareholder))                                    entityName.contains("DESGLOSE") ||

            .findFirst()

            .orElse(null);                // Si llegamos aquí, esta fila es un accionista de la entidad actual                    entityName.contains("COMPOSICION")) {

            

        // Encontrar las relaciones DRA BLUE GOW INC -> hijos                  if (currentEntity != null) {                    continue;

        List<OwnershipRow> gowChildren = originalRows.stream()

            .filter(row -> "DRA BLUE GOW INC".equals(row.entity))                    String shareholder = entityName;                }

            .collect(Collectors.toList());

                                double percentage = percentageInB * 100; // Convertir a porcentaje                

        if (glowRelation != null && !gowChildren.isEmpty()) {

            logger.info("Aplicando corrección: {} será el padre de los hijos de DRA BLUE GOW INC",                                     // Lógica mejorada para detectar entidades raíz vs accionistas

                       glowRelation.shareholder);

                                // Validar que el porcentaje esté en rango válido                double percentageInB = getCellValueAsDouble(cellB);

            // Crear nuevas relaciones: DRA BLUE GLOW INC -> hijos de DRA BLUE GOW INC

            for (OwnershipRow gowChild : gowChildren) {                    if (percentage > 0 && percentage <= 100) {                

                OwnershipRow newRelation = new OwnershipRow(

                    "DRA BLUE GLOW INC",                         OwnershipRow ownershipRow = new OwnershipRow(currentEntity, shareholder, percentage);                // Una entidad raíz se identifica SOLO por:

                    gowChild.shareholder, 

                    gowChild.percentage                        validRows.add(ownershipRow);                // 1. B está vacío/nulo O

                );

                unifiedRows.add(newRelation);                        logger.debug("Fila {} válida: {} -> {} ({}%)",                 // 2. B = 0 O  

                logger.info("✓ Relación corregida: {} -> {} ({}%)", 

                           newRelation.entity, newRelation.shareholder, newRelation.percentage);                                   i + 1, currentEntity, shareholder, percentage);                // 3. B = 1 Y no hay entidad actual (es la primera fila del Excel)

            }

                                }                boolean isEmptyOrZero = (cellB == null || isEmptyCell(cellB) || percentageInB == 0.0);

            // Remover las relaciones originales de DRA BLUE GOW INC

            unifiedRows.removeIf(row -> "DRA BLUE GOW INC".equals(row.entity));                }                boolean isFirstEntity = (percentageInB == 1.0 && currentEntity == null && i < 10); // Solo las primeras filas

            logger.info("✓ Removidas {} relaciones originales de DRA BLUE GOW INC", gowChildren.size());

        }            }                

        

        return unifiedRows;        }                if (isEmptyOrZero || isFirstEntity) {

    }

                                // Esta es la entidad raíz del grupo actual

    /**

     * Genera un archivo Excel limpio con el formato estándar de 3 columnas        // Unificar entidades con nombres similares (DRA BLUE GLOW vs DRA BLUE GOW)                    currentEntity = entityName;

     */

    private void generateCleanExcel(List<OwnershipRow> validRows, String outputPath) throws IOException {        logger.info("Unificando entidades con nombres similares...");                    logger.debug("Entidad raíz detectada: {} (B={}, fila={})", currentEntity, percentageInB, i + 1);

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {

            Sheet sheet = workbook.createSheet("ComposicionAccionaria");        validRows = unifyDraBluEntities(validRows);                    continue;

            

            // Crear encabezados                        }

            Row headerRow = sheet.createRow(0);

            headerRow.createCell(0).setCellValue("Entidad");        logger.info("Se encontraron {} filas válidas después de unificación", validRows.size());                

            headerRow.createCell(1).setCellValue("Accionista");

            headerRow.createCell(2).setCellValue("Participación");                        // Si llegamos aquí, esta fila es un accionista de la entidad actual

            

            // Aplicar estilo a los encabezados        // Mostrar entidades únicas encontradas                if (currentEntity != null) {

            CellStyle headerStyle = workbook.createCellStyle();

            Font headerFont = workbook.createFont();        List<String> uniqueEntities = validRows.stream()                    // La celda A contiene el nombre del accionista

            headerFont.setBold(true);

            headerStyle.setFont(headerFont);            .map(row -> row.entity)                    // La celda B contiene el porcentaje de participación

            

            for (int i = 0; i < 3; i++) {            .distinct()                    String shareholder = entityName;

                headerRow.getCell(i).setCellStyle(headerStyle);

            }            .collect(Collectors.toList());                    double percentage = percentageInB * 100; // Convertir a porcentaje

            

            // Escribir datos        logger.info("Entidades únicas encontradas en columna A: {}", uniqueEntities);                    

            int rowNum = 1;

            for (OwnershipRow ownershipRow : validRows) {                            // DEBUG: Log detallado para entender el procesamiento

                Row row = sheet.createRow(rowNum++);

                row.createCell(0).setCellValue(ownershipRow.entity);        // Generar archivo Excel limpio                    logger.info("DEBUG Fila {}: Entidad={} | Accionista={} | B={} | Porcentaje={}%", 

                row.createCell(1).setCellValue(ownershipRow.shareholder);

                row.createCell(2).setCellValue(ownershipRow.percentage);        String cleanFilePath = originalFilePath.replace(".xlsx", "_cleaned.xlsx");                               i + 1, currentEntity, shareholder, percentageInB, percentage);

            }

                    generateCleanExcel(validRows, cleanFilePath);                    

            // Autoajustar columnas

            sheet.autoSizeColumn(0);                            // Validar que el porcentaje esté en rango válido

            sheet.autoSizeColumn(1);

            sheet.autoSizeColumn(2);        logger.info("Excel limpio generado exitosamente: {}", cleanFilePath);                    if (percentage > 0 && percentage <= 100) {

            

            // Guardar archivo        return cleanFilePath;                        OwnershipRow ownershipRow = new OwnershipRow(currentEntity, shareholder, percentage);

            try (FileOutputStream fos = new FileOutputStream(outputPath)) {

                workbook.write(fos);    }                        validRows.add(ownershipRow);

            }

        }                            logger.debug("Fila {} válida: {} -> {} ({}%)", 

    }

        /**                                   i + 1, currentEntity, shareholder, percentage);

    /**

     * Verifica si una celda está vacía     * Unifica entidades con nombres similares, específicamente DRA BLUE GLOW INC y DRA BLUE GOW INC                    } else {

     */

    private boolean isEmptyCell(Cell cell) {     */                        logger.warn("DEBUG Fila {} RECHAZADA - porcentaje fuera de rango: {}%", i + 1, percentage);

        if (cell == null) {

            return true;    private List<OwnershipRow> unifyDraBluEntities(List<OwnershipRow> originalRows) {                    }

        }

                List<OwnershipRow> unifiedRows = new ArrayList<>(originalRows);                } else {

        CellType cellType = cell.getCellType();

        if (cellType == CellType.FORMULA) {                            logger.warn("DEBUG Fila {} RECHAZADA - no hay entidad actual para: {}", i + 1, entityName);

            cellType = cell.getCachedFormulaResultType();

        }        // Buscar entidades DRA BLUE con nombres similares                }

        

        return switch (cellType) {        boolean hasGlow = originalRows.stream().anyMatch(row ->             }

            case BLANK -> true;

            case STRING -> cell.getStringCellValue().trim().isEmpty();            "DRA BLUE GLOW INC".equals(row.shareholder));        }

            case NUMERIC -> cell.getNumericCellValue() == 0.0;

            default -> false;        boolean hasGow = originalRows.stream().anyMatch(row ->         

        };

    }            "DRA BLUE GOW INC".equals(row.entity));        // Paso 2: Unificar entidades con nombres similares (DRA BLUE GLOW vs DRA BLUE GOW)

    

    /**                    logger.info("Unificando entidades con nombres similares...");

     * Obtiene el valor de una celda como String

     */        if (hasGlow && hasGow) {        validRows = unifyDraBluEntities(validRows);

    private String getCellValueAsString(Cell cell) {

        if (cell == null) {            logger.info("Detectadas entidades DRA BLUE similares - unificando...");        

            return null;

        }                    logger.info("Se encontraron {} filas válidas después de unificación", validRows.size());

        

        CellType cellType = cell.getCellType();            // Encontrar la relación padre -> DRA BLUE GLOW INC        

        if (cellType == CellType.FORMULA) {

            cellType = cell.getCachedFormulaResultType();            OwnershipRow glowRelation = originalRows.stream()        if (validRows.isEmpty()) {

        }

                        .filter(row -> "DRA BLUE GLOW INC".equals(row.shareholder))            throw new IllegalArgumentException("No se encontraron filas válidas en el Excel");

        return switch (cellType) {

            case STRING -> cell.getStringCellValue().trim();                .findFirst()        }

            case NUMERIC -> String.valueOf(cell.getNumericCellValue());

            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());                .orElse(null);        

            case BLANK -> "";

            default -> "";                        // Listar todas las entidades únicas encontradas

        };

    }            // Encontrar las relaciones DRA BLUE GOW INC -> hijos          java.util.Set<String> uniqueEntities = new java.util.HashSet<>();

    

    /**            List<OwnershipRow> gowChildren = originalRows.stream()        for (OwnershipRow row : validRows) {

     * Obtiene el valor de una celda como double

     */                .filter(row -> "DRA BLUE GOW INC".equals(row.entity))            uniqueEntities.add(row.entity);

    private double getCellValueAsDouble(Cell cell) {

        if (cell == null) {                .collect(Collectors.toList());        }

            return 0.0;

        }                        logger.info("Entidades únicas encontradas en columna A: {}", uniqueEntities);

        

        CellType cellType = cell.getCellType();            if (glowRelation != null && !gowChildren.isEmpty()) {        

        if (cellType == CellType.FORMULA) {

            cellType = cell.getCachedFormulaResultType();                // Crear nuevas relaciones: DRA BLUE GLOW INC -> hijos de DRA BLUE GOW INC        // Crear nuevo Excel limpio

        }

                        for (OwnershipRow gowChild : gowChildren) {        createCleanExcel(validRows, outputPath);

        return switch (cellType) {

            case NUMERIC -> cell.getNumericCellValue();                    OwnershipRow newRelation = new OwnershipRow(        

            case STRING -> {

                try {                        "DRA BLUE GLOW INC",         logger.info("Excel limpio generado exitosamente: {}", outputPath);

                    yield Double.parseDouble(cell.getStringCellValue().trim());

                } catch (NumberFormatException e) {                        gowChild.shareholder,         return validRows.size();

                    yield 0.0;

                }                        gowChild.percentage    }

            }

            default -> 0.0;                    );    

        };

    }                    unifiedRows.add(newRelation);    /**

    

    /**                    logger.info("Relación unificada: {} -> {} ({}%)",      * Parsea una fila del Excel y extrae los datos si son válidos.

     * Clase para representar una fila de participación accionaria

     */                               newRelation.entity, newRelation.shareholder, newRelation.percentage);     */

    public static class OwnershipRow {

        public final String entity;                }    private OwnershipRow parseRow(Row row, int rowNumber) {

        public final String shareholder;

        public final double percentage;                        try {

        

        public OwnershipRow(String entity, String shareholder, double percentage) {                // Remover las relaciones originales de DRA BLUE GOW INC            String entity = getCellValue(row.getCell(0));

            this.entity = entity;

            this.shareholder = shareholder;                unifiedRows.removeIf(row -> "DRA BLUE GOW INC".equals(row.entity));            String shareholder = getCellValue(row.getCell(1));

            this.percentage = percentage;

        }                logger.info("Removidas relaciones originales de DRA BLUE GOW INC");            String percentageStr = getCellValue(row.getCell(2));

        

        @Override            }            

        public String toString() {

            return String.format("%s -> %s (%.2f%%)", entity, shareholder, percentage);        }            // Validar que no estén vacíos

        }

    }                    if (entity.isEmpty() || shareholder.isEmpty() || percentageStr.isEmpty()) {

}
        return unifiedRows;                return null;

    }            }

                

    private void generateCleanExcel(List<OwnershipRow> validRows, String outputPath) throws IOException {            // Filtrar filas que son encabezados o texto descriptivo

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {            if (isHeaderOrDescriptiveText(entity, shareholder, percentageStr)) {

            Sheet sheet = workbook.createSheet("ComposicionAccionaria");                logger.debug("Fila {}: Detectado como encabezado o texto descriptivo, saltando...", rowNumber);

                            return null;

            // Crear encabezados            }

            Row headerRow = sheet.createRow(0);            

            headerRow.createCell(0).setCellValue("Entidad");            // Intentar parsear el porcentaje

            headerRow.createCell(1).setCellValue("Accionista");            double percentage = parsePercentage(percentageStr);

            headerRow.createCell(2).setCellValue("Participación");            

                        // Validar rango del porcentaje

            // Escribir datos            if (percentage <= 0 || percentage > 100) {

            int rowNum = 1;                logger.warn("Fila {}: Porcentaje fuera de rango ({}%), saltando...", rowNumber, percentage);

            for (OwnershipRow ownershipRow : validRows) {                return null;

                Row row = sheet.createRow(rowNum++);            }

                row.createCell(0).setCellValue(ownershipRow.entity);            

                row.createCell(1).setCellValue(ownershipRow.shareholder);            return new OwnershipRow(entity.trim(), shareholder.trim(), percentage);

                row.createCell(2).setCellValue(ownershipRow.percentage);            

            }        } catch (Exception e) {

                        logger.debug("Fila {}: Error al parsear - {}", rowNumber, e.getMessage());

            // Autoajustar columnas            return null;

            sheet.autoSizeColumn(0);        }

            sheet.autoSizeColumn(1);    }

            sheet.autoSizeColumn(2);    

                /**

            // Guardar archivo     * Verifica si una celda está vacía.

            try (FileOutputStream fos = new FileOutputStream(outputPath)) {     */

                workbook.write(fos);    /**

            }     * Unifica entidades con nombres similares, específicamente DRA BLUE GLOW INC y DRA BLUE GOW INC

        }     */

    }    private List<OwnershipRow> unifyDraBluEntities(List<OwnershipRow> originalRows) {

            List<OwnershipRow> unifiedRows = new ArrayList<>();

    private boolean isEmptyCell(Cell cell) {        

        if (cell == null) {        // Copiar todas las filas originales

            return true;        unifiedRows.addAll(originalRows);

        }        

                // Buscar entidades DRA BLUE con nombres similares

        CellType cellType = cell.getCellType();        boolean hasGlow = originalRows.stream().anyMatch(row -> 

        if (cellType == CellType.FORMULA) {            "DRA BLUE GLOW INC".equals(row.shareholder));

            cellType = cell.getCachedFormulaResultType();        boolean hasGow = originalRows.stream().anyMatch(row -> 

        }            "DRA BLUE GOW INC".equals(row.entity));

                    

        return switch (cellType) {        if (hasGlow && hasGow) {

            case BLANK -> true;            logger.info("Detectadas entidades DRA BLUE similares - unificando...");

            case STRING -> cell.getStringCellValue().trim().isEmpty();            

            case NUMERIC -> cell.getNumericCellValue() == 0.0;            // Encontrar la relación padre -> DRA BLUE GLOW INC

            default -> false;            OwnershipRow glowRelation = originalRows.stream()

        };                .filter(row -> "DRA BLUE GLOW INC".equals(row.shareholder))

    }                .findFirst()

                    .orElse(null);

    private String getCellValueAsString(Cell cell) {                

        if (cell == null) {            // Encontrar las relaciones DRA BLUE GOW INC -> hijos  

            return null;            List<OwnershipRow> gowChildren = originalRows.stream()

        }                .filter(row -> "DRA BLUE GOW INC".equals(row.entity))

                        .collect(toList());

        CellType cellType = cell.getCellType();                

        if (cellType == CellType.FORMULA) {            if (glowRelation != null && !gowChildren.isEmpty()) {

            cellType = cell.getCachedFormulaResultType();                // Crear nuevas relaciones: DRA BLUE GLOW INC -> hijos de DRA BLUE GOW INC

        }                for (OwnershipRow gowChild : gowChildren) {

                            OwnershipRow newRelation = new OwnershipRow(

        return switch (cellType) {                        "DRA BLUE GLOW INC", 

            case STRING -> cell.getStringCellValue().trim();                        gowChild.shareholder, 

            case NUMERIC -> String.valueOf(cell.getNumericCellValue());                        gowChild.percentage

            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());                    );

            case BLANK -> "";                    unifiedRows.add(newRelation);

            default -> "";                    logger.info("Relación unificada: {} -> {} ({}%)", 

        };                               newRelation.entity, newRelation.shareholder, newRelation.percentage);

    }                }

                    

    private double getCellValueAsDouble(Cell cell) {                // Remover las relaciones originales de DRA BLUE GOW INC

        if (cell == null) {                unifiedRows.removeIf(row -> "DRA BLUE GOW INC".equals(row.entity));

            return 0.0;                logger.info("Removidas relaciones originales de DRA BLUE GOW INC");

        }            }

                }

        CellType cellType = cell.getCellType();        

        if (cellType == CellType.FORMULA) {        return unifiedRows;

            cellType = cell.getCachedFormulaResultType();    }

        }

            private static Stream<OwnershipRow> collect(Collector<OwnershipRow, ?, List<OwnershipRow>> toList) {

        return switch (cellType) {        // TODO Auto-generated method stub

            case NUMERIC -> cell.getNumericCellValue();        throw new UnsupportedOperationException("Unimplemented method 'collect'");

            case STRING -> {    }

                try {

                    yield Double.parseDouble(cell.getStringCellValue().trim());    private static Collector<OwnershipRow, ?, List<OwnershipRow>> toList() {

                } catch (NumberFormatException e) {        // TODO Auto-generated method stub

                    yield 0.0;        throw new UnsupportedOperationException("Unimplemented method 'toList'");

                }    }

            }        if (cell == null) {

            default -> 0.0;            return true;

        };        }

    }        

            CellType cellType = cell.getCellType();

    /**        if (cellType == CellType.FORMULA) {

     * Clase para representar una fila de participación accionaria            cellType = cell.getCachedFormulaResultType();

     */        }

    public static class OwnershipRow {        

        public final String entity;        if (cellType == CellType.BLANK) {

        public final String shareholder;            return true;

        public final double percentage;        }

                

        public OwnershipRow(String entity, String shareholder, double percentage) {        if (cellType == CellType.STRING) {

            this.entity = entity;            return cell.getStringCellValue().trim().isEmpty();

            this.shareholder = shareholder;        }

            this.percentage = percentage;        

        }        if (cellType == CellType.NUMERIC) {

                    return cell.getNumericCellValue() == 0.0;

        @Override        }

        public String toString() {        

            return String.format("%s -> %s (%.2f%%)", entity, shareholder, percentage);        return false;

        }    }

    }    

}    /**
     * Obtiene el valor de una celda como String.
     */
    private String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        
        CellType cellType = cell.getCellType();
        if (cellType == CellType.FORMULA) {
            cellType = cell.getCachedFormulaResultType();
        }
        
        switch (cellType) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                // Si es un número entero, no mostrar decimales
                double numValue = cell.getNumericCellValue();
                if (numValue == (long) numValue) {
                    return String.valueOf((long) numValue);
                }
                return String.valueOf(numValue);
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case BLANK:
                return "";
            default:
                return "";
        }
    }
    
    /**
     * Obtiene el valor de una celda como double.
     */
    private double getCellValueAsDouble(Cell cell) {
        if (cell == null) {
            return 0.0;
        }
        
        CellType cellType = cell.getCellType();
        if (cellType == CellType.FORMULA) {
            cellType = cell.getCachedFormulaResultType();
        }
        
        switch (cellType) {
            case NUMERIC:
                return cell.getNumericCellValue();
            case STRING:
                try {
                    String strValue = cell.getStringCellValue().trim();
                    if (strValue.isEmpty()) {
                        return 0.0;
                    }
                    // Intentar parsear como número
                    return Double.parseDouble(strValue.replace("%", "").replace(",", "."));
                } catch (NumberFormatException e) {
                    return 0.0;
                }
            default:
                return 0.0;
        }
    }
    
    /**
     * Obtiene el valor de una celda como String (método legacy para compatibilidad).
     */
    private String getCellValue(Cell cell) {
        return getCellValueAsString(cell);
    }
    
    /**
     * Detecta si una fila contiene encabezados o texto descriptivo.
     */
    private boolean isHeaderOrDescriptiveText(String col1, String col2, String col3) {
        String combined = (col1 + col2 + col3).toLowerCase();
        
        // Palabras clave que indican encabezados o texto descriptivo
        String[] keywords = {
            "entidad", "accionista", "participación", "participacion",
            "origen", "vía", "via", "través", "traves", "a través",
            "nombre", "beneficiario", "titular", "final",
            "%", "porcentaje", "directa", "indirecta",
            "composición", "composicion", "cálculo", "calculo"
        };
        
        for (String keyword : keywords) {
            if (combined.contains(keyword)) {
                return true;
            }
        }
        
        return false;
    }
    
    /**
     * Parsea un valor de porcentaje en diferentes formatos.
     */
    private double parsePercentage(String value) {
        if (value == null || value.isEmpty()) {
            throw new IllegalArgumentException("Porcentaje vacío");
        }
        
        // Limpiar el valor
        String cleaned = value.trim()
                             .replace("%", "")
                             .replace(",", ".")
                             .replace(" ", "");
        
        try {
            double result = Double.parseDouble(cleaned);
            
            // Si el valor está entre 0 y 1, asumir que ya está normalizado y convertir a 0-100
            if (result > 0 && result < 1) {
                result = result * 100;
            }
            
            return result;
        } catch (NumberFormatException e) {
            throw new IllegalArgumentException("Formato de porcentaje inválido: " + value);
        }
    }
    
    /**
     * Crea un nuevo archivo Excel limpio con las filas válidas.
     */
    private void createCleanExcel(List<OwnershipRow> rows, String outputPath) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Composición Accionaria");
            
            // Crear encabezado
            Row headerRow = sheet.createRow(0);
            createCell(headerRow, 0, "Entidad");
            createCell(headerRow, 1, "Accionista");
            createCell(headerRow, 2, "Participación");
            
            // Agregar filas de datos
            int rowIndex = 1;
            for (OwnershipRow ownershipRow : rows) {
                Row dataRow = sheet.createRow(rowIndex++);
                createCell(dataRow, 0, ownershipRow.entity);
                createCell(dataRow, 1, ownershipRow.shareholder);
                createCell(dataRow, 2, ownershipRow.percentage);
            }
            
            // Ajustar ancho de columnas
            sheet.autoSizeColumn(0);
            sheet.autoSizeColumn(1);
            sheet.autoSizeColumn(2);
            
            // Guardar el archivo
            try (FileOutputStream fos = new FileOutputStream(outputPath)) {
                workbook.write(fos);
            }
        }
    }
    
    /**
     * Crea una celda con valor String.
     */
    private void createCell(Row row, int column, String value) {
        Cell cell = row.createCell(column);
        cell.setCellValue(value);
    }
    
    /**
     * Crea una celda con valor numérico.
     */
    private void createCell(Row row, int column, double value) {
        Cell cell = row.createCell(column);
        cell.setCellValue(value);
    }
    
    /**
     * Clase interna para representar una fila de participación accionaria.
     */
    private static class OwnershipRow {
        final String entity;
        final String shareholder;
        final double percentage;
        
        OwnershipRow(String entity, String shareholder, double percentage) {
            this.entity = entity;
            this.shareholder = shareholder;
            this.percentage = percentage;
        }
    }
}
