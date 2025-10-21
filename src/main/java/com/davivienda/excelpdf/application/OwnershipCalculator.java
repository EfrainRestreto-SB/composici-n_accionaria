package com.davivienda.excelpdf.application;

import com.davivienda.excelpdf.domain.Node;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

/**
 * Calculadora de participaciones accionarias que procesa archivos Excel
 * y calcula recursivamente la composición final de beneficiarios.
 * Incluye detección de ciclos y validaciones robustas.
 * 
 * @author Davivienda
 * @version 1.0
 */
public class OwnershipCalculator {
    
    private static final Logger logger = LoggerFactory.getLogger(OwnershipCalculator.class);
    
    private final Map<String, Node> graph = new HashMap<>();
    private final Map<String, Double> finalResults = new HashMap<>();
    private final Map<String, String> beneficiaryPaths = new HashMap<>();
    private final Set<String> visitedInCurrentPath = new HashSet<>(); // Para detección de ciclos
    
    /**
     * Obtiene o crea un nodo en el grafo.
     * 
     * @param name nombre de la entidad
     * @return nodo existente o nuevo nodo
     */
    private Node getOrCreateNode(String name) {
        return graph.computeIfAbsent(name.trim(), Node::new);
    }
    
    /**
     * Carga las relaciones de propiedad desde un archivo Excel.
     * Formato esperado: Columna A: Entidad, Columna B: Accionista, Columna C: % Participación
     * 
     * @param excelPath ruta del archivo Excel
     * @throws IOException si hay problemas de acceso al archivo
     * @throws IllegalArgumentException si el formato del Excel es inválido
     */
    public void loadFromExcel(String excelPath) throws IOException {
        logger.info("Cargando datos desde: {}", excelPath);
        
        try (FileInputStream fileInputStream = new FileInputStream(excelPath);
             XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream)) {
            
            Sheet sheet = workbook.getSheetAt(0);
            int rowCount = 0;
            int validRowCount = 0;
            
            // Validar que exista al menos una fila de encabezado
            if (sheet.getLastRowNum() < 1) {
                throw new IllegalArgumentException("El archivo Excel debe tener al menos una fila de datos además del encabezado");
            }
            
            // Procesar filas (asumiendo que la primera fila son encabezados)
            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                rowCount++;
                
                if (row == null) {
                    logger.warn("Fila {} está vacía, saltando...", rowIndex + 1);
                    continue;
                }
                
                try {
                    // Extraer datos de las celdas
                    String entity = getCellStringValue(row.getCell(0));
                    String owner = getCellStringValue(row.getCell(1));
                    double percentage = getCellNumericValue(row.getCell(2));
                    
                    // Validar datos
                    if (entity.isEmpty() || owner.isEmpty()) {
                        logger.warn("Fila {}: Entidad o Accionista vacío, saltando...", rowIndex + 1);
                        continue;
                    }
                    
                    if (percentage <= 0 || percentage > 100) {
                        logger.warn("Fila {}: Porcentaje inválido ({}%), debe estar entre 0 y 100", rowIndex + 1, percentage);
                        continue;
                    }
                    
                    // Convertir porcentaje de 0-100 a 0-1
                    double normalizedPercentage = percentage / 100.0;
                    
                    // Crear nodos y relación
                    Node entityNode = getOrCreateNode(entity);
                    Node ownerNode = getOrCreateNode(owner);
                    entityNode.addOwner(ownerNode, normalizedPercentage);
                    
                    validRowCount++;
                    logger.debug("Procesada relación: {} -> {} ({}%)", entity, owner, percentage);
                    
                } catch (Exception e) {
                    logger.error("Error procesando fila {}: {}", rowIndex + 1, e.getMessage());
                    throw new IllegalArgumentException("Error en fila " + (rowIndex + 1) + ": " + e.getMessage());
                }
            }
            
            logger.info("Carga completada. Filas procesadas: {}, Relaciones válidas: {}, Entidades: {}", 
                       rowCount, validRowCount, graph.size());
            
            if (validRowCount == 0) {
                throw new IllegalArgumentException("No se encontraron relaciones válidas en el archivo Excel");
            }
            
            // Validar integridad del grafo
            validateGraphIntegrity();
        }
    }
    
    /**
     * Obtiene el valor string de una celda, manejando diferentes tipos.
     */
    private String getCellStringValue(Cell cell) {
        if (cell == null) return "";
        
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                return String.valueOf((long) cell.getNumericCellValue()).trim();
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue()).trim();
            default:
                return "";
        }
    }
    
    /**
     * Obtiene el valor numérico de una celda.
     */
    private double getCellNumericValue(Cell cell) {
        if (cell == null) {
            throw new IllegalArgumentException("Celda de porcentaje no puede estar vacía");
        }
        
        switch (cell.getCellType()) {
            case NUMERIC:
                return cell.getNumericCellValue();
            case STRING:
                try {
                    return Double.parseDouble(cell.getStringCellValue().trim());
                } catch (NumberFormatException e) {
                    throw new IllegalArgumentException("Valor de porcentaje inválido: " + cell.getStringCellValue());
                }
            default:
                throw new IllegalArgumentException("Tipo de celda no soportado para porcentaje: " + cell.getCellType());
        }
    }
    
    /**
     * Valida la integridad del grafo después de la carga.
     */
    private void validateGraphIntegrity() {
        for (Node node : graph.values()) {
            try {
                node.validateOwnership();
            } catch (IllegalStateException e) {
                logger.warn("Advertencia de integridad: {}", e.getMessage());
                // No lanzamos excepción para permitir casos donde la suma sea menor al 100%
            }
        }
    }
    
    /**
     * Calcula las participaciones finales desde una entidad raíz.
     * 
     * @param rootEntityName nombre de la entidad raíz
     * @throws IllegalArgumentException si la entidad raíz no existe
     */
    public void calculateFinalOwnership(String rootEntityName) {
        logger.info("Calculando participaciones finales desde: {}", rootEntityName);
        
        Node rootNode = graph.get(rootEntityName.trim());
        if (rootNode == null) {
            throw new IllegalArgumentException("Entidad raíz no encontrada: " + rootEntityName);
        }
        
        // Limpiar resultados anteriores
        finalResults.clear();
        beneficiaryPaths.clear();
        visitedInCurrentPath.clear();
        
        // Iniciar cálculo recursivo
        calculateOwnershipRecursive(rootNode, 1.0, rootEntityName, new HashSet<>());
        
        logger.info("Cálculo completado. Beneficiarios finales encontrados: {}", finalResults.size());
    }
    
    /**
     * Cálculo recursivo de participaciones con detección de ciclos.
     * 
     * @param node nodo actual
     * @param accumulatedPercentage porcentaje acumulado hasta este nodo
     * @param path ruta completa hasta este nodo
     * @param visitedInPath nodos visitados en la ruta actual (para detectar ciclos)
     */
    private void calculateOwnershipRecursive(Node node, double accumulatedPercentage, 
                                           String path, Set<String> visitedInPath) {
        
        String nodeName = node.getName();
        
        // Detectar ciclos
        if (visitedInPath.contains(nodeName)) {
            logger.warn("Ciclo detectado en la ruta: {} -> {}", path, nodeName);
            // Tratamos el nodo como beneficiario final para evitar recursión infinita
            finalResults.merge(nodeName, accumulatedPercentage, Double::sum);
            beneficiaryPaths.put(nodeName, path + " [CICLO DETECTADO]");
            return;
        }
        
        // Si no tiene propietarios, es un beneficiario final
        if (!node.hasOwners()) {
            finalResults.merge(nodeName, accumulatedPercentage, Double::sum);
            beneficiaryPaths.put(nodeName, path);
            logger.debug("Beneficiario final: {} ({}%)", nodeName, accumulatedPercentage * 100);
            return;
        }
        
        // Continuar recursión con los propietarios
        Set<String> newVisitedPath = new HashSet<>(visitedInPath);
        newVisitedPath.add(nodeName);
        
        for (Map.Entry<Node, Double> ownerEntry : node.getOwners().entrySet()) {
            Node owner = ownerEntry.getKey();
            double ownershipPercentage = ownerEntry.getValue();
            double newAccumulatedPercentage = accumulatedPercentage * ownershipPercentage;
            String newPath = path + " → " + owner.getName();
            
            calculateOwnershipRecursive(owner, newAccumulatedPercentage, newPath, newVisitedPath);
        }
    }
    
    /**
     * Obtiene los resultados finales de participación.
     * 
     * @return mapa de beneficiario -> porcentaje final (0.0 - 1.0)
     */
    public Map<String, Double> getFinalResults() {
        return new HashMap<>(finalResults);
    }
    
    /**
     * Obtiene las rutas completas hacia cada beneficiario.
     * 
     * @return mapa de beneficiario -> ruta completa
     */
    public Map<String, String> getBeneficiaryPaths() {
        return new HashMap<>(beneficiaryPaths);
    }
    
    /**
     * Obtiene el grafo completo de entidades.
     * 
     * @return mapa de nombre -> nodo
     */
    public Map<String, Node> getGraph() {
        return new HashMap<>(graph);
    }
    
    /**
     * Obtiene estadísticas del grafo cargado.
     * 
     * @return estadísticas como string
     */
    public String getGraphStatistics() {
        int totalEntities = graph.size();
        int entitiesWithOwners = (int) graph.values().stream().filter(Node::hasOwners).count();
        int finalBeneficiaries = totalEntities - entitiesWithOwners;
        
        return String.format("Estadísticas del grafo: %d entidades totales, %d con propietarios, %d beneficiarios finales",
                           totalEntities, entitiesWithOwners, finalBeneficiaries);
    }
}