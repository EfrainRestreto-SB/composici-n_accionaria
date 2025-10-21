package com.davivienda.excelpdf.domain;

import java.util.HashMap;
import java.util.Map;

/**
 * Representa un nodo en el grafo de participaciones accionarias.
 * Cada nodo puede tener múltiples propietarios con sus respectivos porcentajes.
 * 
 * @author Davivienda
 * @version 1.0
 */
public class Node {
    
    private final String name;
    private final Map<Node, Double> owners = new HashMap<>();
    
    /**
     * Constructor del nodo.
     * 
     * @param name nombre de la entidad que representa este nodo
     */
    public Node(String name) {
        if (name == null || name.trim().isEmpty()) {
            throw new IllegalArgumentException("El nombre del nodo no puede estar vacío");
        }
        this.name = name.trim();
    }
    
    /**
     * Agrega un propietario a este nodo con su porcentaje de participación.
     * 
     * @param owner nodo propietario
     * @param percentage porcentaje de participación (0.0 - 1.0)
     * @throws IllegalArgumentException si el porcentaje es inválido
     */
    public void addOwner(Node owner, double percentage) {
        if (owner == null) {
            throw new IllegalArgumentException("El propietario no puede ser null");
        }
        
        if (percentage <= 0.0 || percentage > 1.0) {
            throw new IllegalArgumentException(
                String.format("Porcentaje inválido: %.2f%%. Debe estar entre 0 y 100%%", percentage * 100)
            );
        }
        
        if (owner.equals(this)) {
            throw new IllegalArgumentException("Un nodo no puede ser propietario de sí mismo");
        }
        
        owners.put(owner, percentage);
    }
    
    /**
     * Obtiene todos los propietarios de este nodo.
     * 
     * @return mapa de propietario -> porcentaje
     */
    public Map<Node, Double> getOwners() {
        return new HashMap<>(owners);
    }
    
    /**
     * Verifica si este nodo tiene propietarios.
     * 
     * @return true si tiene propietarios, false si es un beneficiario final
     */
    public boolean hasOwners() {
        return !owners.isEmpty();
    }
    
    /**
     * Obtiene el nombre de la entidad.
     * 
     * @return nombre de la entidad
     */
    public String getName() {
        return name;
    }
    
    /**
     * Valida que la suma de participaciones no exceda el 100%.
     * 
     * @throws IllegalStateException si la suma excede el 100%
     */
    public void validateOwnership() {
        if (owners.isEmpty()) {
            return; // Los beneficiarios finales no necesitan validación
        }
        
        double totalPercentage = owners.values().stream()
            .mapToDouble(Double::doubleValue)
            .sum();
        
        if (totalPercentage > 1.01) { // Permitir pequeños errores de redondeo
            throw new IllegalStateException(
                String.format("La suma de participaciones de '%s' excede el 100%% (%.2f%%)", 
                             name, totalPercentage * 100)
            );
        }
    }
    
    /**
     * Obtiene el porcentaje total de participación distribuida.
     * 
     * @return porcentaje total distribuido (0.0 - 1.0)
     */
    public double getTotalDistributedPercentage() {
        return owners.values().stream()
            .mapToDouble(Double::doubleValue)
            .sum();
    }
    
    /**
     * Obtiene el número de propietarios directos.
     * 
     * @return número de propietarios
     */
    public int getOwnerCount() {
        return owners.size();
    }
    
    @Override
    public boolean equals(Object obj) {
        if (this == obj) return true;
        if (obj == null || getClass() != obj.getClass()) return false;
        Node node = (Node) obj;
        return name.equals(node.name);
    }
    
    @Override
    public int hashCode() {
        return name.hashCode();
    }
    
    @Override
    public String toString() {
        return String.format("Node{name='%s', owners=%d, totalDistributed=%.2f%%}", 
                           name, owners.size(), getTotalDistributedPercentage() * 100);
    }
}