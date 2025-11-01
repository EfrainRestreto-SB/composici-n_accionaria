def is_final_entity_debug(entidad):
    entidad_upper = entidad.upper()
    print(f'DEBUG: Analizando entidad: "{entidad}" -> "{entidad_upper}"')
    
    # Lista base de nombres de personas conocidos
    known_person_indicators = [
        'RODRIGUEZ', 'DIAZ', 'CONTRERAS', 'MERCEDES', 'STELLA', 
        'ALEXANDRA', 'CARLOS', 'JORGE', 'ARTURO', 'ENRIQUE', 'LUZ'
    ]
    
    # Entidades operativas conocidas
    operational_entities = [
        'INVERSIONES MADCOM', 'MADCOM'
    ]
    
    # Verificar patrones conocidos
    for indicator in known_person_indicators:
        if indicator in entidad_upper:
            print(f'  -> DETECTADO como persona por indicador: {indicator}')
            return True
    
    # Verificar entidades operativas
    for entity in operational_entities:
        if entity in entidad_upper:
            print(f'  -> DETECTADO como entidad operativa: {entity}')
            return True
    
    # Detección dinámica
    words = entidad_upper.split()
    print(f'  -> Palabras: {words}')
    
    if len(words) >= 2:
        # Patrones de nombres
        name_patterns = ['MARIA', 'FERNANDEZ', 'LOPEZ', 'JUAN', 'PEDRO', 'ANA', 'JOSE']
        for pattern in name_patterns:
            if pattern in entidad_upper:
                print(f'  -> DETECTADO como persona por patrón dinámico: {pattern}')
                return True
        
        # Verificar palabras corporativas
        corporate_words = ['INC', 'S.A', 'SAS', 'LTDA', 'CORPORATION', 'COMPANY', 'FINANCIAL', 'INVERSIONES']
        is_corporate = any(corp_word in entidad_upper for corp_word in corporate_words)
        print(f'  -> Es corporativo: {is_corporate}')
        print(f'  -> Palabras corporativas encontradas: {[word for word in corporate_words if word in entidad_upper]}')
        
        if not is_corporate and len(words) >= 2:
            print(f'  -> DETECTADO como persona (no corporativo con 2+ palabras)')
            return True
    
    print(f'  -> NO es beneficiario final')
    return False

# Probar con las entidades problemáticas
print('=== PROBANDO DRA BLUE GOW INC ===')
result1 = is_final_entity_debug('DRA BLUE GOW INC')
print(f'Resultado: {result1}')
print()

print('=== PROBANDO DRA BLUE GOW ===')
result3 = is_final_entity_debug('DRA BLUE GOW')
print(f'Resultado: {result3}')
print()

print('=== PROBANDO ALEXANDRA DIAZ RODRIGUEZ ===')
result2 = is_final_entity_debug('ALEXANDRA DIAZ RODRIGUEZ')
print(f'Resultado: {result2}')