def is_final_entity_simple(entidad):
    entidad_upper = entidad.upper()
    
    known_person_indicators = [
        'RODRIGUEZ', 'DIAZ', 'CONTRERAS', 'MERCEDES', 'STELLA', 
        'ALEXANDRA', 'CARLOS', 'JORGE', 'ARTURO', 'ENRIQUE', 'LUZ'
    ]
    
    operational_entities = [
        'INVERSIONES MADCOM', 'MADCOM'
    ]
    
    # Verificar patrones conocidos
    for indicator in known_person_indicators:
        if indicator in entidad_upper:
            return True
    
    # Verificar entidades operativas
    for entity in operational_entities:
        if entity in entidad_upper:
            return True
    
    # Detección dinámica
    words = entidad_upper.split()
    if len(words) >= 2:
        name_patterns = ['MARIA', 'FERNANDEZ', 'LOPEZ', 'JUAN', 'PEDRO', 'ANA', 'JOSE']
        for pattern in name_patterns:
            if pattern in entidad_upper:
                return True
        
        corporate_words = ['INC', 'S.A', 'SAS', 'LTDA', 'CORPORATION', 'COMPANY', 'FINANCIAL', 'INVERSIONES']
        is_corporate = any(corp_word in entidad_upper for corp_word in corporate_words)
        
        if not is_corporate and len(words) >= 2:
            return True
    
    return False

# Lista de entidades con participaciones finales no nulas
entities = [
    'RED COW INC',
    'POWER FINANCIAL S.A',
    'BLACK LAB INC', 
    'CARLOS ARTURO DIAZ RODRIGUEZ',
    'TIERRA ARCO IRIS',
    'LUZ MERCEDES DIAZ RODRIGUEZ',
    'DRA BLUE GLOW INC',
    'BLACK BULL CORPORATION',
    'INVERSIONES MADCOM',
    'STELLA RODRIGUEZ CONTRERAS',
    'ALEXANDRA DIAZ RODRIGUEZ',
    'JORGE ENRIQUE DIAZ RODRIGUEZ'
]

print('=== ANALISIS DE ENTIDADES ===')
final_beneficiaries = []

for entity in entities:
    is_final = is_final_entity_simple(entity)
    status = "SI" if is_final else "NO"
    print(f'{entity}: {status}')
    if is_final:
        final_beneficiaries.append(entity)

print()
print(f'Total beneficiarios finales únicos: {len(set(final_beneficiaries))}')
print(f'Beneficiarios: {list(set(final_beneficiaries))}')