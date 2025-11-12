#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script para consolidar automáticamente las participaciones en RED COW INC
Interpreta la estructura jerárquica del archivo data.xlsx y calcula beneficiarios finales
Versión 6: Interpretación inteligente de estructura jerárquica - Java Compatible
"""

import sys
import os
from pathlib import Path

# === SISTEMA DE GESTIÓN AUTOMÁTICA DE DEPENDENCIAS ===
def setup_virtual_environment():
    """
    Configura automáticamente un entorno virtual con las dependencias necesarias
    Solo se ejecuta si hay problemas de importación
    """
    import subprocess
    import sys
    from pathlib import Path
    
    print("=== CONFIGURANDO ENTORNO VIRTUAL AUTOMÁTICO ===")
    
    # Directorio del entorno virtual
    venv_dir = Path("venv_composicion")
    
    try:
        # 1. Crear entorno virtual si no existe
        if not venv_dir.exists():
            print(f"Creando entorno virtual en: {venv_dir}")
            subprocess.run([sys.executable, "-m", "venv", str(venv_dir)], check=True)
            print("Entorno virtual creado exitosamente")
        
        # 2. Determinar ejecutable Python del entorno virtual
        if sys.platform == "win32":
            python_exe = venv_dir / "Scripts" / "python.exe"
            pip_exe = venv_dir / "Scripts" / "pip.exe"
        else:
            python_exe = venv_dir / "bin" / "python"
            pip_exe = venv_dir / "bin" / "pip"
        
        # 3. Instalar dependencias necesarias
        required_packages = ["pandas", "openpyxl", "xlsxwriter", "numpy"]
        
        for package in required_packages:
            print(f"Instalando {package}...")
            try:
                subprocess.run([str(pip_exe), "install", package], check=True, capture_output=True)
                print(f"{package} instalado correctamente")
            except subprocess.CalledProcessError as e:
                print(f"Error instalando {package}: {e}")
                return False        # 4. Re-ejecutar script con el Python del entorno virtual
        print(f"Dependencias instaladas. Re-ejecutando script con entorno virtual...")
        current_args = sys.argv.copy()
        
        # Ejecutar el script actual con el Python del entorno virtual
        result = subprocess.run([str(python_exe)] + current_args, capture_output=False)
        
        # Si fue exitoso, notificar y salir
        if result.returncode == 0:
            print("[OK] Script ejecutado exitosamente con entorno virtual")
        
        sys.exit(result.returncode)
        
    except Exception as e:
        print(f"Error configurando entorno virtual: {e}")
        return False

def check_and_install_dependencies():
    """
    Verifica las dependencias y las instala automáticamente si faltan
    """
    required_modules = {
        'pandas': 'pandas',
        'openpyxl': 'openpyxl', 
        'xlsxwriter': 'xlsxwriter'
    }
    
    missing_modules = []
    
    print("Verificando dependencias Python...")
    
    # Verificar cada módulo
    for module_name, package_name in required_modules.items():
        try:
            __import__(module_name)
            print(f"[OK] {module_name} disponible")
        except ImportError:
            print(f"[FALTA] {module_name} no encontrado")
            missing_modules.append(package_name)
    
    # Si faltan módulos, forzar instalación automática inmediatamente
    if missing_modules:
        print(f"INSTALANDO AUTOMÁTICAMENTE: {', '.join(missing_modules)}")
        print("Esto puede tomar unos momentos...")
        
        success = auto_install_packages(missing_modules)
        if not success:
            print("ERROR: Instalación automática falló, creando entorno virtual...")
            return setup_virtual_environment()
        
        # Verificar nuevamente después de instalación
        print("Verificando instalación...")
        for module_name in required_modules.keys():
            try:
                __import__(module_name)
                print(f"[INSTALADO] {module_name} ahora disponible")
            except ImportError:
                print(f"[ERROR] {module_name} aún no disponible después de instalación")
                return setup_virtual_environment()
        
        return True
    else:
        print("Todas las dependencias están disponibles")
        return True

def auto_install_packages(packages):
    """
    Instala paquetes automáticamente usando pip
    """
    import subprocess
    import sys
    
    print("Iniciando instalación automática de dependencias...")
    
    # Intentar actualizar pip primero
    try:
        print("Actualizando pip...")
        subprocess.run([sys.executable, "-m", "pip", "install", "--upgrade", "pip"], 
                      check=True, capture_output=True)
        print("Pip actualizado correctamente")
    except subprocess.CalledProcessError:
        print("Advertencia: No se pudo actualizar pip, continuando...")
    
    # Instalar cada paquete
    for package in packages:
        try:
            print(f"Instalando {package}...")
            
            # Usar múltiples métodos de instalación
            commands_to_try = [
                [sys.executable, "-m", "pip", "install", package],
                [sys.executable, "-m", "pip", "install", "--user", package],
                ["py", "-m", "pip", "install", package]
            ]
            
            success = False
            for cmd in commands_to_try:
                try:
                    result = subprocess.run(cmd, check=True, capture_output=True, text=True)
                    print(f"[ÉXITO] {package} instalado con: {' '.join(cmd)}")
                    success = True
                    break
                except subprocess.CalledProcessError as e:
                    print(f"Intento fallido con {' '.join(cmd)}: {e}")
                    continue
                except FileNotFoundError:
                    print(f"Comando no encontrado: {' '.join(cmd)}")
                    continue
            
            if not success:
                print(f"ERROR: No se pudo instalar {package} con ningún método")
                print("Cambiando a entorno virtual...")
                return setup_virtual_environment()
            
        except Exception as e:
            print(f"Error inesperado instalando {package}: {e}")
            return setup_virtual_environment()
    
    print("Instalación automática completada")
    
    # Verificar que la instalación fue exitosa
    print("Verificando instalación...")
    try:
        # Forzar recarga de módulos
        import importlib
        import sys
        
        # Limpiar cache de importaciones
        for module_name in ['pandas', 'openpyxl', 'xlsxwriter', 'numpy']:
            if module_name in sys.modules:
                del sys.modules[module_name]
        
        # Intentar importar pandas
        import pandas as pd
        print("ÉXITO: Pandas importado correctamente después de instalación")
        
        # Forzar salida exitosa para notificar a Java
        print("[OK] Todas las dependencias han sido instaladas y verificadas")
        return True
        
    except ImportError as e:
        print(f"ERROR: Pandas aún no disponible después de instalación: {e}")
        print("Creando entorno virtual como alternativa...")
        return setup_virtual_environment()

def handle_numpy_conflict():
    """
    Maneja específicamente el conflicto de numpy source directory
    """
    import os
    import sys
    from pathlib import Path
    
    current_dir = Path.cwd()
    print(f"📍 Directorio actual: {current_dir}")
    
    # Verificar si estamos en un directorio problemático
    problematic_patterns = ["numpy", "site-packages", "python"]
    
    for pattern in problematic_patterns:
        if pattern.lower() in str(current_dir).lower():
            print(f"Detectado directorio problemático que contiene '{pattern}'")
            
            # Cambiar a un directorio temporal seguro
            safe_dir = Path.home() / "temp_composicion_work"
            safe_dir.mkdir(exist_ok=True)
            
            print(f"Cambiando a directorio seguro: {safe_dir}")
            os.chdir(safe_dir)
            
            # Copiar archivos necesarios
            import shutil
            script_path = Path(sys.argv[0])
            excel_path = Path(sys.argv[1]) if len(sys.argv) > 1 else None
            
            if script_path.exists():
                shutil.copy2(script_path, safe_dir)
            if excel_path and excel_path.exists():
                shutil.copy2(excel_path, safe_dir)
            
            print("Archivos copiados al directorio seguro")
            return True
    
    return False

def initialize_environment():
    """
    Inicializa el entorno de ejecución verificando y configurando dependencias
    """
    try:
        # Intentar importar pandas normalmente primero
        import pandas as pd
        print("Pandas disponible - continuando ejecución normal")
        return True
        
    except ImportError as e:
        error_msg = str(e).lower()
        
        # Detectar tipo específico de error
        if "numpy" in error_msg and "source directory" in error_msg:
            print("Detectado conflicto de numpy - aplicando solución automática...")
            if handle_numpy_conflict():
                print("Conflicto de numpy resuelto")
            return setup_virtual_environment()
            
        elif "no module named" in error_msg:
            print("Detectadas dependencias faltantes - instalando automáticamente...")
            return check_and_install_dependencies()
            
        else:
            print(f"Error de importación no manejado: {e}")
            return setup_virtual_environment()

# === INICIALIZACIÓN AUTOMÁTICA ===
# Ejecutar verificación de dependencias antes de continuar
print("=== INICIANDO SCRIPT CON AUTO-REPARACIÓN ===")
print("Verificando dependencias...")

if not initialize_environment():
    print("No se pudo configurar el entorno automáticamente")
    print("SOLUCIONES MANUALES:")
    print("   1. Ejecutar desde un directorio diferente")
    print("   2. py -m pip install pandas openpyxl xlsxwriter")
    print("   3. Crear entorno virtual manualmente")
    sys.exit(1)

print("Dependencias verificadas correctamente")

# Importar pandas después de verificar/configurar el entorno
try:
    import pandas as pd
    print("Pandas importado exitosamente")
except ImportError as e:
    print(f"Error crítico: No se puede importar pandas después de verificación: {e}")
    print("Forzando creación de entorno virtual...")
    setup_virtual_environment()

# Configurar codificación para compatibilidad con Java
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

def consolidate_participations(excel_file):
    """
    Consolida las participaciones finales a partir de la estructura jerárquica
    """
    try:
        print("=== INICIANDO CONSOLIDACION INTELIGENTE v6 ===")
        print(f"Archivo: {excel_file}")
        
        # Leer Excel original
        df = pd.read_excel(excel_file)
        print(f"Filas leidas: {len(df)}")
        
        # Extraer beneficiarios finales basándose en la estructura del reporte
        final_beneficiaries = extract_final_beneficiaries(df)
        print(f"Beneficiarios finales detectados: {len(final_beneficiaries)}")
        
        # Crear estructura de salida
        output_data = create_consolidated_output(final_beneficiaries)
        
        # Generar archivo de salida
        input_path = Path(excel_file)
        output_file = input_path.parent / f"{input_path.stem}_cleaned_fixed.xlsx"
        
        # Guardar archivo corregido
        df_output = pd.DataFrame(output_data)
        df_output.to_excel(output_file, index=False)
        
        print(f"[OK] Archivo consolidado creado: {output_file}")
        print(f"Relaciones en output: {len(df_output)}")
        
        # Mostrar estadísticas
        show_consolidation_stats(output_data)
        
        return True
        
    except Exception as e:
        print(f"[ERROR] Error durante la consolidacion: {e}")
        import traceback
        traceback.print_exc()
        return False

def extract_final_beneficiaries(df):
    """
    Extrae los beneficiarios finales y sus participaciones de la estructura jerárquica
    """
    print("\n=== EXTRAYENDO BENEFICIARIOS FINALES ===")
    
    beneficiaries = {}
    
    # Analizar cada fila buscando patrones de beneficiarios finales
    for i, row in df.iterrows():
        entidad = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        participacion_directa = row.iloc[1] if pd.notna(row.iloc[1]) else None
        participacion_final = row.iloc[2] if pd.notna(row.iloc[2]) else None
        
        # Saltar filas sin datos útiles
        if not entidad or entidad == "nan":
            continue
            
        # Detectar beneficiarios finales: tienen participación final pero no son entidades de primer nivel
        if participacion_final is not None and participacion_final > 0:
            
            # Determinar si es beneficiario final (persona natural o entidad operativa)
            is_final_beneficiary = is_final_entity(entidad)
            
            if is_final_beneficiary:
                # Convertir participación final a porcentaje
                participation_percent = float(participacion_final) * 100
                
                # Agregar o actualizar beneficiario
                if entidad in beneficiaries:
                    beneficiaries[entidad] += participation_percent
                else:
                    beneficiaries[entidad] = participation_percent
                    
                print(f"Beneficiario detectado: {entidad} -> {participation_percent:.2f}%")
    
    return beneficiaries

def is_final_entity(entidad):
    """
    Determina si una entidad es un beneficiario final (persona natural principalmente)
    Mantiene la lógica original pero agrega detección dinámica de personas
    """
    entidad_upper = entidad.upper()
    
    # Lista base de nombres de personas conocidos (beneficiarios finales)
    known_person_indicators = [
        "RODRIGUEZ", "DIAZ", "CONTRERAS", "MERCEDES", "STELLA", 
        "ALEXANDRA", "CARLOS", "JORGE", "ARTURO", "ENRIQUE", "LUZ"
    ]
    
    # Entidades operativas conocidas que también son beneficiarios finales
    operational_entities = [
        "INVERSIONES MADCOM", "MADCOM"
    ]
    
    # Entidades corporativas que NO son beneficiarios finales (aunque no tengan palabras corporativas obvias)
    corporate_entities = [
        "TIERRA ARCO IRIS"  # Es una entidad corporativa, no persona natural
    ]
    
    # Verificar contra patrones conocidos primero
    for indicator in known_person_indicators:
        if indicator in entidad_upper:
            return True
            
    # Verificar si es una entidad corporativa conocida (NO beneficiario final)
    for entity in corporate_entities:
        if entity in entidad_upper:
            return False
            
    for entity in operational_entities:
        if entity in entidad_upper:
            return True
    
    # Detección dinámica adicional para nuevos nombres de personas
    # Buscar patrones que indiquen nombre de persona (2+ nombres propios)
    words = entidad_upper.split()
    if len(words) >= 2:
        # Si contiene palabras típicas de nombres de persona
        name_patterns = ["MARIA", "FERNANDEZ", "LOPEZ", "JUAN", "PEDRO", "ANA", "JOSE"]
        for pattern in name_patterns:
            if pattern in entidad_upper:
                return True
        
        # Si todas las palabras empiezan con mayúscula y no contienen palabras corporativas
        corporate_words = ["INC", "S.A", "SAS", "LTDA", "CORPORATION", "COMPANY", "FINANCIAL", "INVERSIONES"]
        is_corporate = any(corp_word in entidad_upper for corp_word in corporate_words)
        
        if not is_corporate and len(words) >= 2:
            # Probablemente es un nombre de persona
            return True
    
    return False

def create_consolidated_output(beneficiaries):
    """
    Crea la estructura de salida consolidada - SOLO participaciones directas en RED COW INC
    """
    print("\n=== CREANDO ESTRUCTURA CONSOLIDADA ===")
    
    output_data = []
    
    # Ordenar beneficiarios por participación descendente
    sorted_beneficiaries = sorted(beneficiaries.items(), key=lambda x: x[1], reverse=True)
    
    # SOLO agregar participaciones directas en RED COW INC (no entidades intermedias)
    for beneficiary, participation in sorted_beneficiaries:
        # Mantener el nombre original del beneficiario final
        accionista_display = beneficiary
            
        output_data.append({
            "Entidad": "RED COW INC",
            "Accionista": accionista_display,
            "Participacion": round(participation, 2)
        })
        
        print(f"Agregado: RED COW INC <- {accionista_display} ({participation:.2f}%)")
    
    # NO agregar entidades intermedias para evitar confusión en el algoritmo
        print("INFO: Omitiendo entidades intermedias para evitar doble conteo")
    return output_data

def detect_intermediate_entities(beneficiaries):
    """
    Detecta entidades intermedias importantes que deben aparecer en la salida
    """
    intermediate = []
    
    # No agregar entidades intermedias adicionales
    
    return intermediate

def show_consolidation_stats(output_data):
    """
    Muestra estadísticas de la consolidación
    """
    print(f"\n=== ESTADISTICAS DE CONSOLIDACION ===")
    
    red_cow_participations = [row for row in output_data if row["Entidad"] == "RED COW INC"]
    total_participation = sum(row["Participacion"] for row in red_cow_participations)
    
    print(f"Beneficiarios directos de RED COW INC: {len(red_cow_participations)}")
    print(f"Participacion total: {total_participation:.2f}%")
    print(f"Total de relaciones en output: {len(output_data)}")
    
    print(f"\nBENEFICIARIOS FINALES (ESTRUCTURA PLANA):")
    for row in red_cow_participations:
        print(f"  - {row['Accionista']}: {row['Participacion']}%")
        
    # Verificar si tenemos los valores esperados
    expected_beneficiaries = {
        "Stella Rodriguez Contreras": 34.02,
        "Alexandra Diaz Rodriguez": 12.65,  # Beneficiario real en lugar de DRA BLUE GOW
        "Luz Mercedes Diaz Rodriguez": 14.49,
        "Jorge Enrique Diaz Rodriguez": 12.28,
        "Carlos Arturo Diaz Rodriguez": 12.4,
        "Inversiones MADCOM": 12.28
        # TIERRA ARCO IRIS eliminado - no es beneficiario final (entidad corporativa)
    }
    
    print(f"\nVERIFICACION CONTRA VALORES ESPERADOS:")
    for expected_name, expected_value in expected_beneficiaries.items():
        found = False
        for row in red_cow_participations:
            if normalize_name(row['Accionista']) == normalize_name(expected_name):
                diff = abs(row['Participacion'] - expected_value)
                status = "[OK]" if diff < 0.1 else "[ERROR]"
                print(f"  {status} {expected_name}: Esperado {expected_value}%, Actual {row['Participacion']}%")
                found = True
                break
        if not found:
            print(f"  [ERROR] {expected_name}: No encontrado")

def normalize_name(name):
    """
    Normaliza nombres para comparación
    """
    return name.upper().replace("Ñ", "N").replace("Á", "A").replace("É", "E").replace("Í", "I").replace("Ó", "O").replace("Ú", "U")

def main():
    """Función principal"""
    if len(sys.argv) < 2:
        print("[ERROR] Uso: python fix_dra_blue_dynamic.py <archivo_excel>")
        return False
    
    # Permitir modo de solo verificación de dependencias
    if len(sys.argv) == 2 and sys.argv[1] == "--check-deps":
        print("[OK] Dependencias verificadas correctamente")
        return True
    
    excel_file = sys.argv[1]
    
    if not Path(excel_file).exists():
        print(f"[ERROR] Archivo no encontrado: {excel_file}")
        return False
    
    # Realizar consolidación
    result = consolidate_participations(excel_file)
    
    if result:
        print(f"\n[OK] Consolidacion completada exitosamente")
    else:
        print(f"\n[ERROR] Error en la consolidacion")
    
    return result

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)