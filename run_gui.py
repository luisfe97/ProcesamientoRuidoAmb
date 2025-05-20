#!/usr/bin/env python3
"""
Script para iniciar la interfaz gráfica del sistema de procesamiento acústico y meteorológico.
"""
import sys
import os

# Añadir directorio actual al path para encontrar todos los módulos
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

try:
    # Importar módulo principal de la GUI
    from gui.main_gui import main
    
    # Ejecutar la aplicación
    if __name__ == "__main__":
        main()
        
except ImportError as e:
    print(f"Error al importar los módulos necesarios: {e}")
    print("\nPor favor, asegúrate de haber instalado todas las dependencias:")
    print("python config/install_dependencies.py")
    
except Exception as e:
    print(f"Error al iniciar la aplicación: {e}")