import os

def create_project_structure():
    """
    Crea la estructura de carpetas necesaria para el proyecto.
    """
    # Definir la estructura de carpetas
    structure = {
        'utils': ['__init__.py', 'date_utils.py', 'file_utils.py'],
        'processing': [
            '__init__.py', 
            'acoustic.py', 
            'meteorology.py', 
            'statistics.py', 
            'uncertainty.py', 
            'uncertainty_handler.py',
            'data_handler.py',
            'compliance.py'
        ],
        'data': ['__init__.py', 'constants.py', 'limits.py'],
        'export': ['__init__.py', 'excel.py']
    }
    
    # Crear carpetas y archivos
    for folder, files in structure.items():
        os.makedirs(folder, exist_ok=True)
        for file in files:
            with open(os.path.join(folder, file), 'w') as f:
                f.write('# Módulo generado automáticamente\n')
    
    # Crear carpetas adicionales
    os.makedirs('PTOS_salida', exist_ok=True)
    
    # Crear archivos principales
    with open('main.py', 'w') as f:
        f.write('# Script principal\n')
    
    # Crear README.md con menos contenido para evitar problemas
    with open('README.md', 'w') as f:
        f.write('# Sistema de Procesamiento Acústico\n\nEstructura inicial del proyecto generada con setup.py')
    
    print("✅ Estructura de carpetas creada exitosamente.")

if __name__ == "__main__":
    create_project_structure()