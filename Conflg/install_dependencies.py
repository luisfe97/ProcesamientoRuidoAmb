import subprocess
import sys
import os
import platform

def install_dependencies():
    """
    Instala todas las dependencias necesarias para el proyecto.
    Funciona en Windows, macOS y Linux.
    """
    # Lista de paquetes necesarios
    required_packages = [
        'pandas',
        'numpy',
        'scipy',
        'openpyxl',
        'xlsxwriter',
        'matplotlib',
        'PyQt5',  # Útil para visualizaciones si se necesitan
    ]
    
    # Verificar pip
    print("Verificando instalación de pip...")
    try:
        subprocess.check_call([sys.executable, '-m', 'pip', '--version'])
    except subprocess.CalledProcessError:
        print("❌ Error: pip no está instalado o no está funcionando correctamente.")
        print("Por favor, instala pip manualmente y vuelve a ejecutar este script.")
        sys.exit(1)
    
    # Actualizar pip
    print("Actualizando pip...")
    try:
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', '--upgrade', 'pip'])
    except subprocess.CalledProcessError:
        print("⚠️ Advertencia: No se pudo actualizar pip, pero se continuará con la instalación.")
    
    # Instalar cada paquete
    print("\n📦 Instalando dependencias...")
    
    for package in required_packages:
        print(f"  Instalando {package}...")
        try:
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])
            print(f"  ✅ {package} instalado correctamente.")
        except subprocess.CalledProcessError:
            print(f"  ❌ Error al instalar {package}. Por favor, instálalo manualmente.")
    
    # Verificar que todo se haya instalado correctamente
    print("\n🔍 Verificando instalaciones...")
    all_installed = True
    
    for package in required_packages:
        try:
            __import__(package)
            print(f"  ✅ {package} verificado.")
        except ImportError:
            print(f"  ❌ {package} no se pudo importar. Es posible que la instalación haya fallado.")
            all_installed = False
    
    # Mensaje final
    if all_installed:
        print("\n✅ ¡Todas las dependencias se han instalado correctamente!")
    else:
        print("\n⚠️ Algunas dependencias no se instalaron correctamente. Por favor, revisa los errores e intenta instalarlas manualmente.")
    
    # Información de sistema
    print("\n📋 Información del sistema:")
    print(f"  - Python: {sys.version}")
    print(f"  - Sistema operativo: {platform.system()} {platform.release()}")
    print(f"  - Directorio de trabajo: {os.getcwd()}")

if __name__ == "__main__":
    print("🚀 Iniciando instalación de dependencias para el Sistema de Procesamiento Acústico...")
    install_dependencies()