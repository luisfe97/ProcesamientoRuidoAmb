# Sistema de Procesamiento Acústico

Este proyecto realiza el procesamiento y análisis de datos acústicos según la normativa colombiana.

## Estructura del Proyecto

El proyecto está organizado en una estructura modular:

```
proyecto/
│
├── main.py                      # Script principal que ejecuta todo el proceso
│
├── config/                      # Scripts de configuración
│   ├── setup.py                 # Script para crear estructura de carpetas
│   └── install_dependencies.py  # Script para instalar dependencias
│
├── input/                       # Carpeta para archivos de entrada
│   └── Met_Mar.xlsx             # Archivo Excel con datos de entrada
│
├── plantilla/                   # Carpeta para plantillas de Excel
│   └── Plantilla_Macro.xlsx     # Plantilla para exportación de resultados
│
├── utils/                       # Utilidades generales
│   ├── __init__.py
│   ├── date_utils.py            # Funciones para manejo de fechas
│   └── file_utils.py            # Funciones para manejo de archivos
│
├── processing/                  # Módulos de procesamiento
│   ├── __init__.py
│   ├── acoustic.py              # Funciones de procesamiento acústico
│   ├── meteorology.py           # Funciones de procesamiento meteorológico
│   ├── statistics.py            # Funciones estadísticas
│   ├── uncertainty.py           # Funciones base para cálculo de incertidumbre
│   ├── uncertainty_handler.py   # Gestión de cálculos de incertidumbre
│   ├── data_handler.py          # Funciones para carga y manejo de datos
│   └── compliance.py            # Funciones para evaluación de cumplimiento
│
├── data/                        # Datos compartidos
│   ├── __init__.py
│   ├── constants.py             # Constantes del programa
│   └── limits.py                # Límites regulatorios
│
├── export/                      # Funciones de exportación
│   ├── __init__.py
│   ├── excel.py                 # Funciones para exportar a Excel
│   └── ruido_total.py           # Script para consolidar resultados
│
└── PTOS_salida/                 # Carpeta donde se guardan los resultados
```

## Instalación

1. Clonar el repositorio
2. Instalar las dependencias:
   ```
   pip install pandas numpy scipy openpyxl xlsxwriter
   ```
3. Ejecutar `setup.py` para crear la estructura de carpetas:
   ```
   python setup.py
   ```

## Uso

1. Coloca el archivo `Met_Mar.xlsx` en la carpeta raíz del proyecto
2. Asegúrate de tener el archivo `Plantilla_Macro.xlsx` en la carpeta raíz
3. Ejecuta el script principal:
   ```
   python main.py
   ```
4. Los resultados se guardarán en la carpeta `PTOS_salida`

## Descripción de los Módulos

### utils

- `date_utils.py`: Funciones para el manejo y corrección de fechas y horas
- `file_utils.py`: Funciones para manejo de archivos Excel y combinación de resultados

### processing

- `acoustic.py`: Implementa las funciones de procesamiento acústico (ponderación A, ajuste tonal, etc.)
- `meteorology.py`: Funciones para procesar datos meteorológicos
- `statistics.py`: Funciones estadísticas para el cálculo de promedios logarítmicos y niveles equivalentes
- `uncertainty.py`: Cálculo de incertidumbres según la normativa

### data

- `constants.py`: Almacena constantes utilizadas en todo el proyecto
- `limits.py`: Define los límites regulatorios por tipo de zona

### export

- `excel.py`: Funciones para exportar resultados a archivos Excel con formato