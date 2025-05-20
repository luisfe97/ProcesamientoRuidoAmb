import os
import warnings
import pandas as pd
from data.constants import SHEETS_TO_PROCESS, ARCHIVO_EXCEL, OUTPUT_FOLDER
from utils.file_utils import combine_excel_files
from processing.acoustic import aplicar_Correccion
from processing.meteorology import process_and_export_weather_data
from processing.data_handler import (
    cargar_datos, procesar_tercios_octava, crear_tabla_procesada, 
    filtrar_por_periodos, procesar_diario
)
from processing.statistics import generar_resumenes, calcular_estadisticos, actualizar_resumen, calcular_L_Raseq_dn
from processing.uncertainty_handler import calcular_incertidumbres
from processing.compliance import (
    asignar_limites, asignar_limites_diarios, procesar_compliance_diurno, 
    procesar_compliance_nocturno, finalizar_agrupados
)
from export.excel import export_to_template
from export.ruido_total import (procesar_excel_simple,combinar_excels)

# Configuración inicial
warnings.filterwarnings('ignore')
pd.options.display.float_format = '{:.2f}'.format

# Aseguramos que la carpeta de salida exista
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def procesar_hoja(sheet, pto, archivo_excel=ARCHIVO_EXCEL, file_path=ARCHIVO_EXCEL):
    """
    Procesa una hoja específica del archivo Excel
    
    Args:
        sheet: Nombre de la hoja a procesar
        pto: Número de punto para el archivo de salida
        archivo_excel: Nombre del archivo Excel
        file_path: Ruta del archivo Excel
        
    Returns:
        Número de punto actualizado
    """
    # 1. Cargar datos
    dataframes, nombres, TerciosOctava, dfASlow, dfAImpulse, Estacion = cargar_datos(archivo_excel, sheet)
    
    # 2. Procesar datos meteorológicos
    MET_resultado, MET_Diurno, MET_Nocturno, resumen, MET_resumen_diurno, MET_resumen_nocturno = process_and_export_weather_data(file_path, Estacion, pto)
    
    # 3. Procesar tercios de octava
    TerciosOctava, DfAjusteTonal = procesar_tercios_octava(TerciosOctava)
    
    # 4. Crear tabla procesada
    TablaProcesada = crear_tabla_procesada(TerciosOctava, dfASlow, dfAImpulse, DfAjusteTonal)
    TablaProcesada['LRASeq,i'] = TablaProcesada.apply(aplicar_Correccion, axis=1)
    
    # 5. Filtrar por precipitación
    if MET_resultado is not None:
        TablaProcesada = TablaProcesada[~TablaProcesada['Period start'].isin(
            MET_resultado[MET_resultado['PREC'] > 0.5].index
        )]
    
    # 6. Filtrar por períodos
    diurno_ref, nocturno_ref, diurno_Total, nocturno_Total = filtrar_por_periodos(TerciosOctava, TablaProcesada)
    
    # 7. Procesar datos diarios
    DfAjusteTonal_diurno_ref, DfAjusteTonal_nocturno_ref, diurno_grouped, nocturno_grouped = procesar_diario(TablaProcesada, diurno_ref, nocturno_ref)
    
    # 8. Finalizar agrupados
    diurno_grouped, nocturno_grouped = finalizar_agrupados(diurno_grouped, nocturno_grouped, DfAjusteTonal_diurno_ref, DfAjusteTonal_nocturno_ref)
    
    # 9. Generar resumenes
    resumen_diurno, resumen_nocturno = generar_resumenes(diurno_Total, nocturno_Total, diurno_grouped, nocturno_grouped)
    
    # 10. Calcular estadísticos
    resultados_diurnos_df = calcular_estadisticos(resumen_diurno, diurno_Total)
    resultados_nocturnos_df = calcular_estadisticos(resumen_nocturno, nocturno_Total)
    
    # 11. Actualizar resumenes
    resumen_diurno = actualizar_resumen(resumen_diurno, resultados_diurnos_df)
    resumen_nocturno = actualizar_resumen(resumen_nocturno, resultados_nocturnos_df)
    
    # 12. Calcular día-noche
    dia_noche = pd.DataFrame({
        'TipoDia': resumen_diurno['TipoDia'],
        'Nm,dn': resumen_diurno['Conteo'] + resumen_nocturno['Conteo'],
        'LASeq': calcular_L_Raseq_dn(resumen_diurno['LASeq_k'], resumen_nocturno['LASeq_k']),
        'LRASeq': calcular_L_Raseq_dn(resumen_diurno['LRASeq_k'], resumen_nocturno['LRASeq_k']),
        'LAIeq': calcular_L_Raseq_dn(resumen_diurno['LAIeq_k'], resumen_nocturno['LAIeq_k'])
    })
    
    # 13. Calcular incertidumbres
    IncExp_diu, IncExp_noc = calcular_incertidumbres(resumen_diurno, resumen_nocturno, MET_resumen_diurno, MET_resumen_nocturno)
    
    # 14. Asignar límites
    resumen_diurno, resumen_nocturno = asignar_limites(resumen_diurno, resumen_nocturno, Estacion)
    diurno_grouped, nocturno_grouped = asignar_limites_diarios(diurno_grouped, nocturno_grouped, Estacion)
    
    # 15. Procesar compliance
    resumen_diurno, diurno_grouped = procesar_compliance_diurno(resumen_diurno, diurno_grouped, IncExp_diu)
    resumen_nocturno, nocturno_grouped = procesar_compliance_nocturno(resumen_nocturno, nocturno_grouped, IncExp_noc)
    
    # 16. Exportar resultados
    template_path = "Plantilla/Plantilla_Macro.xlsx"
    output_path = f'{OUTPUT_FOLDER}/PTO{pto}.xlsx'
    TablaProcesada=TablaProcesada.drop(columns=['Fechas'])
    # Eliminar filas completamente nulas de cada DataFrame
    print(diurno_grouped)
    export_to_template(
        TablaProcesada, 
        diurno_grouped, 
        nocturno_grouped, 
        resumen_diurno, 
        resumen_nocturno, 
        dia_noche, 
        template_path, 
        output_path,
        Estacion
    )
    
    return pto + 1

def main():
    """Función principal que ejecuta el flujo completo de procesamiento"""
    
    archivo_excel = ARCHIVO_EXCEL
    file_path = ARCHIVO_EXCEL
    pto = 1
    
    # Procesar todas las hojas
    for sheet in SHEETS_TO_PROCESS:
        print(f"Procesando hoja: {sheet}")
        pto = procesar_hoja(sheet, pto, archivo_excel, file_path)
    
    # Combinar archivos Excel
    combine_excel_files(OUTPUT_FOLDER)

    """Función principal que ejecuta el procesamiento completo"""
    # Configuración de rutas
    ruta_excel = "PTOS_salida/Excel_Intercalado.xlsx"    
    # Procesar el archivo Excel
    dataframes = procesar_excel_simple(ruta_excel, OUTPUT_FOLDER)
    
    # Combinar excels si se requiere
    archivo1 = os.path.join(OUTPUT_FOLDER, "RUIDO TOTAL.xlsx")
    archivo2 = ruta_excel
    archivo_salida = os.path.join(OUTPUT_FOLDER, "20240723 FOM305-25 y 26 Plantilla Ruido Total (Ambiental) v1.xlsx")
    
    combinar_excels(archivo1, archivo2, archivo_salida)
    print("Proceso completado con éxito.")

if __name__ == "__main__":
    main()
