import pandas as pd
import numpy as np
from data.constants import HORAS_REFERENCIA
from utils.date_utils import corregir_fecha_hora
from processing.acoustic import Ponderacion_A, ajuste_tonal, calcular_ki

def cargar_datos(archivo_excel, sheet):
    """
    Carga los datos del archivo Excel para una hoja específica
    
    Args:
        archivo_excel: Ruta del archivo Excel
        sheet: Nombre de la hoja a procesar
        
    Returns:
        Tupla con (dataframes, nombres, TerciosOctava, dfASlow, dfAImpulse, Estacion)
    """
    df = pd.read_excel(archivo_excel, sheet_name=sheet, header=None)
    Nombres = df.iloc[6, :]
    Estacion = df.iloc[4, 1]

    # Preparar datos
    Nombres = Nombres.dropna().reset_index(drop=True)
    df.columns = df.iloc[8, :]
    df = df.iloc[9:].reset_index(drop=True)
    primera_columna = df.iloc[:, 0].rename("Period start")  
    
    # Limpiar nombres para evitar NaN
    Nombres = [str(n).replace("1/3 Oct", "").replace("Hz", "").strip() if pd.notna(n) else f"Desconocido_{idx}" 
            for idx, n in enumerate(Nombres)]

    # Extraer dataframes por grupos de columnas
    dataframes = []
    nombres = []
    j = 1  # Índice independiente para recorrer `Nombres`
    for i in range(1, df.shape[1], 5):
        sub_df = pd.concat([primera_columna, df.iloc[:, i:i+5]], axis=1)
        
        # Verificar si `j` está dentro del rango de `Nombres`
        if j < len(Nombres):
            nombre = Nombres[j]  # Usar índice independiente
        else:
            nombre = f"Desconocido_{j}"
        
        dataframes.append(sub_df)
        nombres.append(nombre)
        
        j += 1  # Avanzamos al siguiente nombre sin saltos
        
    # Crear TerciosOctava de manera más directa
    TerciosOctava = pd.DataFrame({
        nombre: dataframes[idx+2].iloc[:, 1]
        for idx, nombre in enumerate(nombres[2:])
    })
    TerciosOctava.insert(0, 'Period start', primera_columna.reset_index(drop=True))

    # Guardar datasets específicos
    dfASlow = dataframes[0]
    dfAImpulse = dataframes[1]

    # Convertir y corregir fechas de una manera más eficiente
    for df_item in [TerciosOctava, dfASlow, dfAImpulse]:
        df_item['Period start'] = pd.to_datetime(df_item['Period start'], format='%d/%m/%Y %I:%M:%S %p')
        df_item['Period start'] = df_item['Period start'].apply(corregir_fecha_hora)
    
    return dataframes, nombres, TerciosOctava, dfASlow, dfAImpulse, Estacion

def procesar_tercios_octava(TerciosOctava):
    """
    Procesa los datos de tercios de octava con ponderación A
    
    Args:
        TerciosOctava: DataFrame con datos de tercios de octava
        
    Returns:
        DataFrame con tercios de octava procesados y datos de ajuste tonal
    """
    # Optimizar ponderación y ajuste tonal
    # Convertir spectrum a lista de strings para evitar problemas con índices
    spectrum_list = list(TerciosOctava.columns[1:])

    # Usar list comprehension y procesamiento vectorial
    resultados_Ponderados = np.array([
        [Ponderacion_A(freq, level) for freq, level in zip(spectrum_list, row)]
        for row in TerciosOctava.iloc[:, 1:].values
    ])

    # Crear DataFrame de resultados ponderados
    resultados_df = pd.DataFrame(resultados_Ponderados, columns=spectrum_list)
    resultados_df.insert(0, 'Period start', TerciosOctava['Period start'].reset_index(drop=True))
    TerciosOctava_procesado = resultados_df

    # Ajuste tonal optimizado con lista de frecuencias
    resultados = [
        {
            'KT,i': ajuste_tonal(spectrum_list, row)[0],
            'Bandas': ajuste_tonal(spectrum_list, row)[1]
        }
        for row in TerciosOctava_procesado.iloc[:, 1:].values
    ]
    DfAjusteTonal = pd.DataFrame(resultados)
    
    return TerciosOctava_procesado, DfAjusteTonal

def crear_tabla_procesada(TerciosOctava, dfASlow, dfAImpulse, DfAjusteTonal):
    """
    Crea la tabla procesada con todos los datos combinados
    
    Args:
        TerciosOctava: DataFrame con datos de tercios de octava
        dfASlow: DataFrame con datos de nivel lento
        dfAImpulse: DataFrame con datos de nivel impulso
        DfAjusteTonal: DataFrame con datos de ajuste tonal
        
    Returns:
        DataFrame con la tabla procesada
    """
    # Conversión y cálculo de KI más eficiente
    dfASlow['Leq'] = dfASlow['Leq'].astype(float)
    dfAImpulse['Leq'] = dfAImpulse['Leq'].astype(float)

    KI = pd.DataFrame({'KI,i': dfAImpulse['Leq'] - dfASlow['Leq']})
    KI['KI,i'] = KI['KI,i'].apply(calcular_ki)

    # Crear tabla procesada de manera más directa
    TablaProcesada = pd.DataFrame({
        'Period start': TerciosOctava['Period start'],
        'LASeq,i': dfASlow['Leq'],
        'LAIeq,i': dfAImpulse['Leq'],
        'KI,i': KI['KI,i'],
        'KT,i': DfAjusteTonal['KT,i'],
        'Bandas': DfAjusteTonal['Bandas']
    })

    return TablaProcesada

def filtrar_por_periodos(TerciosOctava, TablaProcesada):
    """
    Filtra los datos por períodos diurnos y nocturnos
    
    Args:
        TerciosOctava: DataFrame con datos de tercios de octava
        TablaProcesada: DataFrame con datos procesados
        
    Returns:
        Tupla con (diurno_ref, nocturno_ref, diurno_Total, nocturno_Total)
    """
    # Definir horas de referencia
    hora_diurna_inicio, hora_diurna_fin = HORAS_REFERENCIA["diurna_inicio"].time(), HORAS_REFERENCIA["diurna_fin"].time()
    hora_nocturna_inicio, hora_nocturna_fin = HORAS_REFERENCIA["nocturna_inicio"].time(), HORAS_REFERENCIA["nocturna_fin"].time()
    
    # Filtrar tercios de octava por período
    diurno_ref = TerciosOctava[(TerciosOctava['Period start'].dt.time >= hora_diurna_inicio) &
                         (TerciosOctava['Period start'].dt.time <= hora_diurna_fin)]
    nocturno_ref = TerciosOctava[(TerciosOctava['Period start'].dt.time >= hora_nocturna_inicio) |
                           (TerciosOctava['Period start'].dt.time <= hora_nocturna_fin)]
    
    # Filtrar datos diurnos y nocturnos para la tabla procesada
    diurno_Total = TablaProcesada[
        (TablaProcesada['Period start'].dt.time >= hora_diurna_inicio) &
        (TablaProcesada['Period start'].dt.time <= hora_diurna_fin) &
        (TablaProcesada['LASeq,i'].notna()) &
        (TablaProcesada['LAIeq,i'].notna()) &
        (TablaProcesada['LRASeq,i'].notna())
    ]

    nocturno_Total = TablaProcesada[
        ((TablaProcesada['Period start'].dt.time >= hora_nocturna_inicio) |
        (TablaProcesada['Period start'].dt.time <= hora_nocturna_fin)) &
        (TablaProcesada['LASeq,i'].notna()) &
        (TablaProcesada['LAIeq,i'].notna()) &
        (TablaProcesada['LRASeq,i'].notna())
    ]
    
    # Añadir tipo de día
    for df in [diurno_Total, nocturno_Total]:
        df['TipoDia'] = np.where(df['Period start'].dt.dayofweek == 6, 'Dominical', 'Ordinario')
    
    return diurno_ref, nocturno_ref, diurno_Total, nocturno_Total

def procesar_diario(TablaProcesada, diurno_ref, nocturno_ref):
    """
    Procesa los datos diarios para períodos diurnos y nocturnos
    
    Args:
        TablaProcesada: DataFrame con datos procesados
        diurno_ref: DataFrame con referencia diurna
        nocturno_ref: DataFrame con referencia nocturna
        
    Returns:
        Tupla con (DfAjusteTonal_diurno_ref, DfAjusteTonal_nocturno_ref, diurno_grouped, nocturno_grouped)
    """
    # Convertir 'Period start' a datetime y crear columnas de fecha
    for df in [TablaProcesada, diurno_ref, nocturno_ref]:
        if 'Fechas' not in df.columns:
            df['Fechas'] = df['Period start'].dt.date

    # Identificar columnas de ruido (excluyendo 'Period start' y 'Fechas')
    columnas_ruido_ref = [col for col in diurno_ref.columns if col not in ['Period start', 'Fechas']]
    
    # Asegurar que todas las columnas de ruido sean numéricas
    for df in [diurno_ref, nocturno_ref]:
        df[columnas_ruido_ref] = df[columnas_ruido_ref].apply(pd.to_numeric, errors='coerce')

    from processing.statistics import promedio_logaritmico_ref
    
    # Agrupar por fecha y calcular métricas
    diurno_grouped_ref = diurno_ref.groupby('Fechas').agg(
        {col: lambda x: promedio_logaritmico_ref(x.dropna()) for col in columnas_ruido_ref}
    ).reset_index()

    nocturno_grouped_ref = nocturno_ref.groupby('Fechas').agg(
        {col: lambda x: promedio_logaritmico_ref(x.dropna()) for col in columnas_ruido_ref}
    ).reset_index()

    # Procesar datos diurnos y nocturnos con list comprehension
    resultados_diurno_ref = [{
        'KT,i': ajuste_tonal(
            diurno_grouped_ref.columns[1:].to_list(), 
            diurno_grouped_ref.iloc[idx, 1:].to_list()
        )[0],
        'Bandas': ajuste_tonal(
            diurno_grouped_ref.columns[1:].to_list(), 
            diurno_grouped_ref.iloc[idx, 1:].to_list()
        )[1]
    } for idx in range(len(diurno_grouped_ref))]

    resultados_nocturno_ref = [{
        'KT,i': ajuste_tonal(
            nocturno_grouped_ref.columns[1:].to_list(), 
            nocturno_grouped_ref.iloc[idx, 1:].to_list()
        )[0],
        'Bandas': ajuste_tonal(
            nocturno_grouped_ref.columns[1:].to_list(), 
            nocturno_grouped_ref.iloc[idx, 1:].to_list()
        )[1]
    } for idx in range(len(nocturno_grouped_ref))]

    # Crear los DataFrames finales
    DfAjusteTonal_diurno_ref = pd.DataFrame(resultados_diurno_ref)
    DfAjusteTonal_nocturno_ref = pd.DataFrame(resultados_nocturno_ref)
    
    # Definir horas de referencia
    hora_diurna_inicio, hora_diurna_fin = HORAS_REFERENCIA["diurna_inicio"].time(), HORAS_REFERENCIA["diurna_fin"].time()
    hora_nocturna_inicio, hora_nocturna_fin = HORAS_REFERENCIA["nocturna_inicio"].time(), HORAS_REFERENCIA["nocturna_fin"].time()

    # Filtrar horarios diurno y nocturno
    diurno = TablaProcesada[
        (TablaProcesada['Period start'].dt.time >= hora_diurna_inicio) &
        (TablaProcesada['Period start'].dt.time <= hora_diurna_fin) &
        (TablaProcesada['LASeq,i'].notna()) &
        (TablaProcesada['LAIeq,i'].notna()) & 
        (TablaProcesada['LRASeq,i'].notna())
    ]

    nocturno = TablaProcesada[
        ((TablaProcesada['Period start'].dt.time >= hora_nocturna_inicio) |
        (TablaProcesada['Period start'].dt.time <= hora_nocturna_fin)) &
        (TablaProcesada['LASeq,i'].notna()) &
        (TablaProcesada['LAIeq,i'].notna()) & 
        (TablaProcesada['LRASeq,i'].notna())
    ]

    # Crear columna de fechas para agrupar
    diurno['Fechas'] = diurno['Period start'].dt.date
    nocturno['Fechas'] = nocturno['Period start'].dt.date

    # Agrupar por fecha y calcular métricas para diurno y nocturno
    diurno_grouped = diurno.groupby('Fechas').agg(
        Nm_1d=('LASeq,i', 'size'),
        LASeq_1d=('LASeq,i', lambda x: promedio_logaritmico_ref(x)),
        LAIeq_1d=('LAIeq,i', lambda x: promedio_logaritmico_ref(x)),
        LRASeq_1d=('LRASeq,i', lambda x: promedio_logaritmico_ref(x))
    ).reset_index()

    nocturno_grouped = nocturno.groupby('Fechas').agg(
        Nm_1d=('LASeq,i', 'size'),
        LASeq_1d=('LASeq,i', lambda x: promedio_logaritmico_ref(x)),
        LAIeq_1d=('LAIeq,i', lambda x: promedio_logaritmico_ref(x)),
        LRASeq_1d=('LRASeq,i', lambda x: promedio_logaritmico_ref(x))
    ).reset_index()

    # Calcular KI,1d
    diurno_grouped['KI,1d'] = (diurno_grouped['LAIeq_1d'] - diurno_grouped['LASeq_1d']).apply(calcular_ki)
    nocturno_grouped['KI,1d'] = (nocturno_grouped['LAIeq_1d'] - nocturno_grouped['LASeq_1d']).apply(calcular_ki)
    
    return DfAjusteTonal_diurno_ref, DfAjusteTonal_nocturno_ref, diurno_grouped, nocturno_grouped