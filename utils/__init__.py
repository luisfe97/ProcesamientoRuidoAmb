import numpy as np
import pandas as pd
from processing.acoustic import calcular_L_Raseq_dn

def promedio_logaritmico_ref(grupo):
    """
    Calcula el promedio logarítmico de un grupo de valores en dB
    
    Args:
        grupo: Serie o lista de valores en dB
        
    Returns:
        Promedio logarítmico redondeado a 1 decimal
    """
    valores_lineales_ref = 10 ** (grupo / 10)  # Convertir dB a escala lineal
    promedio_lineal_ref = np.mean(valores_lineales_ref)  # Promedio en escala lineal
    return round(10 * np.log10(promedio_lineal_ref), 1)  # Convertir de vuelta a dB y redondear

def NivelEq_Jornadas(df):
    """
    Calcula el nivel equivalente para jornadas ordinarias y dominicales
    
    Args:
        df: DataFrame con datos de niveles sonoros
        
    Returns:
        Lista con niveles equivalentes [ordinarios, dominicales]
    """
    # Filtrar solo los días ordinarios y seleccionar las columnas que necesitas     
    df_ordinarios = df[df['TipoDia'] == 'Ordinario']      
    df_ordinarios['LRASeq_1d'] = pd.to_numeric(df_ordinarios['LRASeq_1d'])     
    df_ordinarios['Nm_1d'] = pd.to_numeric(df_ordinarios['Nm_1d'])          
    
    # Calcular N_m,ord,1d * 3600 para cada día     
    N_m_tiempo_ord = df_ordinarios['Nm_1d'] * 3600          
    
    # Calcular la primera parte de la suma dentro del logaritmo     
    primera_parte_ord = np.sum(10 ** (0.1 * (df_ordinarios['LRASeq_1d'] + (10 * np.log10(N_m_tiempo_ord)))))          
    
    # Calcular la segunda parte (denominador)     
    segunda_parte_ord = np.sum(N_m_tiempo_ord)          
    
    # Calcular el resultado final     
    resultado_ordinarios = (10 * np.log10(primera_parte_ord)) - (10 * np.log10(segunda_parte_ord))       
    
    df_dominical = df[df['TipoDia'] == 'Dominical']     
    df_dominical['LRASeq_1d'] = pd.to_numeric(df_dominical['LRASeq_1d'])     
    df_dominical['Nm_1d'] = pd.to_numeric(df_dominical['Nm_1d'])          
    
    # Calcular N_m,dom,1d * 3600 para cada día     
    N_m_tiempo_dom = df_dominical['Nm_1d'] * 3600          
    
    # Calcular la primera parte de la suma dentro del logaritmo
    primera_parte_dom = np.sum(10 ** (0.1 * (df_dominical['LRASeq_1d'] + (10 * np.log10(N_m_tiempo_dom)))))
    
    # Calcular la segunda parte (denominador)     
    segunda_parte_dom = np.sum(N_m_tiempo_dom)          
    
    # Calcular el resultado final
    resultado_dominicales = (10 * np.log10(primera_parte_dom)) - (10 * np.log10(segunda_parte_dom))          
    
    return [resultado_ordinarios, resultado_dominicales]

def NivelEq_Jornadas_totales(df):
    """
    Calcula el nivel equivalente para todas las jornadas
    
    Args:
        df: DataFrame con datos de niveles sonoros
        
    Returns:
        Lista con nivel equivalente total
    """
    df['LRASeq_1d'] = pd.to_numeric(df['LRASeq_1d'])     
    df['Nm_1d'] = pd.to_numeric(df['Nm_1d'])          
    
    # Calcular N_m,dom,1d * 3600 para cada día     
    N_m_tiempo= df['Nm_1d'] * 3600          
    
    # Calcular la primera parte de la suma dentro del logaritmo
    primera_parte= np.sum(10 ** (0.1 * (df['LRASeq_1d'] + (10 * np.log10(N_m_tiempo)))))
    
    # Calcular la segunda parte (denominador)     
    segunda_parte= np.sum(N_m_tiempo)          
    
    # Calcular el resultado final
    resultado = (10 * np.log10(primera_parte)) - (10 * np.log10(segunda_parte))    
    
    return [resultado]

def procesar_resumen_periodo(df_total, df_grouped):
    """
    Procesa los datos para un período específico (diurno/nocturno) y genera un resumen.
    
    Args:
        df_total: DataFrame con todos los datos del período
        df_grouped: DataFrame agrupado del período
    
    Returns:
        DataFrame con el resumen completo
    """
    ordinarios, dominical = NivelEq_Jornadas(df_grouped)
    eq_total = NivelEq_Jornadas_totales(df_grouped)
    
    leq_jornadas = pd.DataFrame({
        'TipoDia': ['Dominical', 'Ordinario'],
        'LRASeq_k': [dominical, ordinarios]
    })
    
    resumen = df_total.groupby('TipoDia').agg(
        Conteo=('TipoDia', 'size'),
        LASeq_k=('LASeq,i', promedio_logaritmico_ref),
        LAIeq_k=('LAIeq,i', promedio_logaritmico_ref)
    ).reset_index()
    
    resumen = pd.merge(resumen, leq_jornadas, on='TipoDia', how='outer')
    
    total_row = pd.DataFrame({
        'TipoDia': ['Total'],
        'Conteo': [df_total.shape[0]],
        'LASeq_k': [promedio_logaritmico_ref(df_total['LASeq,i'])],
        'LAIeq_k': [promedio_logaritmico_ref(df_total['LAIeq,i'])],
        'LRASeq_k': [round(eq_total[0], 1)]
    })
    
    resumen_final = pd.concat([resumen, total_row], ignore_index=True)
    columnas_ordenadas = ['TipoDia', 'Conteo', 'LASeq_k', 'LAIeq_k', 'LRASeq_k']
    
    return resumen_final[columnas_ordenadas]

def generar_resumenes(diurno_Total, nocturno_Total, diurno_grouped, nocturno_grouped):
    """
    Genera los resúmenes para períodos diurnos y nocturnos.
    
    Args:
        diurno_Total: DataFrame con todos los datos diurnos
        nocturno_Total: DataFrame con todos los datos nocturnos
        diurno_grouped: DataFrame agrupado de datos diurnos
        nocturno_grouped: DataFrame agrupado de datos nocturnos
        
    Returns:
        Tupla con (resumen_diurno, resumen_nocturno)
    """
    resumen_diurno = procesar_resumen_periodo(diurno_Total, diurno_grouped)
    resumen_nocturno = procesar_resumen_periodo(nocturno_Total, nocturno_grouped)
    
    return resumen_diurno, resumen_nocturno

def calcular_estadisticos(resumen, datos_total):
    """
    Calcula estadísticos para el resumen
    
    Args:
        resumen: DataFrame con resumen
        datos_total: DataFrame con datos totales
        
    Returns:
        DataFrame con estadísticos calculados
    """
    resultados = []
    
    for index, row in resumen.iterrows():
        tipo_dia = row['TipoDia']
        N_mk = row['Conteo']
        LRASeq_k = row['LRASeq_k']
        
        # Filtrar datos
        data_filtrada = datos_total if tipo_dia == 'Total' else datos_total[datos_total['TipoDia'] == tipo_dia]
        
        # Calcular LRASeq,i usando promedio logarítmico
        LRASeq_i = promedio_logaritmico_ref(data_filtrada['LRASeq,i'])
        
        # Calcular s_k^2
        diferencias = (10**(0.1 * data_filtrada['LRASeq,i']) - 10**(0.1 * LRASeq_k))**2
        s_k2 = diferencias.sum() / (N_mk - 1)
        
        # Calcular s_k
        s_k = (s_k2**0.5) / (N_mk**0.5)
        
        # Formatear a notación científica con 2 decimales
        resultados.append({
            'TipoDia': tipo_dia, 
            's_k^2': format(s_k2, '.2e'), 
            's_k': format(s_k, '.2e')
        })
    
    return pd.DataFrame(resultados)

def actualizar_resumen(resumen, resultados_df):
    """
    Actualiza el resumen con los resultados calculados
    
    Args:
        resumen: DataFrame de resumen original
        resultados_df: DataFrame con resultados a incluir
        
    Returns:
        DataFrame de resumen actualizado
    """
    # Limpiar columnas existentes
    for col in ['s_k^2', 's_k']:
        if col in resumen.columns:
            resumen = resumen.drop(columns=[col])
    
    # Merge con resultados
    return pd.merge(resumen, resultados_df, on='TipoDia', how='outer')