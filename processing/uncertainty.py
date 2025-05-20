import numpy as np
import pandas as pd

def Calculo_U(df):
    """
    Calcula la incertidumbre a partir del dataframe proporcionado
    
    Args:
        df: DataFrame con las columnas LRASeq y s
        
    Returns:
        Valor de incertidumbre calculado
    """
    # Convertir las columnas necesarias a números
    df['LRASeq'] = pd.to_numeric(df['LRASeq'], errors='coerce')
    df['s'] = pd.to_numeric(df['s'], errors='coerce')
    
    # Realizar las operaciones
    parte1 = np.power(10, 0.1 * df['LRASeq']) + df['s']
    
    resultado = 10 * np.log10(parte1) - df['LRASeq']
    
    return resultado

def formatear_valores(df):
    """
    Formatea valores mayores o iguales a 1000 en las columnas del DataFrame.
    
    Args:
        df: DataFrame a formatear
        
    Returns:
        DataFrame con valores formateados
    """
    for col in df.columns:
        df[col] = df[col].apply(lambda x: f"{x:.2E}" if x >= 1000 else x)
    return df

def crear_dataframe(tipo_dia, resumen_diurno):
    """
    Crea un DataFrame para el cálculo de incertidumbre de un tipo de día específico
    
    Args:
        tipo_dia: Tipo de día (Dominical, Ordinario, Total)
        resumen_diurno: DataFrame con el resumen diurno
        
    Returns:
        DataFrame formateado para los cálculos
    """
    df = pd.DataFrame()
    df['Nm'] = resumen_diurno[resumen_diurno['TipoDia'] == tipo_dia]['Conteo']
    df['LRASeq'] = resumen_diurno[resumen_diurno['TipoDia'] == tipo_dia]['LRASeq_k']
    df['s'] = resumen_diurno[resumen_diurno['TipoDia'] == tipo_dia]['s_k']
    df['u'] = Calculo_U(df)
    return formatear_valores(df)

def calcular_incertidumbre(ubi, sens, instru, tipo_a):
    """
    Calcula la incertidumbre combinada.
    
    Args:
        ubi: DataFrame con incertidumbre de ubicación
        sens: DataFrame con incertidumbre de sensibilidad
        instru: DataFrame con incertidumbre instrumental
        tipo_a: DataFrame con incertidumbre tipo A
        
    Returns:
        Incertidumbre combinada
    """
    datos = [
        ubi['uloc'].iloc[0],
        sens['umic,H*'].iloc[0], sens['umic,P'].iloc[0], sens['umic,T'].iloc[0],
        instru['uslm'].iloc[0], instru['uresol'].iloc[0],
        tipo_a['u'].iloc[0]
    ]
    return np.sqrt(sum(np.square(datos)))

def calcular_veff(IncComb, Tipo_A, instru, sens, Ubi):
    """
    Calcula los grados de libertad efectivos veff
    
    Args:
        IncComb: Incertidumbre combinada
        Tipo_A: DataFrame con incertidumbre tipo A
        instru: DataFrame con incertidumbre instrumental
        sens: DataFrame con incertidumbre de sensibilidad
        Ubi: DataFrame con incertidumbre de ubicación
        
    Returns:
        Grados de libertad efectivos
    """
    inc_comb = IncComb.iloc[0]
    tipo_u = Tipo_A['u'].iloc[0]
    tipo_nm = Tipo_A['Nm'].iloc[0]
    instru_uslm = instru['uslm'].iloc[0]
    sens_umic_T = sens['umic,T'].iloc[0]
    sens_umic_P = sens['umic,P'].iloc[0]
    sens_umic_H = sens['umic,H*'].iloc[0]
    ubi_uloc = Ubi['uloc'].iloc[0]

    veff = np.round(
        np.power(inc_comb, 4) /
        (
            np.power(tipo_u, 4) / (tipo_nm - 1) +
            np.power(instru_uslm, 4) / 50 +
            np.power(sens_umic_T, 4) / 200 +
            np.power(sens_umic_P, 4) / 200 +
            np.power(sens_umic_H, 4) / 200 +
            np.power(ubi_uloc, 4) / 50
        ), 0
    )
    return veff