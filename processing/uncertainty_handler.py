import pandas as pd
import numpy as np
from scipy.stats import t
from processing.uncertainty import crear_dataframe, calcular_incertidumbre, calcular_veff

def calcular_incertidumbres(resumen_diurno, resumen_nocturno, MET_resumen_diurno, MET_resumen_nocturno):
    """
    Calcula todas las incertidumbres para los datos diurnos y nocturnos
    
    Args:
        resumen_diurno: DataFrame con resumen diurno
        resumen_nocturno: DataFrame con resumen nocturno
        MET_resumen_diurno: DataFrame con resumen meteorológico diurno
        MET_resumen_nocturno: DataFrame con resumen meteorológico nocturno
        
    Returns:
        Tupla con (IncExp_diu, IncExp_noc)
    """
    # Creación de DataFrames para tipos A
    Tipo_A_Dom_Diurno = crear_dataframe('Dominical', resumen_diurno)
    Tipo_A_Ord_Diurno = crear_dataframe('Ordinario', resumen_diurno)
    Tipo_A_Diurno = crear_dataframe('Total', resumen_diurno)

    Tipo_A_Dom_Nocturno = crear_dataframe('Dominical', resumen_nocturno)
    Tipo_A_Ord_Nocturno = crear_dataframe('Ordinario', resumen_nocturno)
    Tipo_A_Nocturno = crear_dataframe('Total', resumen_nocturno)

    # Crear DataFrames para instrumentación
    instru_Diu = pd.DataFrame({'uslm': [0.5], 'uresol': [round(0.1 / np.sqrt(12),3)]})
    instru_Noc = pd.DataFrame({'uslm': [0.5], 'uresol': [round(0.1 / np.sqrt(12),3)]})

    # Crear DataFrames para sensibilidad
    sens_Diu = pd.DataFrame({
        'cT dB/°C': [-0.007], 
        'cP dB/kPa': [-0.010],
        "∆Tmax": [MET_resumen_diurno["∆"].iloc[0]],
        "∆Pmax": [MET_resumen_diurno["∆"].iloc[2]],
        "umic,T": [MET_resumen_diurno["∆"].iloc[0] * np.abs(-0.007)],
        "umic,P": [MET_resumen_diurno["∆"].iloc[2] * np.abs(-0.010)],
        "umic,H*": [0.1]
    })
    
    sens_Noc = pd.DataFrame({
        'cT dB/°C': [-0.007], 
        'cP dB/kPa': [-0.010],
        "∆Tmax": [MET_resumen_nocturno["∆"].iloc[0]],
        "∆Pmax": [MET_resumen_nocturno["∆"].iloc[2]],
        "umic,T": [MET_resumen_nocturno["∆"].iloc[0] * np.abs(-0.007)],
        "umic,P": [MET_resumen_nocturno["∆"].iloc[2] * np.abs(-0.010)],
        "umic,H*": [0.1]
    })

    # Crear DataFrames para ubicación
    Ubi_diu = pd.DataFrame({"uloc": [0]})
    Ubi_noc = pd.DataFrame({"uloc": [0]})

    # Calcular incertidumbres combinadas
    columnas = ["udom", "uord", "u"]
    IncComb_Dia = pd.DataFrame(columns=columnas)
    IncComb_Noc = pd.DataFrame(columns=columnas)

    # Calcular incertidumbre para día
    IncComb_Dia.loc[0] = [
        calcular_incertidumbre(Ubi_diu, sens_Diu, instru_Diu, Tipo_A_Dom_Diurno),
        calcular_incertidumbre(Ubi_diu, sens_Diu, instru_Diu, Tipo_A_Ord_Diurno),
        calcular_incertidumbre(Ubi_diu, sens_Diu, instru_Diu, Tipo_A_Diurno)
    ]

    # Calcular incertidumbre para noche
    IncComb_Noc.loc[0] = [
        calcular_incertidumbre(Ubi_noc, sens_Noc, instru_Noc, Tipo_A_Dom_Nocturno),
        calcular_incertidumbre(Ubi_noc, sens_Noc, instru_Noc, Tipo_A_Ord_Nocturno),
        calcular_incertidumbre(Ubi_noc, sens_Noc, instru_Noc, Tipo_A_Nocturno)
    ]

    # Calcular grados de libertad efectivos (veff)
    veff_values = {
        "veff_noc": {
            "veff,dom": calcular_veff(IncComb_Noc["udom"], Tipo_A_Dom_Nocturno, instru_Noc, sens_Noc, Ubi_noc),
            "veff,ord": calcular_veff(IncComb_Noc["uord"], Tipo_A_Ord_Nocturno, instru_Noc, sens_Noc, Ubi_noc),
            "veff": calcular_veff(IncComb_Noc["u"], Tipo_A_Nocturno, instru_Noc, sens_Noc, Ubi_noc)
        },
        "veff_diu": {
            "veff,dom": calcular_veff(IncComb_Dia["udom"], Tipo_A_Dom_Diurno, instru_Diu, sens_Diu, Ubi_diu),
            "veff,ord": calcular_veff(IncComb_Dia["uord"], Tipo_A_Ord_Diurno, instru_Diu, sens_Diu, Ubi_diu),
            "veff": calcular_veff(IncComb_Dia["u"], Tipo_A_Diurno, instru_Diu, sens_Diu, Ubi_diu)
        }
    }

    # Crear DataFrames con los valores calculados
    veff_noc = pd.DataFrame([veff_values["veff_noc"]])
    veff_diu = pd.DataFrame([veff_values["veff_diu"]])

    # Calcular incertidumbre expandida
    alpha = 0.05

    IncExp_diu = pd.DataFrame({
        "K,dom": [t.ppf(1 - alpha / 2, veff_diu["veff,dom"].iloc[0])],
        "U,dom": [(t.ppf(1 - alpha / 2, veff_diu["veff,dom"].iloc[0])*IncComb_Dia["udom"].iloc[0])],
        "K,Ord": [t.ppf(1 - alpha / 2, veff_diu["veff,ord"].iloc[0])],
        "U,Ord": [(t.ppf(1 - alpha / 2, veff_diu["veff,ord"].iloc[0])*IncComb_Dia["uord"].iloc[0])],
        "K": [t.ppf(1 - alpha / 2, veff_diu["veff"].iloc[0])],
        "U": [(t.ppf(1 - alpha / 2, veff_diu["veff"].iloc[0])*IncComb_Dia["u"].iloc[0])]
    })

    IncExp_noc = pd.DataFrame({
        "K,dom": [t.ppf(1 - alpha / 2, veff_noc["veff,dom"].iloc[0])],
        "U,dom": [(t.ppf(1 - alpha / 2, veff_noc["veff,dom"].iloc[0])*IncComb_Noc["udom"].iloc[0])],
        "K,Ord": [t.ppf(1 - alpha / 2, veff_noc["veff,ord"].iloc[0])],
        "U,Ord": [(t.ppf(1 - alpha / 2, veff_noc["veff,ord"].iloc[0])*IncComb_Noc["uord"].iloc[0])],
        "K": [t.ppf(1 - alpha / 2, veff_noc["veff"].iloc[0])],
        "U": [(t.ppf(1 - alpha / 2, veff_noc["veff"].iloc[0])*IncComb_Noc["u"].iloc[0])]
    })
    
    return IncExp_diu, IncExp_noc