import pandas as pd
import numpy as np
from scipy.stats import t, norm
from data.limits import LIMITE_0627_DIA, LIMITE_0627_NOCHE
from processing.acoustic import calcular_declaracion, calcular_declaracion_diaria, Nivel_Eq_diaria

def procesar_compliance_diurno(resumen_diurno, diurno_grouped, IncExp_diu):
    """
    Procesa el cumplimiento para el período diurno
    
    Args:
        resumen_diurno: DataFrame con resumen diurno
        diurno_grouped: DataFrame con datos diurnos agrupados
        IncExp_diu: DataFrame con incertidumbre expandida diurna
        
    Returns:
        Tupla con (resumen_diurno actualizado, diurno_grouped actualizado)
    """
    # Asignamos valores basados en `TipoDia`
    resumen_diurno['K'] = resumen_diurno['TipoDia'].apply(
        lambda x: IncExp_diu['K,dom'][0] if x == 'Dominical' 
        else IncExp_diu['K,Ord'][0] if x == 'Ordinario' 
        else IncExp_diu['K'][0] if x == 'Total' 
        else None
    )

    resumen_diurno['U'] = resumen_diurno['TipoDia'].apply(
        lambda x: IncExp_diu['U,dom'][0] if x == 'Dominical' 
        else IncExp_diu['U,Ord'][0] if x == 'Ordinario' 
        else IncExp_diu['U'][0] if x == 'Total' 
        else None
    )

    # Asignar valores para las columnas de resumen_diurno
    resumen_diurno["E"] = resumen_diurno.apply(lambda row: max(row["LRASeq_k"] - row["Tu"], 0), axis=1)
    resumen_diurno["w"] = resumen_diurno.apply(lambda row: -row['U'], axis=1)
    resumen_diurno["Au"] = resumen_diurno.apply(lambda row: row['Tu']-row['w'], axis=1)
    resumen_diurno["Z"] = (resumen_diurno["Tu"] - resumen_diurno["LRASeq_k"]) / (resumen_diurno["U"] / resumen_diurno["K"])

    resumen_diurno["Rp*=Pc"] = resumen_diurno.apply(
        lambda row: norm.cdf((row["LRASeq_k"] + (row["U"] / row["K"])*row["Z"]), row["LRASeq_k"], row["U"] / row["K"])
        if row["U"] and row["K"] else "—",
        axis=1
    )
    resumen_diurno["Rc"] = resumen_diurno.apply(
        lambda row: 1-row["Rp*=Pc"] if isinstance(row["Rp*=Pc"], (int, float)) else "—",
        axis=1
    )

    resumen_diurno["Declaracion"] = resumen_diurno.apply(calcular_declaracion, axis=1)
    
    # Asignamos valores para diurno_grouped basados en `tipoDia`
    diurno_grouped['K'] = diurno_grouped['TipoDia'].apply(lambda x: IncExp_diu['K,dom'][0] if x == 'Dominical' else IncExp_diu['K,Ord'][0])
    diurno_grouped['U'] = diurno_grouped['TipoDia'].apply(lambda x: IncExp_diu['U,dom'][0] if x == 'Dominical' else IncExp_diu['U,Ord'][0])
    
    # Calcular campos para diurno_grouped
    diurno_grouped["E"] = diurno_grouped.apply(lambda row: max(row["LRASeq_1d"] - row["Tu"], 0), axis=1)
    diurno_grouped["w"] = diurno_grouped.apply(lambda row: -row['U'], axis=1)
    diurno_grouped['Au'] = diurno_grouped['TipoDia'].apply(lambda x: resumen_diurno.loc[resumen_diurno['TipoDia'] == 'Dominical', "Au"].values[0] if x == 'Dominical' else resumen_diurno.loc[resumen_diurno['TipoDia'] == 'Ordinario', "Au"].values[0])
    diurno_grouped["Z"] = (diurno_grouped["Tu"] - diurno_grouped["LRASeq_1d"]) / (diurno_grouped["U"] / diurno_grouped["K"])
    
    # Calcular la distribución normal acumulativa (CDF)
    diurno_grouped["Rp*=Pc"] = diurno_grouped.apply(
        lambda row: norm.cdf((row["LRASeq_1d"] + (row["U"] / row["K"])*row["Z"]), row["LRASeq_1d"], row["U"] / row["K"])
        if row["U"] and row["K"] else "—",
        axis=1
    )
    diurno_grouped["Rc"] = diurno_grouped.apply(
        lambda row: 1-row["Rp*=Pc"] if isinstance(row["Rp*=Pc"], (int, float)) else "—",
        axis=1
    )
    diurno_grouped["Declaracion"] = diurno_grouped.apply(calcular_declaracion_diaria, axis=1)
    
    return resumen_diurno, diurno_grouped

def procesar_compliance_nocturno(resumen_nocturno, nocturno_grouped, IncExp_noc):
    """
    Procesa el cumplimiento para el período nocturno
    
    Args:
        resumen_nocturno: DataFrame con resumen nocturno
        nocturno_grouped: DataFrame con datos nocturnos agrupados
        IncExp_noc: DataFrame con incertidumbre expandida nocturna
        
    Returns:
        Tupla con (resumen_nocturno actualizado, nocturno_grouped actualizado)
    """
    # Asignamos valores basados en `TipoDia`
    resumen_nocturno['K'] = resumen_nocturno['TipoDia'].apply(
        lambda x: IncExp_noc['K,dom'][0] if x == 'Dominical' 
        else IncExp_noc['K,Ord'][0] if x == 'Ordinario' 
        else IncExp_noc['K'][0] if x == 'Total' 
        else None
    )

    resumen_nocturno['U'] = resumen_nocturno['TipoDia'].apply(
        lambda x: IncExp_noc['U,dom'][0] if x == 'Dominical' 
        else IncExp_noc['U,Ord'][0] if x == 'Ordinario' 
        else IncExp_noc['U'][0] if x == 'Total' 
        else None
    )

    # Asignar valores para las columnas de resumen_nocturno
    resumen_nocturno["E"] = resumen_nocturno.apply(lambda row: max(row["LRASeq_k"] - row["Tu"], 0), axis=1)
    resumen_nocturno["w"] = resumen_nocturno.apply(lambda row: -row['U'], axis=1)
    resumen_nocturno["Au"] = resumen_nocturno.apply(lambda row: row['Tu']-row['w'], axis=1)
    resumen_nocturno["Z"] = (resumen_nocturno["Tu"] - resumen_nocturno["LRASeq_k"]) / (resumen_nocturno["U"] / resumen_nocturno["K"])
    
    # Calcular la distribución normal acumulativa (CDF)
    resumen_nocturno["Rp*=Pc"] = resumen_nocturno.apply(
        lambda row: norm.cdf((row["LRASeq_k"] + (row["U"] / row["K"])*row["Z"]), row["LRASeq_k"], row["U"] / row["K"])
        if row["U"] and row["K"] else "—",
        axis=1
    )
    resumen_nocturno["Rc"] = resumen_nocturno.apply(
        lambda row: 1-row["Rp*=Pc"] if isinstance(row["Rp*=Pc"], (int, float)) else "—",
        axis=1
    )
    resumen_nocturno["Declaracion"] = resumen_nocturno.apply(calcular_declaracion, axis=1)
    
    # Asignamos valores para nocturno_grouped
    nocturno_grouped['K'] = nocturno_grouped['TipoDia'].apply(lambda x: IncExp_noc['K,dom'][0] if x == 'Dominical' else IncExp_noc['K,Ord'][0])
    nocturno_grouped['U'] = nocturno_grouped['TipoDia'].apply(lambda x: IncExp_noc['U,dom'][0] if x == 'Dominical' else IncExp_noc['U,Ord'][0])
    
    # Calcular campos para nocturno_grouped
    nocturno_grouped["E"] = nocturno_grouped.apply(lambda row: max(row["LRASeq_1d"] - row["Tu"], 0), axis=1)
    nocturno_grouped["w"] = nocturno_grouped.apply(lambda row: -row['U'], axis=1)
    nocturno_grouped['Au'] = nocturno_grouped['TipoDia'].apply(lambda x: resumen_nocturno.loc[resumen_nocturno['TipoDia'] == 'Dominical', "Au"].values[0] if x == 'Dominical' else resumen_nocturno.loc[resumen_nocturno['TipoDia'] == 'Ordinario', "Au"].values[0])
    nocturno_grouped["Z"] = (nocturno_grouped["Tu"] - nocturno_grouped["LRASeq_1d"]) / (nocturno_grouped["U"] / nocturno_grouped["K"])
    
    # Calcular la distribución normal acumulativa (CDF)
    nocturno_grouped["Rp*=Pc"] = nocturno_grouped.apply(
        lambda row: norm.cdf((row["LRASeq_1d"] + (row["U"] / row["K"])*row["Z"]), row["LRASeq_1d"], row["U"] / row["K"])
        if row["U"] and row["K"] else "—",
        axis=1
    )
    nocturno_grouped["Rc"] = nocturno_grouped.apply(
        lambda row: 1-row["Rp*=Pc"] if isinstance(row["Rp*=Pc"], (int, float)) else "—",
        axis=1
    )
    nocturno_grouped["Declaracion"] = nocturno_grouped.apply(calcular_declaracion_diaria, axis=1)
    
    return resumen_nocturno, nocturno_grouped

def asignar_limites(resumen_diurno, resumen_nocturno, Estacion):
    """
    Asigna los límites regulatorios a los DataFrames
    
    Args:
        resumen_diurno: DataFrame con resumen diurno
        resumen_nocturno: DataFrame con resumen nocturno
        Estacion: Código de la estación
        
    Returns:
        Tupla con (resumen_diurno actualizado, resumen_nocturno actualizado)
    """
    # Asignar el valor del diccionario basado en la variable 'estacion' a todas las filas
    resumen_diurno['Tu'] = LIMITE_0627_DIA.get(Estacion, None)
    resumen_nocturno['Tu'] = LIMITE_0627_NOCHE.get(Estacion, None)
    
    return resumen_diurno, resumen_nocturno

def asignar_limites_diarios(diurno_grouped, nocturno_grouped, Estacion):
    """
    Asigna los límites regulatorios a los DataFrames diarios
    
    Args:
        diurno_grouped: DataFrame con datos diurnos agrupados
        nocturno_grouped: DataFrame con datos nocturnos agrupados
        Estacion: Código de la estación
        
    Returns:
        Tupla con (diurno_grouped actualizado, nocturno_grouped actualizado)
    """
    # Asignar el valor del diccionario basado en la variable 'estacion' a todas las filas
    diurno_grouped['Tu'] = LIMITE_0627_DIA.get(Estacion, None)
    nocturno_grouped['Tu'] = LIMITE_0627_NOCHE.get(Estacion, None)
    
    return diurno_grouped, nocturno_grouped

def finalizar_agrupados(diurno_grouped, nocturno_grouped, DfAjusteTonal_diurno_ref, DfAjusteTonal_nocturno_ref):
    """
    Finaliza el procesamiento de los DataFrames agrupados
    
    Args:
        diurno_grouped: DataFrame con datos diurnos agrupados
        nocturno_grouped: DataFrame con datos nocturnos agrupados
        DfAjusteTonal_diurno_ref: DataFrame con ajuste tonal diurno
        DfAjusteTonal_nocturno_ref: DataFrame con ajuste tonal nocturno
        
    Returns:
        Tupla con (diurno_grouped actualizado, nocturno_grouped actualizado)
    """
    # Añadir columnas KT,1d y Bandas
    if len(diurno_grouped) > 0 and len(DfAjusteTonal_diurno_ref) > 0:
        diurno_grouped['KT,1d'] = DfAjusteTonal_diurno_ref["KT,i"].reset_index(drop=True)
        diurno_grouped['Bandas'] = DfAjusteTonal_diurno_ref["Bandas"].reset_index(drop=True)

    if len(nocturno_grouped) > 0 and len(DfAjusteTonal_nocturno_ref) > 0:
        nocturno_grouped['KT,1d'] = DfAjusteTonal_nocturno_ref["KT,i"].reset_index(drop=True)
        nocturno_grouped['Bandas'] = DfAjusteTonal_nocturno_ref["Bandas"].reset_index(drop=True)

    # Calcular LRASeq,1d
    from processing.acoustic import Nivel_Eq_diaria
    diurno_grouped['LRASeq,1d'] = diurno_grouped.apply(Nivel_Eq_diaria, axis=1)

    # Añadir columna TipoDia
    import pandas as pd
    import numpy as np
    
    diurno_grouped['TipoDia'] = np.where(
        pd.to_datetime(diurno_grouped['Fechas']).dt.dayofweek == 6, 
        'Dominical', 
        'Ordinario'
    )

    nocturno_grouped['TipoDia'] = np.where(
        pd.to_datetime(nocturno_grouped['Fechas']).dt.dayofweek == 6, 
        'Dominical', 
        'Ordinario'
    )

    # Reorganizar columnas
    column_order = ['Fechas', 'TipoDia', 'Nm_1d', 'LASeq_1d', 'LAIeq_1d', 'KI,1d', 'KT,1d', 'Bandas', 'LRASeq_1d']
    diurno_grouped = diurno_grouped[column_order]
    nocturno_grouped = nocturno_grouped[column_order]
    
    return diurno_grouped, nocturno_grouped