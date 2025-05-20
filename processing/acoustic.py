import numpy as np
import pandas as pd
from data.constants import FREQUENCIES, PONDERATION

def Ponderacion_A(spectrum_k1_str, valor_entrada):
    """
    Corrige el valor de entrada sumando la ponderación A correspondiente
    
    Args:
        spectrum_k1_str: Frecuencia en forma de string
        valor_entrada: Valor de nivel sonoro a corregir
    
    Returns:
        Valor corregido con ponderación A
    """
    # Convertir la frecuencia de la cadena (por ejemplo '10k') a Hz
    if 'k' in spectrum_k1_str or 'K' in spectrum_k1_str:
        frecuencia_entrada = float(spectrum_k1_str.replace('k', '').replace('K', '')) * 1000
    else:
        frecuencia_entrada = float(spectrum_k1_str)

    # Buscar el índice de la frecuencia más cercana en la tabla
    idx = np.searchsorted(FREQUENCIES, frecuencia_entrada)

    # Verificar si la frecuencia ingresada está fuera del rango de la tabla
    if idx == 0:
        ponderacion = PONDERATION[0]
    elif idx == len(FREQUENCIES):
        ponderacion = PONDERATION[-1]
    else:
        # Verificar cuál es la frecuencia más cercana
        if abs(FREQUENCIES[idx] - frecuencia_entrada) < abs(FREQUENCIES[idx - 1] - frecuencia_entrada):
            ponderacion = PONDERATION[idx]
        else:
            ponderacion = PONDERATION[idx - 1]

    # Sumar la ponderación al valor de entrada
    valor_corregido = valor_entrada + ponderacion
    return valor_corregido

def ajuste_tonal(spectrum, levels):
    """
    Calcula el ajuste tonal según la Resolución colombiana 627 de 2006.

    Args:
        spectrum: Lista de frecuencias centrales de las bandas de tercio de octava.
        levels: Lista de niveles de presión sonora correspondientes.
    
    Returns:
        Valor de ajuste tonal más alto y el rango de bandas donde se encontró el ajuste tonal.
    """
    # Inicialización de variables
    adj = [0, 'No hay ajuste tonal']
    b = [0, 0, 0]
    b_range = ['<= 125 Hz', '>= 160 Hz & <= 400 Hz', '>= 500 Hz']
    
    # Convertir levels a floats si es una lista o una Serie de pandas
    if isinstance(levels, list):
        levels = list(map(float, levels))
    elif hasattr(levels, 'astype'):  # Si levels es un pandas.Series
        levels = levels.astype(float)

    # Verificación de que el número de elementos en spectrum y levels coincidan
    if len(spectrum) == len(levels):
        # Recorre las bandas de frecuencia
        for k1 in range(1, len(spectrum) - 1):
            # Convierte frecuencias con 'k' o 'K' a valor numérico en Hz
            spectrum_k1_str = str(spectrum[k1])  # Asegurarse de que sea cadena
            if 'k' in spectrum_k1_str.lower():
                spectrum[k1] = float(spectrum_k1_str.replace('k', '').replace('K', '')) * 1000

            lt = levels[k1]  # Usar el índice para acceder
            ls = (levels[k1 - 1] + levels[k1 + 1]) / 2  # Usar el índice para acceder
            l = lt - ls

            # Evalúa el ajuste tonal en las bandas de frecuencias
            if 20 <= float(spectrum[k1]) <= 125:  # Baja frecuencia
                if -ls <= l <= 8:
                    if adj[0] < 3:
                        adj[0] = 0
                elif 8 < l <= 12:
                    if adj[0] < 6:
                        adj[0] = 3
                        b[0] = 1
                elif 12 < l <= lt:
                    adj[0] = 6
                    b[0] = 1
            elif 160 <= float(spectrum[k1]) <= 400:  # Frecuencias medias
                if -ls <= l <= 5:
                    if adj[0] < 3:
                        adj[0] = 0
                elif 5 < l <= 8:
                    if adj[0] < 6:
                        adj[0] = 3
                        b[1] = 1
                elif 8 < l <= lt:
                    adj[0] = 6
                    b[1] = 1
            elif float(spectrum[k1]) >= 500:  # Alta frecuencia
                if -ls <= l <= 3:
                    if adj[0] < 3:
                        adj[0] = 0
                elif 3 < l <= 5:
                    if adj[0] < 6:
                        adj[0] = 3
                        b[2] = 1
                elif 5 < l <= lt:
                    adj[0] = 6
                    b[2] = 1

        # Verificación de los rangos de bandas en los que se encontró ajuste tonal
        x = 0
        for a in range(3):
            if b[a] == 1:
                x += 1
                if x == 1:
                    adj[1] = b_range[a]
                else:
                    adj[1] += f"; {b_range[a]}"

        if x == 0:
            adj[1] = "No hay ajuste tonal"
    else:
        adj[0] = "Error"
        adj[1] = "El número de celdas no coincide"
    
    return adj

def calcular_ki(diff):
    """
    Calcula el factor de corrección KI según la diferencia entre LAIeq y LASeq
    
    Args:
        diff: Diferencia entre LAIeq y LASeq
        
    Returns:
        Factor de corrección KI
    """
    if pd.isna(diff):
        return np.nan
    if diff < 3:
        return 0
    elif diff < 6:
        return 3
    else:
        return 6

def aplicar_Correccion(row):
    """
    Aplica la corrección al nivel LASeq,i con el máximo de KI,i y KT,i
    
    Args:
        row: Fila del DataFrame con los valores LASeq,i, KI,i y KT,i
        
    Returns:
        Valor corregido LRASeq,i
    """
    try:
        # Obtener los valores de las columnas correspondientes
        LASeq_i = row['LASeq,i']
        KI_i = row['KI,i']
        KT_i = row['KT,i']
        
        # Realizar la suma y el máximo
        return LASeq_i + max(KI_i, KT_i)
    except Exception as e:
        # Si ocurre un error, retornar el valor '—'
        return '—'

def Nivel_Eq_diaria(row):
    """
    Calcula el nivel equivalente diario LRASeq,1d a partir de LASeq_1d y aplicando
    el máximo de las correcciones KI,1d y KT,1d
    
    Args:
        row: Fila del DataFrame con los valores LASeq_1d, KI,1d y KT,1d
        
    Returns:
        Valor de nivel equivalente diario LRASeq,1d
    """
    try:
        # Obtener los valores de las columnas correspondientes
        LASeq_i = row['LASeq_1d']
        KI_i = row['KI,1d']
        KT_i = row['KT,1d']
        
        # Realizar la suma y el máximo
        return LASeq_i + max(KI_i, KT_i)
    except Exception as e:
        # Si ocurre un error, retornar el valor '—'
        return '—'

def calcular_L_Raseq_dn(L_Raseq_d, L_Raseq_n):
    """
    Calcula el nivel equivalente día-noche según la fórmula establecida
    
    Args:
        L_Raseq_d: Nivel equivalente diurno
        L_Raseq_n: Nivel equivalente nocturno
        
    Returns:
        Nivel equivalente día-noche
    """
    # Calculando el valor de L_Raseq_dn según la fórmula dada usando NumPy
    parte1 = 14 * np.power(10, 0.1 * L_Raseq_d)
    parte2 = 10 * np.power(10, 0.1 * (L_Raseq_n + 10))
    resultado = 10 * np.log10((1 / 24) * (parte1 + parte2))
    return resultado

def calcular_declaracion(row):
    """
    Aplica la lógica de la columna "Declaracion" según los criterios establecidos
    para períodos
    
    Args:
        row: Fila con los campos necesarios para la evaluación
        
    Returns:
        Texto con la declaración de cumplimiento
    """
    if row["w"] == 0:
        return "—"
    elif row["w"] > 0:
        if row["LRASeq_k"] <= row["Au"]:
            return "Pasa"
        elif row["LRASeq_k"] <= row["Tu"]:
            return "Pasa condicional"
        elif row["LRASeq_k"] <= row["Tu"] + row["w"]:
            return "No pasa condicional"
        else:
            return "No pasa"
    else:
        if row["LRASeq_k"] <= row["Tu"]:
            return "Pasa"
        elif row["LRASeq_k"] <= row["Au"]:
            return "Pasa condicional"
        else:
            return "No pasa"

def calcular_declaracion_diaria(row):
    """
    Aplica la lógica de la columna "Declaracion" según los criterios establecidos
    para días individuales
    
    Args:
        row: Fila con los campos necesarios para la evaluación
        
    Returns:
        Texto con la declaración de cumplimiento
    """
    if row["w"] == 0:
        return "—"
    elif row["w"] > 0:
        if row["LRASeq_1d"] <= row["Au"]:
            return "Pasa"
        elif row["LRASeq_1d"] <= row["Tu"]:
            return "Pasa condicional"
        elif row["LRASeq_1d"] <= row["Tu"] + row["w"]:
            return "No pasa condicional"
        else:
            return "No pasa"
    else:
        if row["LRASeq_1d"] <= row["Tu"]:
            return "Pasa"
        elif row["LRASeq_1d"] <= row["Au"]:
            return "Pasa condicional"
        else:
            return "No pasa"