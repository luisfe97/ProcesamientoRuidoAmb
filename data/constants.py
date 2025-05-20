import pandas as pd
import numpy as np

# Horas de referencia
HORAS_REFERENCIA = {
    "diurna_inicio": pd.Timestamp('07:00:00'),
    "diurna_fin": pd.Timestamp('20:00:00'),
    "nocturna_inicio": pd.Timestamp('21:00:00'),
    "nocturna_fin": pd.Timestamp('06:00:00')
}

# Frecuencias (en Hz) y ponderaciones de la tabla proporcionada
FREQUENCIES = np.array([6.3, 8, 10, 12.5, 16, 20, 25, 31.5, 40, 50, 63, 80, 100, 125, 160, 200, 250, 
                        315, 400, 500, 630, 800, 1000, 1250, 1600, 2000, 2500, 3150, 4000, 5000, 
                        6300, 8000, 10000, 12500, 16000, 20000])

PONDERATION = np.array([-85.4, -77.8, -70.4, -63.4, -56.7, -50.5, -44.7, -39.4, -34.6, -30.2, -26.2, 
                        -22.5, -19.1, -16.1, -13.4, -10.9, -8.6, -6.6, -4.8, -3.2, -1.9, -0.8, 0, 
                        0.6, 1, 1.2, 1.3, 1.2, 1, 0.5, -0.1, -1.1, -2.5, -4.3, -6.6, -9.3])

# Lista de hojas a procesar
SHEETS_TO_PROCESS = [
    "EMRI1", "EMRI2", "EMRI3", "EMRI4", "EMRI5",
    "EMRI7", "EMRI8", "EMRI10", "EMRI11", "EMRI13", "EMRI15", "EMRI17", "EMRI18",
    "EMRI19", "EMRI20", "EMRI21", "EMRI23", "EMRI24", "EMRI28", "EMRI29",
    "EMRI32", "EMRI33", "EMRI34", "EMRI35", "EMRI37", "EMRI38", "EMRI39"
]

# Archivo Excel
ARCHIVO_EXCEL = "Input/Met_Mar.xlsx"

# Carpeta de salida
OUTPUT_FOLDER = 'PTOS_salida'

# Diccionario de estaciones meteorológicas
ESTACIONES_MET = {
    "EMRI_1": "EMRI 8 CE0331",
    "EMRI_2": "EMRI 2 CE0337",
    "EMRI_3": "EMRI 3 CE0117",
    "EMRI_4": "EMRI 4 CE0338",
    "EMRI_5": "EMRI 4 CE0338",
    "EMRI_7": "EMRI 7 CE0335",
    "EMRI_8": "EMRI 8 CE0331",
    "EMRI_10": "EMRI 2 CE0337",
    "EMRI_11": "EMRI 4 CE0338",
    "EMRI_13": "CA Engativa CE0350",
    "EMRI_15": "CA Fontibon CE0349",
    "EMRI_17": "EMRI 17 CE0333",
    "EMRI_18": "EMRI 2 CE0337",
    "EMRI_19": "CA4 Funza CE0351",
    "EMRI_20": "CA4 Funza CE0351",
    "EMRI_21": "EMRI 7 CE0335",
    "EMRI_23": "EMRI 3 CE0117",
    "EMRI_24": "EMRI 4 CE0338",
    "EMRI_39": "EMRI 3 CE0117",
    "EMRI_38": "EMRI 3 CE0117",
    "EMRI_28": "CA Engativa CE0350",
    "EMRI_29": "EMRI 4 CE0338",
    "EMRI_37": "EMRI 3 CE0117",
    "EMRI_32": "CA Fontibon CE0349",
    "EMRI_33": "EMRI 2 CE0337",
    "EMRI_34": "SDA",
    "EMRI_35": "SDA"
}