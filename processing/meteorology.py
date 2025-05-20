import pandas as pd
import numpy as np
from data.constants import ESTACIONES_MET, HORAS_REFERENCIA
from utils.date_utils import corregir_fecha_hora
import os

def filter_by_time_range(df, start_time, end_time, is_night_range=False):
    """
    Filter DataFrame by time range, safely handling type mismatches
    
    Args:
        df (pd.DataFrame): DataFrame with datetime index
        start_time (time): Start time for filtering
        end_time (time): End time for filtering
        is_night_range (bool): Whether this is a nighttime range that crosses midnight
        
    Returns:
        pd.DataFrame: Filtered DataFrame
    """
    if df is None or df.empty:
        return pd.DataFrame()
    
    try:
        # Ensure the index is datetime type
        if not isinstance(df.index, pd.DatetimeIndex):
            print("Warning: Index is not a DatetimeIndex, attempting to convert")
            try:
                df.index = pd.to_datetime(df.index)
            except:
                print("Failed to convert index to datetime. Returning empty DataFrame")
                return pd.DataFrame()
        
        # Extract time component from datetime index
        time_index = df.index.time
        
        # For nighttime (crossing midnight), use OR condition
        if is_night_range:
            mask = (time_index >= start_time) | (time_index <= end_time)
        # For daytime (no midnight crossing), use AND condition
        else:
            mask = (time_index >= start_time) & (time_index <= end_time)
            
        return df.loc[mask]
        
    except Exception as e:
        print(f"Error in time filtering: {str(e)}")
        return pd.DataFrame()

def process_and_export_weather_data(file_path, Estacion, numero):
    """
    Procesa y exporta datos meteorológicos para una estación específica
    
    Args:
        file_path: Ruta al archivo Excel con datos meteorológicos
        Estacion: Código de la estación a procesar
        numero: Número para el archivo de salida
        
    Returns:
        Tuple con DataFrames de resultados meteorológicos
    """
    # Fix Estacion key format - convert to string and replace hyphen if needed
    Estacion_key = str(Estacion).replace("-", "_")
    
    if Estacion_key not in ESTACIONES_MET:
        print(f"La estación '{Estacion}' (key: '{Estacion_key}') no existe en el diccionario.")
        return None, None, None, None, None, None

    station_column_name = ESTACIONES_MET[Estacion_key].strip()
    
    # Special handling for SDA stations or other special values
    if station_column_name == 'SDA':
        # Create empty dataframes with proper structure for SDA stations
        empty_df = pd.DataFrame({'Fecha_Hora': pd.date_range(start='2024-04-01', periods=24, freq='1H')})
        empty_df.set_index('Fecha_Hora', inplace=True)
        
        # Add empty columns for required meteorological parameters
        for col in ['TEMP', 'HUM', 'PRES', 'PREC']:
            empty_df[col] = np.nan
            
        # Create empty summary dataframes
        empty_summary = pd.DataFrame({'Variable': ['TEMP', 'HUM', 'PRES', 'PREC'], 
                                     'MAX': [np.nan]*4, 
                                     'MIN': [np.nan]*4, 
                                     '∆': [np.nan]*4})
        
        # Create output directory if it doesn't exist
        output_dir = 'PTOS_salida'
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            print(f"Created directory: {output_dir}")
        
        output_file = f'{output_dir}/MET{numero}.xlsx'
        
        # Export empty dataframes
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            empty_df.to_excel(writer, sheet_name='MET', index=True, startcol=0)
            empty_summary.to_excel(writer, sheet_name='MET', index=False, startcol=empty_df.shape[1] + 2)
            
        print(f"SDA station: Empty data exported to {output_file}")
        return empty_df, empty_df, empty_df, empty_summary, empty_summary, empty_summary
    
    # Normal processing for regular stations
    try:
        excel_file = pd.ExcelFile(file_path)
        final_df = None

        for sheet_name in ['TEMP', 'HUM', 'PRES', 'PREC']:
            if sheet_name in excel_file.sheet_names:
                df = excel_file.parse(sheet_name)
                df.columns = df.columns.str.strip()
                matching_columns = [col for col in df.columns if station_column_name in col]

                if matching_columns:
                    if final_df is None:
                        final_df = pd.DataFrame({'Fecha_Hora': pd.to_datetime(df['Fecha'], errors='coerce')})
                        final_df['Fecha_Hora'] = final_df['Fecha_Hora'].apply(corregir_fecha_hora)
                    final_df[sheet_name] = df[matching_columns[0]].values
                else:
                    print(f"No se encontró columna para {sheet_name} con estación {station_column_name}")

        if final_df is None or final_df.empty:
            print(f"No se encontraron datos para la estación '{Estacion}'")
            return None, None, None, None, None, None

        final_df.set_index('Fecha_Hora', inplace=True)

        # Use the safe filtering function instead of direct comparison
        MET_Diurno = filter_by_time_range(final_df, 
                                        HORAS_REFERENCIA["diurna_inicio"].time(), 
                                        HORAS_REFERENCIA["diurna_fin"].time(), 
                                        is_night_range=False)
        
        MET_Nocturno = filter_by_time_range(final_df, 
                                          HORAS_REFERENCIA["nocturna_inicio"].time(), 
                                          HORAS_REFERENCIA["nocturna_fin"].time(), 
                                          is_night_range=True)

        # Safe calculation of statistics with empty dataframe handling
        summary_df = pd.DataFrame([{ 
            'Variable': col, 
            'MAX': final_df[col].max(), 
            'MIN': final_df[col].min(),
            '∆': final_df[col].max() - final_df[col].min()
        } for col in final_df.columns])
        
        diurno_summary_df = pd.DataFrame([{ 
            'Variable': col, 
            'MAX': MET_Diurno[col].max() if not MET_Diurno.empty else np.nan, 
            'MIN': MET_Diurno[col].min() if not MET_Diurno.empty else np.nan,
            '∆': (MET_Diurno[col].max() - MET_Diurno[col].min()) if not MET_Diurno.empty else np.nan
        } for col in MET_Diurno.columns])
        
        nocturno_summary_df = pd.DataFrame([{ 
            'Variable': col, 
            'MAX': MET_Nocturno[col].max() if not MET_Nocturno.empty else np.nan, 
            'MIN': MET_Nocturno[col].min() if not MET_Nocturno.empty else np.nan,
            '∆': (MET_Nocturno[col].max() - MET_Nocturno[col].min()) if not MET_Nocturno.empty else np.nan
        } for col in MET_Nocturno.columns])
        
        # Create output directory if it doesn't exist
        output_dir = 'PTOS_salida'
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            print(f"Created directory: {output_dir}")
        
        output_file = f'{output_dir}/MET{numero}.xlsx'
        
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            final_df.to_excel(writer, sheet_name='MET', index=True, startcol=0)
            summary_df.to_excel(writer, sheet_name='MET', index=False, startcol=final_df.shape[1] + 2)

        print(f"Datos exportados exitosamente a: {output_file}")
        
        return final_df, MET_Diurno, MET_Nocturno, summary_df, diurno_summary_df, nocturno_summary_df
        
    except Exception as e:
        print(f"Error processing meteorological data for station {Estacion}: {str(e)}")
        return None, None, None, None, None, None