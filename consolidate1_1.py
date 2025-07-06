#This project is licensed under the MIT License - see the LICENSE file for details.
#Recuerda que necesitas tener instalado openpyxl y pandas.
#Prototipo 1.1 
#Con amor para D. de D. 

import pandas as pd
import logging
from pathlib import Path
from typing import List, Optional

# Variable definitions
# Aqui se define la ruta de entrada y el nombre del archivo de salida. 
# Predeterminadamente se toma la ruta del script.
# Pega en la misma carpeta los archivos a consolidar.
INPUT_DIRECTORY = Path(__file__).parent if __file__ else Path.cwd()
OUTPUT_FILENAME = "FINAL_REPORT.xlsx"


# Numero de columnas y filas para considerar validas de una hoja.
# Aqui puedes borrar lo que no quieras que se considere 
# ej. la suma de totales o subtotales de una hoja segun el numero de columnas y filas pobladas)
COLUMN_THRESHOLD = 4
ROW_THRESHOLD = 6

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

def find_excel_files(input_dir: Path) -> List[Path]:
    logging.info(f"Searching for .xlsx files in '{input_dir}'...")
    if not input_dir.is_dir():
        raise FileNotFoundError(f"Error: The directory '{input_dir}' was not found."
                                f"Asegúrate de que la ruta sea correcta. Si tienes dudas llamame!")
    excel_files = list(input_dir.glob('*.xlsx'))
    if not excel_files:
        logging.warning(f"No .xlsx files found in '{input_dir}'. ")
        logging.warning(f"Olvidaste pegar tus archivos .xlsx a la carpeta '{input_dir}'.")
    else:
        logging.info(f"Found {len(excel_files)} Excel file(s).")
    return excel_files

def read_and_consolidate_sheets(excel_files: List[Path]) -> Optional[pd.DataFrame]:
    all_dataframes = []
    for file_path in excel_files:
        logging.info(f"Processing file: {file_path.name}")
        try:
            sheets_dict = pd.read_excel(file_path, sheet_name=None)
            for sheet_name, sheet_df in sheets_dict.items():
                sheet_df['source_filename'] = file_path.name
                sheet_df['source_sheet'] = sheet_name
                all_dataframes.append(sheet_df)
        except Exception as e:
            logging.error(f"Could not read file '{file_path.name}'. Error: {e}")
            continue
    if not all_dataframes:
        logging.warning("No data was processed. Aborting consolidation.")
        return None
    logging.info("Combining all data into a single DataFrame...")
    return pd.concat(all_dataframes, ignore_index=True)

def clean_and_write_data(
    consolidated_df: pd.DataFrame, 
    output_path: Path,
    col_thresh: int,
    row_thresh: int
) -> None:
    logging.info(f"Preparing to write final report to '{output_path}'...")
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        grouped = consolidated_df.groupby('source_sheet')
        for sheet_name, group_df in grouped:
            logging.info(f"  - Cleaning and creating sheet: '{sheet_name}'")
            non_null_counts = group_df.count()
            cols_to_drop = non_null_counts[non_null_counts < col_thresh].index
            cleaned_df = group_df.drop(columns=cols_to_drop)
            cleaned_df.dropna(thresh=row_thresh, inplace=True)
            if 'source_sheet' in cleaned_df.columns:
                cleaned_df = cleaned_df.drop(columns=['source_sheet'])
            cleaned_df.to_excel(writer, sheet_name=sheet_name, index=False)
    logging.info(f"✅ Success! Final report created at '{output_path}'.")
    logging.info(f"Ya quedo maric!! Revisa el archivo {output_path.name} en la carpeta {output_path.parent}.")

def main():
    input_path = Path(INPUT_DIRECTORY)
    output_path = input_path / OUTPUT_FILENAME
    try:
        excel_files = find_excel_files(input_path)
        if not excel_files:
            return
        consolidated_df = read_and_consolidate_sheets(excel_files)
        if consolidated_df is None:
            return
        clean_and_write_data(
            consolidated_df, 
            output_path,
            COLUMN_THRESHOLD,
            ROW_THRESHOLD
        )
    except FileNotFoundError as e:
        logging.error(e)
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    main()
    
