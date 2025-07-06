Script Consolidator (consolidate1_1.py)
Resumen

Este script de Python está diseñado para consolidar datos de múltiples archivos Excel (.xlsx) y sus hojas en un único archivo Excel de salida. Limpia los datos eliminando columnas con pocos valores no nulos y filas con pocos valores no nulos, y luego organiza los datos por hoja original en el nuevo archivo consolidado.

Propósito:

    Consolidar datos de múltiples hojas y archivos Excel en un solo archivo FINAL_REPORT.xlsx.

    Limpiar los datos eliminando columnas y filas con datos insuficientes según umbrales definidos.

    Mantener la estructura original de las hojas en el archivo de salida, agrupando los datos por su hoja de origen.

Características

    Busca automáticamente todos los archivos .xlsx en el directorio de entrada.

    Lee todas las hojas de cada archivo Excel encontrado.

    Agrega columnas source_filename y source_sheet para rastrear el origen de los datos.

    Consolida todos los datos en un único DataFrame de pandas.

    Limpia los datos eliminando columnas con menos de COLUMN_THRESHOLD valores no nulos.

    Limpia los datos eliminando filas con menos de ROW_THRESHOLD valores no nulos.

    Escribe los datos limpios en un nuevo archivo Excel, con cada hoja del archivo de salida correspondiendo a una hoja de origen de los datos consolidados.

    Incluye un registro detallado del proceso.

Prerrequisitos

Antes de ejecutar este script, asegúrate de tener lo siguiente instalado:

    Python: Versión 3.x (probado con Python 3.8+)

    Librerías Python Requeridas:

        pandas (para manipulación de datos)

        openpyxl (motor para leer y escribir archivos Excel)

Puedes instalar estas librerías usando pip:

pip install pandas openpyxl

Instalación

    Descargar el script:
    Descarga el archivo consolidate1_1.py en el directorio donde tienes tus archivos Excel a consolidar.

Uso

    Coloca tus archivos Excel:
    Asegúrate de que todos los archivos .xlsx que deseas consolidar estén en el mismo directorio que el script consolidate1_1.py.

    Configuración (Opcional):
    Puedes ajustar las siguientes variables directamente en el script si es necesario:

        INPUT_DIRECTORY: La ruta donde el script buscará los archivos Excel. Por defecto, es el directorio donde se encuentra el script.

        OUTPUT_FILENAME: El nombre del archivo Excel de salida (por defecto: "FINAL_REPORT.xlsx").

        COLUMN_THRESHOLD: El número mínimo de valores no nulos que una columna debe tener para no ser eliminada (por defecto: 4).

        ROW_THRESHOLD: El número mínimo de valores no nulos que una fila debe tener para no ser eliminada (por defecto: 6).

    Ejecución:
    Abre una terminal o línea de comandos, navega al directorio donde guardaste el script y tus archivos Excel, y ejecuta:

    python consolidate1_1.py

El script buscará los archivos, los procesará y creará el FINAL_REPORT.xlsx en el mismo directorio.
Solución de Problemas

    ModuleNotFoundError: Si ves un error como ModuleNotFoundError: No module named 'pandas', significa que no has instalado las librerías requeridas. Ejecuta pip install pandas openpyxl.

    Archivos Excel no encontrados: Asegúrate de que tus archivos .xlsx estén en el mismo directorio que el script. El script mostrará una advertencia si no encuentra ninguno.

    El archivo de salida está vacío o faltan datos:

        Verifica los umbrales COLUMN_THRESHOLD y ROW_THRESHOLD. Si son demasiado altos, el script podría estar eliminando demasiadas columnas o filas.

        Asegúrate de que tus archivos Excel no estén corruptos o protegidos con contraseña.


Consolidator Script (consolidate1_1.py)
Overview

This Python script is designed to consolidate data from multiple Excel (.xlsx) files and their sheets into a single output Excel file. It cleans the data by dropping columns with too few non-null values and rows with too few non-null values, then organizes the data by original sheet in the new consolidated file.

Purpose:

    To consolidate data from multiple Excel files and sheets into a single FINAL_REPORT.xlsx file.

    To clean data by removing columns and rows with insufficient data based on defined thresholds.

    To maintain the original sheet structure in the output file, grouping data by its source sheet.

Features

    Automatically searches for all .xlsx files in the input directory.

    Reads all sheets from each found Excel file.

    Adds source_filename and source_sheet columns to track data origin.

    Consolidates all data into a single pandas DataFrame.

    Cleans data by dropping columns with fewer than COLUMN_THRESHOLD non-null values.

    Cleans data by dropping rows with fewer than ROW_THRESHOLD non-null values.

    Writes the cleaned data to a new Excel file, with each sheet in the output file corresponding to a source sheet from the consolidated data.

    Includes detailed logging of the process.

Prerequisites

Before running this script, ensure you have the following installed:

    Python: Version 3.x (tested with Python 3.8+)

    Required Python Libraries:

        pandas (for data manipulation)

        openpyxl (engine for reading and writing Excel files)

You can install these libraries using pip:

pip install pandas openpyxl

Installation

    Download the script:
    Download the consolidate1_1.py file to the directory where you have your Excel files to consolidate.

Usage

    Place your Excel files:
    Ensure all .xlsx files you want to consolidate are in the same directory as the consolidate1_1.py script.

    Configuration (Optional):
    You can adjust the following variables directly in the script if needed:

        INPUT_DIRECTORY: The path where the script will look for Excel files. By default, it's the directory where the script is located.

        OUTPUT_FILENAME: The name of the output Excel file (default: "FINAL_REPORT.xlsx").

        COLUMN_THRESHOLD: The minimum number of non-null values a column must have to not be dropped (default: 4).

        ROW_THRESHOLD: The minimum number of non-null values a row must have to not be dropped (default: 6).

    Execution:
    Open a terminal or command prompt, navigate to the directory where you saved the script and your Excel files, and run:

    python consolidate1_1.py

The script will search for files, process them, and create FINAL_REPORT.xlsx in the same directory.
Troubleshooting

    ModuleNotFoundError: If you see an error like ModuleNotFoundError: No module named 'pandas', it means you haven't installed the required libraries. Run pip install pandas openpyxl.

    Excel files not found: Make sure your .xlsx files are in the same directory as the script. The script will show a warning if it doesn't find any.

    Output file is empty or missing data:

        Check the COLUMN_THRESHOLD and ROW_THRESHOLD values. If they are too high, the script might be dropping too many columns or rows.

        Ensure your Excel files are not corrupted or password-protected.



