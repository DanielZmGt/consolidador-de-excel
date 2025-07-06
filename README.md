# Consolidador de Excel
## Las instrucciones para correr este script en SPYDER 6 las envio en breve.
Este script de Python (`consolidate1_1.py`) está diseñado para consolidar múltiples hojas de cálculo de Excel (`.xlsx`) de diferentes archivos en un único archivo de salida, organizando los datos por el nombre de la hoja original.

## Requisitos

Para ejecutar este script, necesitas tener instalado lo siguiente:

1.  **Python 3**: Asegúrate de tener Python instalado en tu sistema. Puedes descargarlo desde [python.org](https://www.python.org/downloads/). Durante la instalación en Windows, **asegúrate de marcar la opción "Add Python to PATH"**.
2.  **Módulos de Python**: Los módulos `pandas` y `openpyxl` son esenciales.

## Configuración y Ejecución

Sigue estos pasos para poner en marcha el consolidador:



### 1. Instalar las Dependencias de Python

1.  Abre **PowerShell** (puedes buscarlo en el menú Inicio de Windows).
2.  Navega hasta la carpeta `Excel_Consolidator` usando el comando `cd`. Por ejemplo, si la creaste en tu escritorio:
    ```powersells
    cd "$env:USERPROFILE\Desktop\Excel_Consolidator"
    ```
    (Reemplaza la ruta si tu carpeta está en otro lugar).
3.  Una vez dentro de la carpeta, ejecuta el siguiente comando para instalar los módulos necesarios:
    
    ```powershell
    pip install -r requirements.txt
    ```
    Este comando leerá el archivo `requirements.txt` e instalará `pandas` y `openpyxl`.

### 2. Colocar los Archivos de Excel de Entrada

1.  Coloca todos los archivos `.xlsx` que deseas consolidar **dentro de la misma carpeta `Excel_Consolidator`** donde se encuentra el script.
2.  El script buscará automáticamente todos los archivos `.xlsx` en esta carpeta.

### 3. Ejecutar el Script

1.  Asegúrate de que tu sesión de PowerShell aún esté en la carpeta `Excel_Consolidator`.
2.  Ejecuta el script usando el comando `python` o `py`:
    ```powershell
    python consolidate1_1.py
    ```
    o
    ```powershell
    py consolidate1_1.py
    ```

### 4. Resultado

* Una vez que el script termine de ejecutarse, se creará un nuevo archivo llamado `FINAL_REPORT.xlsx` en la misma carpeta `Excel_Consolidator`.
* Este archivo contendrá todas las hojas consolidadas de tus archivos de entrada, con limpieza básica de filas y columnas según los umbrales definidos en el script.

---

**Notas Importantes:**

* **Umbrales de Limpieza:** El script tiene dos variables `COLUMN_THRESHOLD` y `ROW_THRESHOLD` (definidas al principio del script) que controlan la limpieza de datos. Las columnas con menos de `COLUMN_THRESHOLD` valores no nulos y las filas con menos de `ROW_THRESHOLD` valores no nulos serán eliminadas. Puedes ajustar estos valores directamente en el archivo `consolidate1_1.py` si es necesario.
* **Mensajes en Consola:** El script mostrará mensajes de progreso y posibles advertencias o errores directamente en la consola de PowerShell.
* **Errores Comunes:**
    * `'python' no se reconoce como un comando...`: Asegúrate de haber marcado "Add Python to PATH" durante la instalación de Python, o añádelo manualmente a las variables de entorno.
    * `No se encuentra el archivo...`: Asegúrate de estar en la carpeta correcta en PowerShell o de que los archivos de Excel estén en la misma carpeta que el script.

---