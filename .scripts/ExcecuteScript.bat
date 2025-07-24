@echo off
REM Cambia al directorio donde está el script (opcional pero recomendable)
cd /d "C:\Users\se109874\OneDrive - Repsol\Documentos\UiPath\Reporte-Regulatorio_Refacturacion\.scripts"

REM Ejecuta el script con Python
python "get_tdc_banxico.py"

if %errorlevel% neq 0 (
    echo Hubo un error al ejecutar el script Python. Código de error: %errorlevel% > "C:\Users\se109874\OneDrive - Repsol\Documentos\UiPath\Reporte-Regulatorio_Refacturacion\.scripts\output.txt"
)