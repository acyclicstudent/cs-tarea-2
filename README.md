# cs-tarea-2
Tarea 2 Computer Security ESCOM

En este repo se encuentra el script de la tarea 2, el cuál tiene el objetivo de buscar en un directorio archivos:
- docx
- pdf
- xlsx
Con la finalidad de extraer su metadata y mostrarla en consola.

Para ejecutar el script:
```
// Ingresar a la carpeta ráiz del repositorio
cd cs-tarea-2
// Instalar dependencias
python -m venv .env && source .env/bin/activate && pip install -r requirements.txt
// Ejecutar el script
python main.py RUTA_DIRECTORIO
```

El script recibe los siguientes argumentos:
- RUTA_DIRECTORIO: Dirección relativa o absoluta, en caso de que la ruta tenga espacios se ingresa entre comillas.

Nota: El script no soporta PDF's con cifrado.