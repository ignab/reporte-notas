# Reporte de notas en Excel

Proyecto de consola para Python 3.10 o superior. Permite cargar alumnos con
distinta cantidad de notas, organizarlas en un DataFrame de pandas y exportarlas
a `reporte_notas.xlsx` mediante openpyxl. El archivo incluye una tabla con filtros,
encabezados y dos decimales visibles.

## Preparación

Abrí una terminal en esta carpeta y creá un entorno virtual:

```console
python -m venv .venv
```

Activación en Windows (CMD):

```bat
.venv\Scripts\activate
```

Activación en Windows (PowerShell):

```powershell
.\.venv\Scripts\Activate.ps1
```

Activación en Linux o macOS:

```sh
source .venv/bin/activate
```

Instalá las dependencias de la entrega:

```console
python -m pip install -r requirements.txt
```

Para reproducir los pasos originales del ejercicio desde cero:

```console
python -m pip install pandas openpyxl
python -m pip freeze > requirements.txt
```

El archivo incluido contiene las versiones congeladas del entorno utilizado.
Si PowerShell bloquea la activación, podés usar CMD o ejecutar directamente
`.venv\Scripts\python.exe reporte_notas.py`, sin activar el entorno.

## Ejecución

```console
python reporte_notas.py
```

1. Ingresá el nombre del alumno.
2. Ingresá sus notas separadas por espacios o punto y coma: `7 8,5 6`.
3. Repetí la carga y escribí `fin` como nombre para terminar.
4. Revisá el resumen en la consola y el Excel en esta misma carpeta.

Ejemplo de entradas, una por línea:

```text
Ana
7 8,5 6
Luis
4;5
fin
```

Ana tendrá promedio 7,17 y condición Aprobado; Luis, 4,50 y Desaprobado.

## Criterios y validaciones

- Escala inclusiva de 0 a 10. Se acepta punto o coma decimal; la coma no separa
  notas. No se aceptan valores infinitos, NaN ni texto en las notas.
- Se aprueba con promedio mayor o igual a 6. Las constantes al comienzo del
  programa permiten identificar el criterio; el ejercicio adopta esta escala.
- La condición se calcula con el promedio sin redondear. Excel muestra dos
  decimales; ese formato no modifica el valor utilizado para decidir la condición.
- Se exige nombre no vacío (máximo 100 caracteres) y al menos una nota por alumno.
- Los nombres repetidos se conservan como registros independientes.
- Las notas faltantes de alumnos con menos evaluaciones quedan vacías en Excel
  y no se cuentan como ceros al calcular su promedio.
- `fin` no distingue mayúsculas y minúsculas y está reservado para finalizar.
- Una lista vacía no genera archivo. Los datos incorrectos se vuelven a pedir.
- Ctrl+C o fin de entrada cancelan la carga sin guardar datos parciales.
- Si ya hay un reporte, se pide confirmación antes de reemplazarlo.
- El guardado usa un archivo temporal y reemplazo final para proteger un reporte
  anterior si falla la escritura. Los errores de permisos, archivo abierto,
  almacenamiento y dependencias ausentes se comunican por consola.

Se manejan las excepciones previsibles y existe una barrera final para fallos
inesperados. Ningún programa puede recuperarse de un corte de energía o de la
terminación forzada de su proceso.

## Archivos

- `reporte_notas.py`: programa con funciones documentadas y anotaciones de tipos.
- `requirements.txt`: dependencias directas y transitivas del entorno.
- `README.md`: preparación, ejecución y criterios.
- `.gitignore`: excluye entorno virtual, caché y reportes generados.

La carpeta `.venv` se crea localmente y no se distribuye: sus rutas dependen de
cada equipo. El programa devuelve código 0 al finalizar correctamente y 1 ante
un error o una carga interrumpida.
