# EA1 + EA2 + EA3 - Ingesta, Limpieza y Enriquecimiento de Datos

## Descripción breve de la solución

Este proyecto implementa las etapas de **ingesta (EA1)**, **preprocesamiento/limpieza (EA2)** y **enriquecimiento (EA3)** del proyecto integrador de Big Data, usando como fuente principal el API público de **[API Colombia](https://api-colombia.com/)** (endpoint `TouristicAttraction`).

### EA1 - Ingesta (`src/ingestion.py`)

1. Se conecta al API y extrae la lista de atracciones turísticas del país.
2. Almacena la información en una base de datos **SQLite** (`src/db/ingestion.db`), en dos tablas relacionadas: `ciudades` y `atracciones`.
3. Genera un archivo **Excel** (`src/xlsx/ingestion.xlsx`) con una muestra representativa de los registros.
4. Genera un archivo de **auditoría** (`src/static/auditoria/ingestion.txt`) que compara lo extraído del API contra lo almacenado en la base de datos.

### EA2 - Preprocesamiento y limpieza (`src/cleaning.py`)

Simula un entorno de nube: lee `src/db/ingestion.db` (generada en la EA1) y aplica:
- Eliminación de duplicados exactos (por `id`) y lógicos (mismo nombre + ciudad).
- Corrección de tipos de datos (`latitud`/`longitud` a `float`, `poblacion` a entero).
- Manejo de nulos (relleno de descripciones vacías, imputación de población por mediana, eliminación de filas sin coordenadas).
- Eliminación de outliers geográficos (coordenadas fuera de Colombia).
- Transformaciones adicionales: normalización de texto, `num_imagenes`, escalado min-max de `poblacion_normalizada`.

Genera `src/xlsx/cleaned_data.xlsx` y `src/static/auditoria/cleaning_report.txt`.

### EA3 - Enriquecimiento (`src/enrichment.py`)

Carga el dataset limpio de la EA2 y lo enriquece con **6 fuentes adicionales en distintos formatos**, ubicadas en `src/fuentes/`:

| Fuente | Formato | Información que aporta |
|---|---|---|
| `regiones.json` | JSON | Región natural de Colombia (Andina, Caribe, Pacífica, Orinoquía, Amazonía, Insular) |
| `capitales.csv` | CSV | Capital y superficie (km²) del departamento |
| `clima.xlsx` | XLSX | Clima predominante del departamento |
| `idiomas.xml` | XML | Lengua indígena representativa del departamento |
| `categorias.html` | HTML | Tabla de palabras clave → categoría turística (Playa y Costa, Montaña y Naturaleza, Patrimonio y Cultura, Recreación y Aventura) |
| `festivos_relevantes.txt` | TXT | Festividad representativa del departamento |

**Estrategia de cruce:** el dataset base solo trae el nombre de la *ciudad*, no el del *departamento* en texto. Por eso `enrichment.py` primero **infiere el departamento a partir de la ciudad**, usando un diccionario de mapeo basado en la división político-administrativa real de Colombia. Con esa columna ya calculada, se hace `merge()` contra 5 de las 6 fuentes (todas descritas por departamento). La sexta fuente (`categorias.html`) se cruza por **texto**: se buscan palabras clave dentro del nombre/descripción de cada atracción para asignarle una categoría turística.

Genera `src/xlsx/enriched_data.xlsx` y `src/static/auditoria/enrichment_report.txt`, este último con el detalle de cuántos registros coincidieron con cada una de las 6 fuentes.

> **Nota:** las ciudades que no están en el diccionario de mapeo quedan marcadas como "Desconocido" y se documentan explícitamente en el reporte de auditoría, en vez de generar un cruce incorrecto.

Todo el proceso está automatizado mediante **GitHub Actions**, que ejecuta el pipeline completo (ingesta → limpieza → enriquecimiento) y deja evidencia como artefactos descargables y commiteados en el repositorio.

## Estructura del proyecto

```
├── setup.py                              # Declara el paquete y sus dependencias
├── README.md                             # Este archivo
├── .github/workflows/bigdata.yml         # Workflow de automatización (EA1 + EA2 + EA3)
└── src/
    ├── ingestion.py                      # EA1: script de ingesta desde el API
    ├── cleaning.py                       # EA2: script de limpieza y preprocesamiento
    ├── enrichment.py                     # EA3: script de enriquecimiento
    ├── fuentes/                          # EA3: las 6 fuentes adicionales
    │   ├── regiones.json
    │   ├── capitales.csv
    │   ├── clima.xlsx
    │   ├── idiomas.xml
    │   ├── categorias.html
    │   └── festivos_relevantes.txt
    ├── db/ingestion.db                   # Base de datos SQLite (salida EA1 / entrada EA2, simula la nube)
    ├── xlsx/
    │   ├── ingestion.xlsx                # EA1: muestra de datos crudos
    │   ├── cleaned_data.xlsx             # EA2: muestra de datos limpios
    │   └── enriched_data.xlsx            # EA3: muestra de datos enriquecidos
    └── static/auditoria/
        ├── ingestion.txt                 # EA1: reporte de auditoría de ingesta
        ├── cleaning_report.txt           # EA2: reporte de auditoría de limpieza
        └── enrichment_report.txt         # EA3: reporte de auditoría de enriquecimiento
```

> **Nota sobre el nombre del script de la EA3:** el enunciado sugiere `enrichement.py` (con un error tipográfico). En este repositorio se usó el nombre correctamente escrito, `enrichment.py`.

## Instrucciones para clonar y ejecutar localmente

```bash
# 1. Clonar el repositorio
git clone https://github.com/<tu-usuario>/<tu-repo>.git
cd <tu-repo>

# 2. Instalar dependencias (usa setup.py)
pip install -e .

# 3. Ejecutar el pipeline completo, en orden
python src/ingestion.py      # EA1: genera db/ingestion.db
python src/cleaning.py       # EA2: lee la BD, genera cleaned_data.xlsx
python src/enrichment.py     # EA3: lee cleaned_data.xlsx, genera enriched_data.xlsx
```

Al finalizar, se generarán/actualizarán automáticamente los 7 archivos de evidencia listados en la estructura del proyecto.

## Automatización con GitHub Actions

El archivo `.github/workflows/bigdata.yml` define un workflow que:

1. Se dispara manualmente (`workflow_dispatch`), en cada `push` a `main`, o de forma programada (`cron` diario).
2. Instala Python 3.11 y las dependencias del proyecto (`requests`, `pandas`, `openpyxl`, `lxml`) vía `pip install -e .`.
3. Ejecuta en orden `ingestion.py` (EA1), `cleaning.py` (EA2) y `enrichment.py` (EA3).
4. Verifica que los 7 archivos de evidencia existan y muestra el contenido del reporte de enriquecimiento en el log.
5. Sube todos los archivos generados como **artefacto descargable** (`evidencias-proyecto`) desde la pestaña **Actions**.
6. Hace commit y push de los archivos generados directamente al repositorio, dejando evidencia permanente del proceso.

### Cómo verificar la ejecución

1. Ir a la pestaña **Actions** del repositorio.
2. Seleccionar la ejecución más reciente del workflow "Ingesta, Limpieza y Enriquecimiento - Proyecto Big Data".
3. Revisar los logs de cada paso (ingesta, limpieza, enriquecimiento, verificación de archivos).
4. Descargar el artefacto **evidencias-proyecto** para inspeccionar la base de datos, los tres Excel y los tres reportes.
5. Alternativamente, revisar directamente en el repositorio las rutas `src/db/`, `src/xlsx/` y `src/static/auditoria/`, actualizadas automáticamente por el workflow.
