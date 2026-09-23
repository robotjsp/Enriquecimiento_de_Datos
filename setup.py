"""
setup.py
--------
Define el proyecto como un paquete instalable y declara sus dependencias.
No ejecuta la lógica de ingesta/limpieza/enriquecimiento (eso lo hacen los
scripts en src/); su función es dejar claro qué librerías necesita el
proyecto para funcionar, tanto en instalación local como en GitHub Actions.
"""

from setuptools import setup, find_packages  # setup: función que registra el proyecto / find_packages: detecta paquetes automáticamente

setup(
    name="ingestion-bigdata",           # Nombre del proyecto/paquete
    version="1.2.0",                    # Versión del proyecto (incrementada: EA1 + EA2 + EA3)
    description="Ingesta, limpieza y enriquecimiento de datos del API de Colombia",  # Descripción corta
    packages=find_packages(where="src"),  # Busca automáticamente subpaquetes dentro de src/ (si los hubiera)
    package_dir={"": "src"},              # Indica que el código fuente vive dentro de la carpeta src/
    install_requires=[                    # Lista de dependencias necesarias para correr el proyecto
        "requests>=2.31.0",              # Para hacer las peticiones HTTP al API (EA1)
        "pandas>=2.0.0",                 # Para procesar los datos en todas las etapas (EA1, EA2, EA3)
        "openpyxl>=3.1.0",               # Motor que usa Pandas para leer/escribir archivos .xlsx
        "lxml>=5.0.0",                   # Motor que usa pandas.read_html para parsear la fuente HTML (EA3)
    ],
    python_requires=">=3.10",             # Versión mínima de Python requerida
)
