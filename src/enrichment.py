"""
enrichment.py
-------------
Script de enriquecimiento de datos para la EA3 del proyecto integrador.

Simula un entorno de procesamiento en la nube: carga el dataset limpio generado
en la EA2 (src/xlsx/cleaned_data.xlsx) y lo enriquece cruzándolo con 6 fuentes
adicionales en distintos formatos (JSON, CSV, XLSX, XML, HTML, TXT), todas
ubicadas en src/fuentes/.

Estrategia de cruce:
- El dataset base solo trae el nombre de la ciudad, no el nombre del departamento
  en texto. Por eso, primero se INFIERE el departamento a partir de la ciudad,
  usando un diccionario de mapeo ciudad -> departamento (basado en la división
  político-administrativa real de Colombia).
- Con el departamento ya inferido, se hace merge() contra 5 de las 6 fuentes
  (todas descritas por departamento).
- La sexta fuente (categorias.html) se cruza por texto: se buscan palabras clave
  del nombre/descripción de la atracción para asignarle una categoría turística.

Flujo del script:
1. Carga el dataset limpio de la EA2.
2. Infiere el departamento de cada atracción a partir de su ciudad.
3. Lee las 6 fuentes adicionales (una función de lectura por formato).
4. Enriquece el dataset con las 5 fuentes basadas en departamento (merge).
5. Enriquece el dataset con la fuente de categorías (cruce por palabra clave).
6. Exporta una muestra del dataset enriquecido a Excel.
7. Genera un archivo de auditoría .txt con el detalle del proceso de integración.
"""

# --- Importación de librerías ---
import pandas as pd                              # Para manipular los datos en DataFrames
import os                                         # Para manejar rutas y creación de carpetas
import json                                       # Para leer la fuente en formato JSON
import xml.etree.ElementTree as ET                # Para leer la fuente en formato XML (librería nativa de Python)
from datetime import datetime                     # Para registrar la fecha/hora de ejecución en el reporte

# --- Rutas del proyecto (relativas a la carpeta src/) ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))                              # Carpeta donde vive este script (src/)
CLEANED_XLSX_PATH = os.path.join(BASE_DIR, "xlsx", "cleaned_data.xlsx")            # Dataset base, generado en la EA2
FUENTES_DIR = os.path.join(BASE_DIR, "fuentes")                                    # Carpeta con las 6 fuentes adicionales
ENRICHED_XLSX_PATH = os.path.join(BASE_DIR, "xlsx", "enriched_data.xlsx")          # Salida: muestra del dataset enriquecido
REPORT_PATH = os.path.join(BASE_DIR, "static", "auditoria", "enrichment_report.txt")  # Salida: reporte de auditoría

# --- Diccionario de mapeo ciudad -> departamento ---
# Cubre las ciudades más comunes que suelen aparecer como sede de atracciones
# turísticas en el API. Las ciudades no incluidas quedarán como "Desconocido"
# y ese caso se documenta explícitamente en el reporte de auditoría.
CIUDAD_A_DEPARTAMENTO = {
    "Bogotá D.C.": "Bogotá D.C.", "Bogota D.c.": "Bogotá D.C.", "Bogotá": "Bogotá D.C.",
    "Medellín": "Antioquia", "Medellin": "Antioquia", "Guatapé": "Antioquia", "Guatape": "Antioquia",
    "Cali": "Valle Del Cauca", "Buenaventura": "Valle Del Cauca",
    "Cartagena": "Bolívar", "Cartagena De Indias": "Bolívar",
    "Santa Marta": "Magdalena",
    "Barranquilla": "Atlántico",
    "Bucaramanga": "Santander",
    "Cúcuta": "Norte De Santander", "Cucuta": "Norte De Santander",
    "Tunja": "Boyacá", "Villa De Leyva": "Boyacá",
    "Manizales": "Caldas",
    "Pereira": "Risaralda",
    "Armenia": "Quindío", "Salento": "Quindío",
    "Ibagué": "Tolima", "Ibague": "Tolima",
    "Neiva": "Huila", "San Agustín": "Huila", "San Agustin": "Huila",
    "Popayán": "Cauca", "Popayan": "Cauca",
    "Pasto": "Nariño",
    "Quibdó": "Chocó", "Quibdo": "Chocó", "Nuquí": "Chocó", "Nuqui": "Chocó",
    "Valledupar": "Cesar",
    "Montería": "Córdoba", "Monteria": "Córdoba",
    "Sincelejo": "Sucre",
    "Riohacha": "La Guajira",
    "San Andrés": "San Andrés Y Providencia", "San Andres": "San Andrés Y Providencia",
    "Villavicencio": "Meta",
    "Yopal": "Casanare",
    "Arauca": "Arauca",
    "Leticia": "Amazonas",
    "Florencia": "Caquetá",
    "Mocoa": "Putumayo",
}


def cargar_dataset_base(ruta_excel: str) -> pd.DataFrame:
    """
    Carga el dataset limpio generado en la EA2 (cleaned_data.xlsx).
    Este archivo hace las veces de "fuente en la nube" ya procesada.
    """
    df = pd.read_excel(ruta_excel)                         # Lee el Excel de datos limpios en un DataFrame
    print(f"[OK] Dataset base cargado: {len(df)} registros desde {ruta_excel}")  # Log informativo
    return df                                              # Devuelve el DataFrame base


def inferir_departamento(df: pd.DataFrame, reporte: dict) -> pd.DataFrame:
    """
    Agrega una columna 'departamento' al DataFrame, inferida a partir del nombre
    de la ciudad usando el diccionario CIUDAD_A_DEPARTAMENTO. Las ciudades sin
    mapeo conocido quedan marcadas como 'Desconocido'.
    """
    df["departamento"] = df["ciudad"].map(CIUDAD_A_DEPARTAMENTO).fillna("Desconocido")  # Aplica el mapeo, o "Desconocido" si no está

    sin_departamento = int((df["departamento"] == "Desconocido").sum())    # Cuenta cuántas filas quedaron sin departamento
    reporte["registros_sin_departamento_inferido"] = sin_departamento       # Guarda el conteo en el reporte
    reporte["ciudades_sin_mapeo"] = sorted(                                 # Lista de ciudades que no se pudieron mapear
        df.loc[df["departamento"] == "Desconocido", "ciudad"].unique().tolist()
    )
    print(f"[OK] Departamento inferido para {len(df) - sin_departamento} de {len(df)} registros.")  # Log
    return df                                                                # Devuelve el DataFrame con la columna nueva


def leer_json(ruta: str) -> pd.DataFrame:
    """Lee la fuente de región natural en formato JSON."""
    with open(ruta, "r", encoding="utf-8") as archivo:   # Abre el archivo JSON en modo lectura
        datos = json.load(archivo)                        # Parsea el contenido JSON a una lista de diccionarios
    df = pd.DataFrame(datos)                              # Convierte la lista de diccionarios en un DataFrame
    print(f"[OK] JSON leído: {len(df)} filas desde {os.path.basename(ruta)}")  # Log informativo
    return df                                              # Devuelve el DataFrame con columnas: departamento, region_natural


def leer_csv(ruta: str) -> pd.DataFrame:
    """Lee la fuente de capitales y superficie en formato CSV."""
    df = pd.read_csv(ruta)                                 # Pandas lee el CSV directamente a un DataFrame
    print(f"[OK] CSV leído: {len(df)} filas desde {os.path.basename(ruta)}")   # Log informativo
    return df                                              # Devuelve el DataFrame con columnas: departamento, capital, superficie_km2


def leer_xlsx(ruta: str) -> pd.DataFrame:
    """Lee la fuente de clima predominante en formato XLSX."""
    df = pd.read_excel(ruta)                               # Pandas lee el Excel directamente a un DataFrame
    print(f"[OK] XLSX leído: {len(df)} filas desde {os.path.basename(ruta)}")  # Log informativo
    return df                                              # Devuelve el DataFrame con columnas: departamento, clima_predominante


def leer_xml(ruta: str) -> pd.DataFrame:
    """Lee la fuente de lenguas indígenas en formato XML usando ElementTree."""
    arbol = ET.parse(ruta)                                 # Parsea el archivo XML y construye el árbol de nodos
    raiz = arbol.getroot()                                 # Obtiene el nodo raíz (<departamentos>)

    filas = []                                             # Lista donde se acumulan los registros extraídos
    for nodo_depto in raiz.findall("departamento"):        # Itera sobre cada nodo <departamento>
        nombre_depto = nodo_depto.get("nombre")            # Lee el atributo "nombre" del nodo
        lengua = nodo_depto.find("lengua_indigena").text   # Lee el texto del hijo <lengua_indigena>
        filas.append({"departamento": nombre_depto, "lengua_indigena": lengua})  # Agrega el registro a la lista

    df = pd.DataFrame(filas)                               # Convierte la lista de registros en un DataFrame
    print(f"[OK] XML leído: {len(df)} filas desde {os.path.basename(ruta)}")    # Log informativo
    return df                                              # Devuelve el DataFrame con columnas: departamento, lengua_indigena


def leer_html(ruta: str) -> pd.DataFrame:
    """Lee la fuente de categorías turísticas en formato HTML (tabla) usando pandas.read_html."""
    tablas = pd.read_html(ruta)                            # Extrae todas las tablas <table> encontradas en el HTML
    df = tablas[0]                                         # Toma la primera (y única) tabla del archivo
    print(f"[OK] HTML leído: {len(df)} filas desde {os.path.basename(ruta)}")   # Log informativo
    return df                                              # Devuelve el DataFrame con columnas: palabra_clave, categoria_turistica


def leer_txt(ruta: str) -> pd.DataFrame:
    """Lee la fuente de festividades en formato TXT, con parseo manual línea por línea."""
    filas = []                                             # Lista donde se acumulan los registros extraídos
    with open(ruta, "r", encoding="utf-8") as archivo:     # Abre el archivo de texto en modo lectura
        for linea in archivo:                              # Itera línea por línea sobre el archivo
            linea = linea.strip()                          # Quita espacios y saltos de línea sobrantes
            if not linea or linea.startswith("#"):         # Ignora líneas vacías o comentarios (empiezan con #)
                continue                                    # Salta a la siguiente línea del archivo
            departamento, festividad = linea.split(":", 1)  # Separa "departamento: festividad" por el primer ":"
            filas.append({                                  # Agrega el registro ya limpio a la lista
                "departamento": departamento.strip(),
                "festividad_representativa": festividad.strip(),
            })
    df = pd.DataFrame(filas)                               # Convierte la lista de registros en un DataFrame
    print(f"[OK] TXT leído: {len(df)} filas desde {os.path.basename(ruta)}")    # Log informativo
    return df                                              # Devuelve el DataFrame con columnas: departamento, festividad_representativa


def enriquecer_por_departamento(df: pd.DataFrame, fuente: pd.DataFrame, nombre_fuente: str, reporte: dict) -> pd.DataFrame:
    """
    Hace un merge (left join) entre el DataFrame principal y una fuente adicional,
    usando 'departamento' como clave. Registra en el reporte cuántos registros
    del dataset principal encontraron coincidencia en esa fuente.
    """
    columnas_antes = set(df.columns)                        # Guarda las columnas actuales, para detectar las nuevas
    df = df.merge(fuente, on="departamento", how="left")     # Une por departamento, conservando todas las filas del dataset principal
    columnas_nuevas = list(set(df.columns) - columnas_antes)  # Detecta qué columnas trajo esta fuente

    if columnas_nuevas:                                       # Si la fuente aportó al menos una columna nueva
        coincidencias = int(df[columnas_nuevas[0]].notna().sum())  # Cuenta cuántas filas quedaron con dato (no nulo) en esa columna
    else:
        coincidencias = 0                                      # Caso borde: la fuente no aportó columnas nuevas

    reporte["fuentes"].append({                                # Agrega el detalle de esta fuente al reporte
        "nombre_fuente": nombre_fuente,
        "registros_en_fuente": len(fuente),
        "registros_coincidentes": coincidencias,
        "registros_sin_coincidencia": len(df) - coincidencias,
    })
    print(f"[OK] Enriquecido con {nombre_fuente}: {coincidencias}/{len(df)} coincidencias.")  # Log informativo
    return df                                                  # Devuelve el DataFrame ya enriquecido con esta fuente


def asignar_categoria_por_texto(df: pd.DataFrame, tabla_categorias: pd.DataFrame, reporte: dict) -> pd.DataFrame:
    """
    Asigna una categoría turística a cada atracción, buscando palabras clave
    de 'tabla_categorias' dentro del nombre o la descripción de la atracción.
    Si ninguna palabra clave coincide, se asigna la categoría 'Sin categorizar'.
    """
    texto_busqueda = (df["nombre"].fillna("") + " " + df["descripcion"].fillna("")).str.lower()  # Une nombre+descripción en minúsculas

    def buscar_categoria(texto: str) -> str:
        """Recorre la tabla de categorías y devuelve la primera coincidencia encontrada en el texto."""
        for _, fila in tabla_categorias.iterrows():           # Itera sobre cada palabra clave de la tabla
            palabra_clave = str(fila["palabra_clave"]).lower()  # Normaliza la palabra clave a minúsculas
            if palabra_clave in texto:                          # Si la palabra clave aparece dentro del texto
                return fila["categoria_turistica"]               # Devuelve la categoría asociada a esa palabra clave
        return "Sin categorizar"                                 # Si no hubo ninguna coincidencia, devuelve este valor por defecto

    df["categoria_turistica"] = texto_busqueda.apply(buscar_categoria)  # Aplica la búsqueda fila por fila

    sin_categoria = int((df["categoria_turistica"] == "Sin categorizar").sum())  # Cuenta cuántas quedaron sin categoría
    reporte["fuentes"].append({                                 # Agrega el detalle de esta fuente al reporte
        "nombre_fuente": "categorias.html",
        "registros_en_fuente": len(tabla_categorias),
        "registros_coincidentes": len(df) - sin_categoria,
        "registros_sin_coincidencia": sin_categoria,
    })
    print(f"[OK] Categoría textual asignada: {len(df) - sin_categoria}/{len(df)} con categoría reconocida.")  # Log
    return df                                                    # Devuelve el DataFrame con la columna categoria_turistica


def exportar_dataset_enriquecido(df: pd.DataFrame, ruta_salida: str, n: int = 50) -> None:
    """Exporta una muestra representativa del dataset enriquecido a un archivo Excel."""
    os.makedirs(os.path.dirname(ruta_salida), exist_ok=True)  # Crea la carpeta xlsx/ si no existe
    tamano_muestra = min(n, len(df))                            # Evita pedir más filas de las que existen
    muestra = df.sample(n=tamano_muestra, random_state=42) if tamano_muestra > 0 else df  # Toma una muestra aleatoria reproducible
    muestra.to_excel(ruta_salida, index=False)                  # Exporta la muestra a Excel sin la columna de índice
    print(f"[OK] Dataset enriquecido exportado a: {ruta_salida}")  # Log informativo


def generar_reporte_enriquecimiento(reporte: dict, registros_base: int, registros_finales: int, ruta_txt: str) -> None:
    """
    Escribe el archivo de auditoría .txt documentando el proceso de integración:
    cuántos registros coincidieron con cada fuente y observaciones relevantes.
    """
    os.makedirs(os.path.dirname(ruta_txt), exist_ok=True)      # Crea la carpeta static/auditoria/ si no existe

    lineas = [                                                  # Construye el contenido del reporte línea por línea
        "REPORTE DE AUDITORÍA - ENRIQUECIMIENTO DE DATOS",
        f"Fecha y hora de ejecución: {datetime.now().isoformat()}",
        "=" * 70,
        f"Registros del dataset base (EA2): {registros_base}",
        f"Registros del dataset enriquecido (EA3): {registros_finales}",
        f"Registros sin departamento inferido: {reporte['registros_sin_departamento_inferido']}",
        f"Ciudades sin mapeo a departamento: {reporte['ciudades_sin_mapeo'] if reporte['ciudades_sin_mapeo'] else 'Ninguna'}",
        "=" * 70,
        "DETALLE DE INTEGRACIÓN POR FUENTE",
    ]

    for fuente in reporte["fuentes"]:                            # Recorre cada fuente integrada y agrega su detalle
        lineas.append(f"  Fuente: {fuente['nombre_fuente']}")
        lineas.append(f"    Registros en la fuente: {fuente['registros_en_fuente']}")
        lineas.append(f"    Registros coincidentes: {fuente['registros_coincidentes']}")
        lineas.append(f"    Registros sin coincidencia: {fuente['registros_sin_coincidencia']}")

    lineas += [
        "=" * 70,
        "Conclusión: El dataset fue enriquecido exitosamente con información "
        "geográfica, climática, cultural y de categorización turística proveniente "
        "de 6 fuentes en formatos distintos (JSON, CSV, XLSX, XML, HTML, TXT). "
        "Los registros sin coincidencia corresponden a ciudades sin mapeo de "
        "departamento definido o a atracciones sin palabras clave reconocidas.",
    ]

    with open(ruta_txt, "w", encoding="utf-8") as archivo:      # Abre el archivo de auditoría en modo escritura
        archivo.write("\n".join(lineas))                        # Escribe todas las líneas separadas por saltos de línea

    print(f"[OK] Reporte de enriquecimiento generado en: {ruta_txt}")  # Log informativo


def main():
    """
    Orquesta el pipeline completo de enriquecimiento:
    cargar base -> inferir departamento -> leer 6 fuentes -> enriquecer (merge/texto)
    -> exportar -> generar reporte de auditoría.
    """
    reporte = {"fuentes": []}                                    # Diccionario acumulador de estadísticas para el reporte

    df = cargar_dataset_base(CLEANED_XLSX_PATH)                  # Paso 1: carga el dataset limpio de la EA2
    registros_base = len(df)                                     # Guarda la cantidad de registros base, antes de enriquecer

    df = inferir_departamento(df, reporte)                       # Paso 2: infiere el departamento a partir de la ciudad

    df_regiones = leer_json(os.path.join(FUENTES_DIR, "regiones.json"))              # Paso 3a: lee fuente JSON
    df_capitales = leer_csv(os.path.join(FUENTES_DIR, "capitales.csv"))              # Paso 3b: lee fuente CSV
    df_clima = leer_xlsx(os.path.join(FUENTES_DIR, "clima.xlsx"))                    # Paso 3c: lee fuente XLSX
    df_idiomas = leer_xml(os.path.join(FUENTES_DIR, "idiomas.xml"))                  # Paso 3d: lee fuente XML
    df_categorias = leer_html(os.path.join(FUENTES_DIR, "categorias.html"))          # Paso 3e: lee fuente HTML
    df_festivos = leer_txt(os.path.join(FUENTES_DIR, "festivos_relevantes.txt"))     # Paso 3f: lee fuente TXT

    df = enriquecer_por_departamento(df, df_regiones, "regiones.json", reporte)      # Paso 4a: enriquece con región natural
    df = enriquecer_por_departamento(df, df_capitales, "capitales.csv", reporte)     # Paso 4b: enriquece con capital/superficie
    df = enriquecer_por_departamento(df, df_clima, "clima.xlsx", reporte)            # Paso 4c: enriquece con clima
    df = enriquecer_por_departamento(df, df_idiomas, "idiomas.xml", reporte)         # Paso 4d: enriquece con lengua indígena
    df = enriquecer_por_departamento(df, df_festivos, "festivos_relevantes.txt", reporte)  # Paso 4e: enriquece con festividad

    df = asignar_categoria_por_texto(df, df_categorias, reporte)  # Paso 5: enriquece con categoría turística (cruce por texto)

    exportar_dataset_enriquecido(df, ENRICHED_XLSX_PATH, n=50)              # Paso 6: exporta muestra del dataset enriquecido
    generar_reporte_enriquecimiento(reporte, registros_base, len(df), REPORT_PATH)  # Paso 7: genera el archivo de auditoría

    print(f"[OK] Pipeline de enriquecimiento finalizado. Registros finales: {len(df)}")  # Log final


if __name__ == "__main__":    # Verifica que el script se ejecute directamente (no como import)
    main()                     # Llama a la función principal para correr todo el pipeline
