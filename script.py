import os
import json
import base64
from datetime import datetime
import pandas as pd
import requests
import sqlalchemy

# -------------------------------------------------------------
# 1. LECTURA DE CONFIGURACIÓN Y SECRETOS
# -------------------------------------------------------------
GITHUB_TOKEN = os.environ.get("GITHUB_TOKEN")
DB_CONFIG_RAW = os.environ.get("DB_CONNECTION_STRING")

if not GITHUB_TOKEN or not DB_CONFIG_RAW:
    raise ValueError("Faltan variables de entorno necesarias (GITHUB_TOKEN o DB_CONNECTION_STRING).")

# Extraer URL de conexión y Nombre de la Tabla del JSON guardado en GitHub Secrets
db_config = json.loads(DB_CONFIG_RAW)
CONNECTION_URL = db_config["connection"]
NOMBRE_TABLA = db_config["tabla"]


def subir_a_github_api(ruta_archivo_local, token, fecha_hora_str):
    """
    Subes/Sobrescribe un archivo directamente en GitHub mediante su API REST (Git Trees/Blobs).
    """
    try:
        print("\n--- Sobrescribiendo archivo en GitHub (vía API REST) ---")
        owner = "usuariocenabast"
        repo = "Resguardo_Excel"
        branch = "main"
        
        nombre_archivo = os.path.basename(ruta_archivo_local)
        path_en_repo = f"Excels Compra historica/{nombre_archivo}"
        
        url_base = f"https://api.github.com/repos/{owner}/{repo}"
        headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/vnd.github+json"
        }

        # 1. Leer el archivo Excel local y pasarlo a base64
        with open(ruta_archivo_local, "rb") as f:
            content_b64 = base64.b64encode(f.read()).decode("utf-8")

        # 2. Crear Blob
        res_blob = requests.post(
            f"{url_base}/git/blobs",
            headers=headers,
            json={"content": content_b64, "encoding": "base64"}
        )
        res_blob.raise_for_status()
        blob_sha = res_blob.json()["sha"]

        # 3. Obtener el commit actual de main
        res_ref = requests.get(f"{url_base}/git/ref/heads/{branch}", headers=headers)
        res_ref.raise_for_status()
        latest_commit_sha = res_ref.json()["object"]["sha"]

        # 4. Crear un árbol con la ruta del archivo actualizada
        res_tree = requests.post(
            f"{url_base}/git/trees",
            headers=headers,
            json={
                "base_tree": latest_commit_sha,
                "tree": [{
                    "path": path_en_repo,
                    "mode": "100644",
                    "type": "blob",
                    "sha": blob_sha
                }]
            }
        )
        res_tree.raise_for_status()
        tree_sha = res_tree.json()["sha"]

        # 5. Crear el commit
        res_commit = requests.post(
            f"{url_base}/git/commits",
            headers=headers,
            json={
                "message": f"Actualización automática sobrescrita: {fecha_hora_str}",
                "tree": tree_sha,
                "parents": [latest_commit_sha]
            }
        )
        res_commit.raise_for_status()
        new_commit_sha = res_commit.json()["sha"]

        # 6. Actualizar la rama main (Push)
        res_push = requests.patch(
            f"{url_base}/git/refs/heads/{branch}",
            headers=headers,
            json={"sha": new_commit_sha}
        )
        res_push.raise_for_status()

        print("¡Éxito! Archivo sobrescrito correctamente en la carpeta del repositorio.")

    except Exception as e:
        print(f"Error al subir a GitHub vía API: {e}")
        raise e

# -------------------------------------------------------------
# 2. EXTRACCIÓN DE DATOS Y GENERACIÓN DEL EXCEL
# -------------------------------------------------------------
if __name__ == "__main__":
    fecha_hora_exacta = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    nombre_base = "Histórico de Compras Cenabast.xlsx"
    ruta_destino = os.path.abspath(nombre_base)

    print(f"Conectando a la BD para consultar la tabla {NOMBRE_TABLA}...")
    engine = sqlalchemy.create_engine(CONNECTION_URL)

    query = f"SELECT * FROM {NOMBRE_TABLA}"
    df = pd.read_sql(query, con=engine)

    # Agregar la columna de fecha y hora de actualización
    df["Fecha y Hora Actualización"] = fecha_hora_exacta

    print("Creando el archivo Excel de salida...")
    with pd.ExcelWriter(ruta_destino, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name="Histórico de Compras", index=False)

    # -------------------------------------------------------------
    # 3. SUBIDA A GITHUB
    # -------------------------------------------------------------
    subir_a_github_api(ruta_destino, GITHUB_TOKEN, fecha_hora_exacta)
