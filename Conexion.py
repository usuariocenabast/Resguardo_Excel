import os
import base64
from datetime import datetime
import pandas as pd
import requests
import sqlalchemy

# -------------------------------------------------------------
# CONFIGURACIÓN VÍA VARIABLES DE ENTORNO (SECRETS DE GITHUB)
# -------------------------------------------------------------
GITHUB_TOKEN = os.environ.get("GITHUB_TOKEN")
DB_CONNECTION_STRING = os.environ.get("CONEXION")

def subir_a_github_api(ruta_archivo_local, token, fecha_hora_str):
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

        with open(ruta_archivo_local, "rb") as f:
            content_b64 = base64.b64encode(f.read()).decode("utf-8")

        # 1. Crear el blob
        res_blob = requests.post(
            f"{url_base}/git/blobs",
            headers=headers,
            json={"content": content_b64, "encoding": "base64"}
        )
        res_blob.raise_for_status()
        blob_sha = res_blob.json()["sha"]

        # 2. Obtener el commit actual
        res_ref = requests.get(f"{url_base}/git/ref/heads/{branch}", headers=headers)
        res_ref.raise_for_status()
        latest_commit_sha = res_ref.json()["object"]["sha"]

        # 3. Crear el nuevo árbol
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

        # 4. Crear el commit
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

        # 5. Actualizar la rama (Push)
        res_push = requests.patch(
            f"{url_base}/git/refs/heads/{branch}",
            headers=headers,
            json={"sha": new_commit_sha}
        )
        res_push.raise_for_status()

        print("¡Éxito! Archivo sobrescrito correctamente en GitHub.")

    except Exception as e:
        print(f"Error al subir a GitHub vía API: {e}")
        raise e

# -------------------------------------------------------------
# EJECUCIÓN
# -------------------------------------------------------------
fecha_hora_exacta = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
nombre_base = "Histórico de Compras Cenabast.xlsx"
ruta_destino = os.path.abspath(nombre_base)

# 1. Consulta SQL directa a la base de datos remota
engine = sqlalchemy.create_engine(DB_CONNECTION_STRING)

query = """
    SELECT 
        * 
    FROM 
        tabla_historico_compras
"""

print("Conectando a la BD en la nube y ejecutando consulta...")
df = pd.read_sql(query, con=engine)

# 2. Agregar la columna de fecha y hora
df["Fecha y Hora Actualización"] = fecha_hora_exacta

# 3. Guardar el archivo Excel limpio
print("Creando archivo Excel...")
with pd.ExcelWriter(ruta_destino, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name="Histórico de Compras", index=False)

# 4. Subir a GitHub
subir_a_github_api(ruta_destino, GITHUB_TOKEN, fecha_hora_exacta)