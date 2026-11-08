import os
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for

app = Flask(__name__)

# Ruta relativa segura para el archivo de inventario
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(BASE_DIR, 'inventario.csv')

def cargar_inventario():
    if os.path.exists(DATA_FILE):
        return pd.read_csv(DATA_FILE)
    else:
        # Estructura base si el archivo aún no existe
        df = pd.DataFrame(columns=['Codigo', 'Articulo', 'Cantidad', 'Sucursal'])
        df.to_csv(DATA_FILE, index=False)
        return df

@app.route('/')
def index():
    df = cargar_inventario()
    # Reemplaza o integra esto con tu plantilla HTML actual (ej: render_template('index.html', inventario=df))
    return f"<h1>Sistema de Inventario Yamaha</h1><p>Conectado correctamente. Total de registros: {len(df)}</p>"

if __name__ == '__main__':
    # Puerto dinámico obligatorio para Render
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)