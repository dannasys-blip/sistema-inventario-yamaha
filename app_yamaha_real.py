import os
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for

app = Flask(__name__)

# Ruta relativa segura para el archivo de inventario
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(BASE_DIR, 'inventario.csv')

def cargar_inventario():
    if os.path.exists(DATA_FILE):
        df = pd.read_csv(DATA_FILE)
        if df.empty:
            df = pd.DataFrame(columns=['Codigo', 'Articulo', 'Cantidad', 'Sucursal'])
            df.to_csv(DATA_FILE, index=False)
        return df
    else:
        # Estructura base con columnas requeridas si el archivo no existe
        df = pd.DataFrame(columns=['Codigo', 'Articulo', 'Cantidad', 'Sucursal'])
        df.to_csv(DATA_FILE, index=False)
        return df

@app.route('/')
def index():
    df = cargar_inventario()
    # Convierte el DataFrame a registros para mostrarlos en la plantilla HTML
    registros = df.to_dict(orient='records')
    return render_template('index.html', inventario=registros)

@app.route('/agregar', methods=['POST'])
def agregar():
    codigo = request.form.get('codigo')
    articulo = request.form.get('articulo')
    cantidad = request.form.get('cantidad')
    sucursal = request.form.get('sucursal')
    
    df = cargar_inventario()
    
    # Crear un nuevo registro y agregarlo usando Pandas
    nuevo_reg = pd.DataFrame([{'Codigo': codigo, 'Articulo': articulo, 'Cantidad': cantidad, 'Sucursal': sucursal}])
    df = pd.concat([df, nuevo_reg], ignore_index=True)
    df.to_csv(DATA_FILE, index=False)
    
    return redirect(url_for('index'))

if __name__ == '__main__':
    # Puerto dinámico obligatorio para Render
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)