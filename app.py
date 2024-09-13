from flask import Flask, render_template, request, redirect, url_for, send_file
import pandas as pd

app = Flask(__name__)

# Nombre del archivo Excel
EXCEL_FILE = 'datos.xlsx'

# Cargar datos desde el archivo Excel al iniciar
def cargar_datos():
    try:
        df = pd.read_excel(EXCEL_FILE)
        return df.to_dict(orient='records')  # Convierte el DataFrame en una lista de diccionarios
    except FileNotFoundError:
        # Si no se encuentra el archivo, inicializa una lista vacía
        return []

# Guardar datos en el archivo Excel
def guardar_datos():
    df = pd.DataFrame(datos)
    df.to_excel(EXCEL_FILE, index=False)

# Inicializa los datos
datos = cargar_datos()

# Ruta principal para mostrar los datos y el formulario
@app.route('/')
def index():
    return render_template('index.html', datos=datos)

# Ruta para crear un nuevo dato
@app.route('/crear', methods=['POST'])
def crear():
    nuevo_dato = {
        "id": len(datos) + 1,
        "dia": request.form['dia'],
        "fecha": request.form['fecha'],
        "unidad_numero": request.form['unidad_numero'],
        "curso": request.form['curso'],
        "caracter_clase": request.form['caracter_clase'],
        "contenidos_tematicos": request.form['contenidos_tematicos'],
        "actividades": request.form['actividades'],
        "observaciones": request.form['observaciones']
    }
    datos.append(nuevo_dato)
    guardar_datos()  # Guarda en el archivo Excel después de agregar
    return redirect(url_for('index'))

@app.route('/actualizar/<int:id>', methods=['POST'])
def actualizar(id):
    for dato in datos:
        if dato['id'] == id:
            dato.update({
                "dia": request.form.get('dia', dato.get('dia', '')),
                "fecha": request.form.get('fecha', dato.get('fecha', '')),
                "unidad_numero": request.form.get('unidad_numero', dato.get('unidad_numero', '')),
                "curso": request.form.get('curso', dato.get('curso', '')),
                "caracter_clase": request.form.get('caracter_clase', dato.get('caracter_clase', '')),
                "contenidos_tematicos": request.form.get('contenidos_tematicos', dato.get('contenidos_tematicos', '')),
                "actividades": request.form.get('actividades', dato.get('actividades', '')),
                "observaciones": request.form.get('observaciones', dato.get('observaciones', ''))
            })
            break
    guardar_datos()  # Guarda en el archivo Excel después de actualizar
    return redirect(url_for('index'))

# Ruta para eliminar un dato
@app.route('/eliminar/<int:id>', methods=['POST'])
def eliminar(id):
    global datos
    datos = [dato for dato in datos if dato['id'] != id]
    guardar_datos()  # Guarda en el archivo Excel después de eliminar
    return redirect(url_for('index'))

# Ruta para descargar el archivo Excel
@app.route('/descargar_excel')
def descargar_excel():
    return send_file(EXCEL_FILE, as_attachment=True)

if __name__ == "__main__":
    app.run(app.run(host='0.0.0.0'), debug=True)