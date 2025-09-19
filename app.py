from flask import Flask, render_template, request, redirect, url_for, flash, send_file, session
import os
from werkzeug.utils import secure_filename
from auth import verificar_password, hash_password
from modules.insertar_columna import procesar_excel
from modules.pasar_data import procesar_transferencia, obtener_hojas_analisis
from modules.cambiar_password import cambiar_password_web, generar_hash_password  # ✅ Nuevo import
from datetime import datetime

app = Flask(__name__)
app.secret_key = 'tu_clave_secreta_aqui'
app.config['UPLOAD_FOLDER'] = 'static/uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

@app.route('/cambiar_password', methods=['GET', 'POST'])
def cambiar_password():
    if 'logged_in' not in session:
        return redirect(url_for('login'))

    if request.method == 'POST':
        password_actual = request.form.get('password_actual')
        nueva_password = request.form.get('nueva_password')
        confirmar_password = request.form.get('confirmar_password')

        if not all([password_actual, nueva_password, confirmar_password]):
            flash('Todos los campos son obligatorios', 'error')
            return redirect(request.url)

        # Cambiar contraseña
        resultado, mensaje = cambiar_password_web(password_actual, nueva_password, confirmar_password)

        if resultado:
            flash(mensaje, 'success')
            # Cerrar sesión después de cambiar la contraseña
            session.pop('logged_in', None)
            return redirect(url_for('login'))
        else:
            flash(mensaje, 'error')
            return redirect(request.url)

    return render_template('cambiar_password.html')


@app.route('/')
def index():
    if 'logged_in' not in session:
        return redirect(url_for('login'))
    return render_template('index.html', now=datetime.now())


@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        password = request.form.get('password')
        if verificar_password(password):
            session['logged_in'] = True
            return redirect(url_for('index'))
        else:
            flash('Contraseña incorrecta', 'error')
    return render_template('login.html')


@app.route('/logout')
def logout():
    session.pop('logged_in', None)
    return redirect(url_for('login'))


@app.route('/insertar_columna', methods=['GET', 'POST'])
def insertar_columna():
    if 'logged_in' not in session:
        return redirect(url_for('login'))

    if request.method == 'POST':
        if 'file' not in request.files:
            flash('No se seleccionó ningún archivo', 'error')
            return redirect(request.url)

        file = request.files['file']
        if file.filename == '':
            flash('No se seleccionó ningún archivo', 'error')
            return redirect(request.url)

        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)

            resultado, archivo_procesado, patrones_encontrados = procesar_excel(filepath)

            if resultado:
                return render_template('resultado.html',
                                       exitoso=True,
                                       archivo=filename,
                                       patrones_encontrados=patrones_encontrados,
                                       archivo_descarga=archivo_procesado,
                                       now=datetime.now())
            else:
                flash('Error al procesar el archivo', 'error')
                return redirect(request.url)

    return render_template('insertar_columna.html')


# ✅ NUEVA RUTA PARA TRANSFERIR DATOS
@app.route('/pasar_data', methods=['GET', 'POST'])
def pasar_data():
    if 'logged_in' not in session:
        return redirect(url_for('login'))

    if request.method == 'POST':
        # Verificar contraseña
        password = request.form.get('password')
        if not verificar_password(password):
            flash('Contraseña incorrecta', 'error')
            return redirect(request.url)

        # Verificar archivos
        if 'file_origen' not in request.files or 'file_destino' not in request.files:
            flash('Debe seleccionar ambos archivos', 'error')
            return redirect(request.url)

        file_origen = request.files['file_origen']
        file_destino = request.files['file_destino']

        if file_origen.filename == '' or file_destino.filename == '':
            flash('Debe seleccionar ambos archivos', 'error')
            return redirect(request.url)

        if (file_origen and allowed_file(file_origen.filename) and
                file_destino and allowed_file(file_destino.filename)):

            # Guardar archivos
            filename_origen = secure_filename(file_origen.filename)
            filepath_origen = os.path.join(app.config['UPLOAD_FOLDER'], filename_origen)
            file_origen.save(filepath_origen)

            filename_destino = secure_filename(file_destino.filename)
            filepath_destino = os.path.join(app.config['UPLOAD_FOLDER'], filename_destino)
            file_destino.save(filepath_destino)

            # Obtener hojas disponibles
            try:
                hojas_analisis = obtener_hojas_analisis(filepath_origen)
                hoja_seleccionada = hojas_analisis[0]  # Tomar la primera por defecto

                if len(hojas_analisis) > 1:
                    # Si hay múltiples hojas, usar la seleccionada por el usuario
                    hoja_seleccionada = request.form.get('hoja_analisis', hojas_analisis[0])

            except Exception as e:
                flash(f'Error al leer hojas: {str(e)}', 'error')
                return redirect(request.url)

            # Procesar transferencia
            resultado, mensaje, resumen, archivo_procesado = procesar_transferencia(
                filepath_origen, filepath_destino, hoja_seleccionada, password
            )

            if resultado:
                return render_template('resultado_transferencia.html',
                                       exitoso=True,
                                       mensaje=mensaje,
                                       resumen=resumen,
                                       archivo_descarga=archivo_procesado,
                                       now=datetime.now())
            else:
                flash(mensaje, 'error')
                return redirect(request.url)

    return render_template('pasar_data.html')


@app.route('/descargar/<filename>')
def descargar_archivo(filename):
    if 'logged_in' not in session:
        return redirect(url_for('login'))

    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    return send_file(filepath, as_attachment=True)


def allowed_file(filename):
    return '.' in filename and \
        filename.rsplit('.', 1)[1].lower() in {'xls', 'xlsx', 'xlsm'}


# Filtro personalizado para obtener el nombre base del archivo
@app.template_filter('basename')
def basename_filter(path):
    return os.path.basename(path) if path else ''


if __name__ == '__main__':
    app.run(debug=True)
