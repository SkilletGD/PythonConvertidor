# Importamos las bibliotecas necesarias para este proyecto
from flask import Flask, request, render_template, send_file
import pandas as pd
import docx
import os


def crear_app():
    # Creamos una instancia para la aplicacion Flask
    app = Flask(__name__)

    # Definimos la ruta donde se guardaran los archivos subidos
    CARPETA_CARGA = 'uploads'

    # Creamos una carpeta de uploads si no existe
    if not os.path.exists(CARPETA_CARGA):
        os.makedirs(CARPETA_CARGA)

    # Ruta para la cual mandamos a llamar al archivo html
    @app.route('/')
    def index():
        return render_template('index.html')

    # Ruta para subir archivos, solo solicitudes POST
    @app.route('/upload', methods=['POST'])
    def subir_archivos():
        # Verificamos si se ha enviado un archivo en la solicitud
        if 'file' not in request.files:
            mensaje_error = 'No hay Archivo'
            tipo_mensaje = 'danger'
            return render_template('index.html', mensaje=mensaje_error, categoria=tipo_mensaje)
        
        # Obtenemos el archivo subido
        archivo = request.files['file']

        # Verificamos si se ha seleccionado un archivo
        if archivo.filename == '':
            mensaje_error = 'Archivo no seleccionado'
            tipo_mensaje = 'danger'
            return render_template('index.html', mensaje=mensaje_error, categoria=tipo_mensaje)

        # Verificamos si el archivo es un archivo docx
        if not archivo.filename.lower().endswith('.docx'):
            mensaje_error = 'Archivo no compatible por el momento, por favor cargue un archivo docx.'
            tipo_mensaje = 'danger'
            return render_template('index.html', mensaje=mensaje_error, categoria=tipo_mensaje)

        # Si todo es válido, guardamos el archivo en la carpeta de uploads
        if archivo:
                ruta_archivo = os.path.join(CARPETA_CARGA, 'input.docx')
                archivo.save(ruta_archivo)

                # Convertimos el archivo de docx a csv despues de subirlo
                ruta_salida = os.path.join(CARPETA_CARGA, 'Reporte.csv')
                try:
                    # Convertimos el archivo a csv con la funcion convert_docx_to_csv
                    convert_docx_to_csv(ruta_archivo, ruta_salida)
                except Exception as e:
                    # Por si nos sale un error podremos imprimir cual a sido el error
                    mensaje_error = f'Error al convertir el archivo: {e}'
                    tipo_mensaje = 'danger'

                # Devolvemos un mensaje de éxito y un enlace para descargar el archivo
                return render_template('index.html', mensaje='¡Archivo convertido con exito!', descargar=True, archivo=ruta_salida)
        
        # Si no se ha seleccionado un archivo, devolvemos un mensaje de error
        return render_template('index.html', mensaje='No se ha seleccionado un archivo', categoria='danger')
        
    # Hacemos una ruta para descargar el archivo convetido
    @app.route('/descargar/<path:archivo>')
    def descargar(archivo):
        return send_file(archivo, as_attachment=True)

    # Se actualizo esta parte del codigo debido a que pandas dejo de funcionar con 
    # df.append desde versiones anteriores, la mejor adaptacion fue df.loc
    def addReporte(df, movimientos, solicitud):
        for reg in range(1,len(movimientos)):
            df.loc[len(df)] = {
                'ORDEN': solicitud,
                'TIEMPO DE INICIO': movimientos[reg][1] + " " + movimientos[reg][2],
                'ACTIVIDAD': movimientos[reg][0],
                'RESPONSABLE': movimientos[reg][4],
                'OBSERVACIONES': movimientos[reg][6]
            }
        return df

    def convert_docx_to_csv(input_path, output_path):
        archivo = docx.Document(input_path)
        tablas = archivo.tables
        parrafos = archivo.paragraphs

        reporte = pd.DataFrame(columns=['ORDEN', 'TIEMPO DE INICIO', 'ACTIVIDAD', 'RESPONSABLE', 'OBSERVACIONES'])
        indP = indT = indB = 0
        indTmovimientos = 0
        solicitud = ''
        for parrafo in archivo.paragraphs:
            if (parrafo.text == ''):
                indP += 1
                continue
            if (parrafo.text == 'INFORMACIÓN GENERAL'):
                solicitud = [[cell.text for cell in row.cells] for row in tablas[indT].rows][0][1]
            if (parrafo.text == 'BITACORA DE MOVIMIENTOS'):
                indTmovimientos += 1
                movimientos = [[cell.text for cell in row.cells] for row in tablas[indB].rows]
                reporte = addReporte(reporte, movimientos, solicitud)
            indP += 1
            indT += 1
            indB += 1
        print('Se encontraron ' + str(indTmovimientos) + 'tablas de movimientos.')
        reporte.to_csv(output_path, index=False)
    return app

if __name__ == '__main__':
    app = crear_app()
    app.run()