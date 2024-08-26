import os

import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Font
from Clases.Validador import Validador
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Image, PageBreak, Spacer
from PIL import Image as PILImage  # Para obtener las dimensiones originales de la imagen



class Controlador:
    def __init__(self, ruta_txt, ruta_excel, ruta_excel_salida, ruta_pdf_salida):
        self.ruta_txt = ruta_txt
        self.ruta_excel = ruta_excel
        self.ruta_excel_salida = ruta_excel_salida
        self.ruta_pdf_salida = ruta_pdf_salida
        self.lineas_txt = []
        self.tipo_caracteres = []
        self.requerido = []
        self.largo = []
        self.inicio = []
        self.id_validacion = []
        self.nombre_campo = []
        self.detalleValidacion = []

    def leer_archivo_txt(self):
        try:
            with open(self.ruta_txt, 'r', encoding='utf-8') as archivo:
                self.lineas_txt = archivo.readlines()
            self.lineas_txt = [linea.strip() for linea in self.lineas_txt]
            print("Archivo TXT leído correctamente.")
        except FileNotFoundError:
            print(f"Error: El archivo {self.ruta_txt} no se encontró.")
        except Exception as e:
            print(f"Se produjo un error al leer el archivo TXT: {e}")

    def leer_archivo_excel(self):
        try:
            df = pd.read_excel(self.ruta_excel, engine='openpyxl')
            print("Primeras filas del DataFrame:")

            # Asegúrate de que los índices sean correctos
            self.nombre_campo = df.iloc[:, 1].tolist()  # Columna B (índice 1)
            self.tipo_caracteres = df.iloc[:, 3].tolist()  # Columna G (índice 3)
            self.requerido = df.iloc[:, 4].tolist()  # Columna H (índice 4)
            self.largo = df.iloc[:, 5].tolist()  # Columna I (índice 5)
            self.inicio = df.iloc[:, 6].tolist()  # Columna J (índice 6)
            self.id_validacion = df.iloc[:, 8].tolist()  # Columna I (índice 8)
            self.detalleValidacion = df.iloc[:, 9].tolist()  # Columna J (índice 9)

            print("Archivo Excel leído correctamente.")
        except FileNotFoundError:
            print(f"Error: El archivo {self.ruta_excel} no se encontró.")
        except Exception as e:
            print(f"Se produjo un error al leer el archivo Excel: {e}")

    def obtener_datos_txt(self):
        return self.lineas_txt

    def obtener_datos_excel(self):
        return {
            'tipo_caracteres': self.tipo_caracteres,
            'requerido': self.requerido,
            'largo': self.largo,
            'inicio': self.inicio,
            'id_validacion': self.id_validacion,
            'nombre_campo': self.nombre_campo
        }

    def extraer_subcadenas(self, cadena, inicio, largo):
        """
        Extrae subcadenas de la cadena principal usando las posiciones de inicio y los largos proporcionados.

        :param cadena: La cadena principal de la cual extraer subcadenas.
        :param inicio: Lista de posiciones de inicio para cada subcadena.
        :param largo: Lista de largos de subcadenas a extraer.
        :return: Lista de subcadenas extraídas.
        """
        subcadenas = []

        # Verificar que las listas de inicio y largo tengan la misma longitud
        if len(inicio) != len(largo):
            raise ValueError("Las listas de inicio y largo deben tener la misma longitud.")

        for i in range(len(inicio)):
            start = inicio[i] - 1
            length = largo[i]

            # Asegurarse de que el rango de extracción esté dentro de los límites de la cadena
            if start < 0 or start >= len(cadena):
                subcadenas.append("")  # Agregar una cadena vacía si el índice de inicio está fuera del rango
                continue

            end = min(start + length, len(cadena))  # Ajustar el final si excede la longitud de la cadena
            subcadenas.append(cadena[start:end])

        return subcadenas

    def validar_campos(self):

        reportes = []
        # Suponiendo que vamos a validar la primera línea del archivo TXT
        list_tipo = self.tipo_caracteres
        list_requerido = self.requerido
        list_id_validacion = self.id_validacion
        list_nombre_campo = self.nombre_campo

        for n in range(len(self.lineas_txt)):
            registro1 = self.lineas_txt[n] if self.lineas_txt else ""
            # Dividir el registro en subcadenas
            registro_parseado = self.extraer_subcadenas(registro1, self.inicio, self.largo)
            reporte = []
            for i in range(len(registro_parseado)):
                validador = Validador(
                    registro_parseado[i],
                    list_tipo[i],
                    list_requerido[i],
                    list_id_validacion[i],
                    list_nombre_campo[i],
                    registro_parseado
                )
                # Ejecutar todas las validaciones
                es_valido, mensaje = validador.validar_todo(registro_parseado[i])
                if not es_valido:
                    reporte.append(mensaje)
            reportes.append(reporte)

        return reportes

    def generar_excel(self, campos_parseados):
        # Crear un nuevo workbook y hoja
        wb = openpyxl.Workbook()
        hoja_casos_prueba = wb.active
        hoja_casos_prueba.title = "Casos de Prueba"

        # Encabezados
        encabezados = ["RF", "Caso de Prueba", "DESCRIPCIÓN", "RESULTADO ESPERADO"]
        hoja_casos_prueba.append(encabezados)

        # Estilo para los encabezados
        for col in range(1, 5):  # De la columna A a D
            celda = hoja_casos_prueba.cell(row=1, column=col)
            celda.font = Font(bold=True)
            celda.alignment = Alignment(horizontal="center")
        numCaso = 1

        # Procesar los datos para agregar las filas
        for i, campo in enumerate(campos_parseados):  # Asegúrate de estar iterando con índices
            if self.tipo_caracteres[i] == "Alfa":
                tipo = "Alfanumérico"
            else:
                tipo = "Numérico"
            rf = "RF-01-Trama 8650"

            # Extraer información relevante de cada fila cp1 del campo
            caso_de_prueba = f"CP 0{numCaso} - Campo {self.nombre_campo[i]} - Largo del Campo"
            numCaso = numCaso + 1
            descripcion = "\n".join([
                f"Paso 1: Definir paso 1",
                f"Paso 2: Definir paso 2",
                f"Paso 3: Definir paso 3",
                f"Paso n: Validar que el campo tenga un largo de {self.largo[i]} caracteres."
            ])
            resultado_esperado = "El campo debe cumplir con el requisito de validación."
            hoja_casos_prueba.append([rf, caso_de_prueba, descripcion, resultado_esperado])
            # Extraer información relevante de cada fila cp2 del campo
            caso_de_prueba = f"CP 0{numCaso} - Campo {self.nombre_campo[i]} - Tipo de dato"
            numCaso = numCaso + 1
            descripcion = "\n".join([
                f"Paso 1: Definir paso 1",
                f"Paso 2: Definir paso 2",
                f"Paso 3: Definir paso 3",
                f"Paso n: Validar que el campo sea {tipo}."
            ])
            hoja_casos_prueba.append([rf, caso_de_prueba, descripcion, resultado_esperado])
            if self.requerido[i] == 1:
                # Extraer información relevante de cada fila cp3 del campo
                caso_de_prueba = f"CP 0{numCaso} - Campo {self.nombre_campo[i]} - Requerido"
                numCaso = numCaso + 1
                descripcion = "\n".join([
                    f"Paso 1: Definir paso 1",
                    f"Paso 2: Definir paso 2",
                    f"Paso 3: Definir paso 3",
                    f"Paso n: Validar que el campo es Obligatorio (No puede ser nulo)."
                ])
                hoja_casos_prueba.append([rf, caso_de_prueba, descripcion, resultado_esperado])
                if self.id_validacion[i] != 0:
                    # Extraer información relevante de cada fila cp4 del campo
                    caso_de_prueba = f"CP 0{numCaso} - Campo {self.nombre_campo[i]} - Lógica"
                    numCaso = numCaso + 1
                    descripcion = "\n".join([
                        f"Paso 1: Definir paso 1",
                        f"Paso 2: Definir paso 2",
                        f"Paso 3: Definir paso 3",
                        f"Paso n: Validar que el campo tenga la siguiente lógica: {self.detalleValidacion[i]}."
                    ])
                    hoja_casos_prueba.append([rf, caso_de_prueba, descripcion, resultado_esperado])
            else:
                if self.id_validacion[i] != 0:
                    # Extraer información relevante de cada fila cp4 del campo
                    caso_de_prueba = f"CP 0{numCaso} - Campo {self.nombre_campo[i]} - Lógica"
                    numCaso = numCaso + 1 = numCaso + 1
                    descripcion = "\n".join([
                        f"Paso 1: Definir paso 1",
                        f"Paso 2: Definir paso 2",
                        f"Paso 3: Definir paso 3",
                        f"Paso n: Validar que el campo tenga la siguiente lógica: {self.detalleValidacion[i]}."
                    ])
                    hoja_casos_prueba.append([rf, caso_de_prueba, descripcion, resultado_esperado])

        # Guardar el archivo Excel
        wb.save(self.ruta_excel_salida)
        print(f"Excel guardado en {self.ruta_excel_salida} con {numCaso} casos de prueba")

    def generar_evidencia_pdf(self, ruta_excel):
        # Cargar el archivo Excel
        wb = openpyxl.load_workbook(ruta_excel)
        hoja = wb.active

        # Iterar sobre las filas del Excel (a partir de la segunda fila para saltar los encabezados)
        for i, fila in enumerate(hoja.iter_rows(min_row=2, values_only=True), start=1):
            # Extraer los datos necesarios de cada fila del Excel
            rf = fila[0]  # Columna A es el RF
            nombre_caso_prueba = fila[1]  # Columna B es el nombre del caso de prueba
            descripcion_caso = fila[2]  # Columna C es la descripción del caso
            resultado_esperado = fila[3]  # Columna D es el resultado esperado

            # Crear un archivo PDF por cada caso de prueba
            nombre_archivo = f"Evidencia_CP_{i}.pdf"
            ruta_completa = os.path.join(self.ruta_pdf_salida, nombre_archivo)
            pdf = SimpleDocTemplate(ruta_completa, pagesize=letter)

            # Crear una lista para almacenar los elementos del PDF
            elementos = []

            # Crear estilos
            styles = getSampleStyleSheet()
            estilo_titulo = styles['Title']
            estilo_titulo.alignment = 1  # Centrar el texto

            # Primera página - Título, imagen y tabla
            # Ruta de la imagen
            image_path = r'C:\Users\Usuario\PycharmProjects\TramasNexus\imagenes\Imagen1.png'

            # Obtener dimensiones originales de la imagen
            with PILImage.open(image_path) as img:
                original_width, original_height = img.size

            # Definir ancho máximo deseado para la imagen en el PDF
            max_width = 4 * inch  # Ajusta el ancho máximo que desees
            # Calcular el alto manteniendo la proporción de la imagen
            aspect_ratio = original_height / original_width
            adjusted_height = max_width * aspect_ratio

            # Insertar la imagen con las dimensiones ajustadas
            img = Image(image_path, width=max_width, height=adjusted_height)
            img.hAlign = 'CENTER'
            elementos.append(Spacer(1, 1 * inch))  # Espacio superior
            elementos.append(img)

            # Espaciado entre imagen y texto
            elementos.append(Spacer(1, 0.5 * inch))

            # Títulos en el documento
            elementos.append(Paragraph("DOCUMENTO DE EVIDENCIAS", estilo_titulo))
            elementos.append(Spacer(1, 0.3 * inch))  # Espaciado entre títulos
            elementos.append(Paragraph("PROYECTO: NXS-264 - Ciclo 1- Balance de Cartera", estilo_titulo))

            # Espaciado entre título y tabla
            elementos.append(Spacer(1, 0.5 * inch))

            # Crear la tabla con el índice vertical
            datos_tabla = [
                ["ID Caso de Prueba", rf],
                ["Caso de Prueba", nombre_caso_prueba],
                ["Ciclo de Certificación", "2"],
                ["Estado de Ejecución", "OK"],
                ["Ambiente", "VP-TNR"],
                ["Emisor", "661"],
                ["Ingeniero de pruebas", "Jose Alejandro Perez"]
            ]

            # Crear y estilizar la tabla
            tabla = Table(datos_tabla, colWidths=[150, 300])
            tabla.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (0, -1), colors.grey),
                ('TEXTCOLOR', (0, 0), (0, -1), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 12),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (1, 0), (1, -1), colors.beige),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ]))
            elementos.append(tabla)

            # Agregar un salto de página
            elementos.append(PageBreak())

            # Segunda página - Descripción y Resultado
            elementos.append(Paragraph("Descripción del Caso", styles['Heading2']))
            elementos.append(Paragraph(descripcion_caso, styles['BodyText']))
            elementos.append(Paragraph("Resultado Obtenido:", styles['Heading2']))
            elementos.append(Paragraph(resultado_esperado, styles['BodyText']))

            # Crear el PDF
            pdf.build(elementos)
            print(f"PDF generado: {nombre_archivo}")

    # Llama a este método después de generar el Excel
    def generar_excel_y_evidencia(self, campos_parseados, ruta_excel_salida):
        self.generar_excel(campos_parseados)
        self.generar_evidencia_pdf(ruta_excel_salida)

# Ejemplo de uso:
ruta_txt = r"C:\Users\Usuario\PycharmProjects\TramasNexus\archivos\MPLUS 8650 QA.txt"
ruta_excel = r"C:\Users\Usuario\PycharmProjects\TramasNexus\archivos\Informacion Trama.xlsx"
ruta_excel_salida = r"C:\Users\Usuario\PycharmProjects\TramasNexus\archivos\Casos de Prueba.xlsx"
ruta_pdf_salida = "C:/Users/Usuario/PycharmProjects/TramasNexus/archivos/"

controlador = Controlador(ruta_txt, ruta_excel, ruta_excel_salida, ruta_pdf_salida)
controlador.leer_archivo_txt()
controlador.leer_archivo_excel()

lineas_txt = controlador.obtener_datos_txt()
datos_excel = controlador.obtener_datos_excel()

# Ejemplo de cadena y listas de inicio y largo
cadena_ejemplo = lineas_txt[0]  # Suponiendo que quieres procesar la primera línea del archivo TXT
inicio = datos_excel['inicio']
largo = datos_excel['largo']

# Llamar al método extraer_subcadenas
subcadenas = controlador.extraer_subcadenas(cadena_ejemplo, inicio, largo)

controlador.generar_excel(subcadenas)
controlador.generar_evidencia_pdf(ruta_excel_salida)

