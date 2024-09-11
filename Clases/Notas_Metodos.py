# Genera documentos de evidencias en PDF  para cada caso de prueba de un excel
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

        # Genera casos de pruebas en planilla Excel
        def generar_CP(self, campos_parseados):
            # Crear un nuevo workbook y hoja
            wb = openpyxl.Workbook()
            hoja_casos_prueba = wb.active
            hoja_casos_prueba.title = "Casos de Prueba"

            # Encabezados
            encabezados = ["RF", "Caso de Prueba", "DESCRIPCIÓN", "RESULTADO ESPERADO"]
            hoja_casos_prueba.append(encabezados)
            numCaso = 1
            # Estilo para los encabezados
            for col in range(1, 5):  # De la columna A a D
                celda = hoja_casos_prueba.cell(row=1, column=col)
                celda.font = Font(bold=True)
                celda.alignment = Alignment(horizontal="center")

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
                        numCaso = numCaso + 1
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
