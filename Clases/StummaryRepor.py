import os
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from functions.utilities import format_path_source


class StummaryRepor:
    ##define el nombre del directorio o la ruta completa
    old_dir = "Linea_Base"
    new_dir = "New_Conection"
    diff_dir = "Comparativa"

    ##Define la ruta del archivo excel con las Api a probar
    excel_path = "/Users/equipo2laboratorio/Documents/Api_DUOC_Results/RFC_Validos.xlsx"

    ##define la ruta de los reportes resultantes
    output_path="/Users/equipo2laboratorio/Documents/Api_DUOC_Results/summary_report.pdf"
    excel_path_report = "/Users/equipo2laboratorio/Documents/Api_DUOC_Results/report.xlsx"

    ##Define  responsable de la prueba.
    name_analyst = 'Rckimudwall'
    work_position = 'Analista Técnico-Verity.'

    def __init__(self, old_dir, new_dir, diff_dir, output_path, excel_path, excel_path_report):
        self.old_dir = format_path_source(old_dir)
        self.new_dir = format_path_source(new_dir)
        self.diff_dir = format_path_source(diff_dir)
        self.output_path = format_path_source(output_path)
        self.excel_path = format_path_source(excel_path)
        self.excel_path_report = format_path_source(excel_path_report)
        self.services = []
        self.report_data = []
        self.service_page_numbers = {}
        self.ordered_service_names = self.get_ordered_service_names()

    def get_ordered_service_names(self):
        df = pd.read_excel(self.excel_path)
        return df['service_name'].tolist()

    def scan_directories(self):
        old_files = [f for f in os.listdir(self.old_dir) if f.endswith('_old.txt')]
        for old_file in old_files:
            service_name = old_file.replace('_old.txt', '').strip()
            old_content = self.read_file_content(os.path.join(self.old_dir, old_file))
            new_content = self.read_file_content(os.path.join(self.new_dir, f"{service_name}_new.txt"))
            diff_content = self.read_file_content(os.path.join(self.diff_dir, f"{service_name}_diff.txt"))
            self.add_service(service_name, old_content, new_content, diff_content)

    def read_file_content(self, file_path):
        if os.path.exists(file_path):
            with open(file_path, 'r') as f:
                return f.read().strip()
        return "File not found."

    def add_service(self, service_name, old_content, new_content, diff_content):
        self.services.append({
            "name": service_name,
            "old_content": old_content,
            "new_content": new_content,
            "diff_content": diff_content
        })

    def generate_pdf(self, user_name, user_title):
        # Crear un canvas temporal para calcular las posiciones de las páginas
        temp_canvas = canvas.Canvas(self.output_path, pagesize=letter)
        self.page_num = 1

        # Portada
        self.add_cover_page(temp_canvas, user_name, user_title)

        # Calcular el espacio ocupado por el índice
        num_index_pages = self.calculate_index_pages(temp_canvas)

        # Contenido del Reporte y recolección de números de página
        for service_name in self.ordered_service_names:
            service = next((s for s in self.services if s['name'] == service_name), None)
            if service:
                self.add_service_page(temp_canvas, service)

        # Resumen Final
        self.add_summary_page(temp_canvas)

        temp_canvas.save()

        # Crear el PDF final con el índice correcto
        final_canvas = canvas.Canvas(self.output_path, pagesize=letter)
        self.page_num = 1

        # Portada
        self.add_cover_page(final_canvas, user_name, user_title)

        # Índice con números de página correctos
        self.add_index(final_canvas, num_index_pages)

        # Contenido del Reporte
        for service_name in self.ordered_service_names:
            service = next((s for s in self.services if s['name'] == service_name), None)
            if service:
                self.add_service_page(final_canvas, service)

        # Resumen Final
        self.add_summary_page(final_canvas)

        final_canvas.save()

    def calculate_index_pages(self, c):
        width, height = letter
        c.setFont("Helvetica-Bold", 18)
        c.drawCentredString(width / 2, height - inch, "Indice")
        c.setFont("Helvetica", 12)
        y = height - 2 * inch
        pages = 1
        for i, service_name in enumerate(self.ordered_service_names):
            y -= 0.5 * inch
            if y < inch:
                pages += 1
                y = height - 2 * inch
        return pages

    def add_cover_page(self, c, user_name, user_title):
        width, height = letter
        image_path = os.path.join(os.path.dirname(self.excel_path),
                                  '/Users/equipo2laboratorio/Documents/Duoc_Verity.jpeg')
        if os.path.exists(image_path):
            c.drawImage(image_path, width / 4, height / 2, width / 2, height / 4)
        c.setFont("Helvetica-Bold", 20)
        c.drawCentredString(width / 2, height / 1.4, "DUOC")
        c.drawCentredString(width / 2, height / 1.5, "Proyecto Actualización de Endpoint + AKS / Grupo 1")
        c.setFont("Helvetica", 12)
        c.drawRightString(width - inch, inch, f"{user_name} - {user_title}")
        c.showPage()
        self.page_num += 1

    def add_index(self, c, num_index_pages):
        width, height = letter
        c.setFont("Helvetica-Bold", 18)
        c.drawCentredString(width / 2, height - inch, "Indice")
        c.setFont("Helvetica", 12)
        y = height - 2 * inch
        for i, service_name in enumerate(self.ordered_service_names):
            if service_name in self.service_page_numbers:
                c.drawString(inch, y,
                             f"{i + 1}. {service_name} - Página {self.service_page_numbers[service_name] + num_index_pages}")
            else:
                c.drawString(inch, y, f"{i + 1}. {service_name} - Página No Encontrada")
            y -= 0.5 * inch
            if y < inch:  # Check if we need to start a new page
                c.showPage()
                self.page_num += 1
                c.setFont("Helvetica-Bold", 18)
                c.drawCentredString(width / 2, height - inch, "Indice (continuación)")
                c.setFont("Helvetica", 12)
                y = height - 2 * inch
        c.drawString(inch, y,
                     f"{len(self.ordered_service_names) + 1}. Resumen Final - Página {self.service_page_numbers.get('Resumen Final', 'No Encontrada') + num_index_pages}")
        c.showPage()
        self.page_num += 1

    def add_service_page(self, c, service):
        width, height = letter
        service_name = service['name'].strip()
        self.service_page_numbers[service_name] = self.page_num
        c.setFont("Helvetica-Bold", 16)
        c.drawCentredString(width / 2, height - inch, service_name)
        c.setFont("Helvetica", 8)
        y = height - 2 * inch

        y = self.add_text_block(c, service['old_content'], y, inch, width, height, f"{service_name}_old.txt")
        y = self.add_text_block(c, service['new_content'], y, inch, width, height, f"{service_name}_new.txt")
        y = self.add_text_block(c, service['diff_content'], y, inch, width, height, f"{service_name}_diff.txt")
        c.showPage()
        self.page_num += 1

    def add_text_block(self, c, content, y, x, width, height, filename):
        c.setFont("Helvetica", 12)
        c.drawString(x, y, filename)
        y -= 0.5 * inch

        c.setFont("Helvetica", 10)
        lines = content.split('\n')
        for line in lines:
            if y < inch:
                c.drawRightString(width - inch, inch, f"Página {self.page_num}")
                c.showPage()
                self.page_num += 1
                y = height - inch
                c.setFont("Helvetica", 12)
                c.drawString(x, y, filename)
                y -= 0.5 * inch
                c.setFont("Helvetica", 10)
            c.drawString(x, y, line)
            y -= 0.2 * inch
        c.drawRightString(width - inch, inch, f"Página {self.page_num}")
        return y

    def add_summary_page(self, c):
        width, height = letter
        self.service_page_numbers['Resumen Final'] = self.page_num
        c.setFont("Helvetica-Bold", 16)
        c.drawCentredString(width / 2, height - inch, "Resumen Final")
        c.setFont("Helvetica", 12)
        c.drawRightString(width - inch, height - inch, f"Página {self.page_num}")

        # Ajustar la posición y la distancia entre el título y la tabla
        y = height - 2 * inch

        # Añadir espacio de 2 líneas entre el título y la tabla (ajustando el espacio manualmente)
        y -= 0.5 * inch  # Reduce la distancia manualmente

        # Leer los datos desde el archivo Excel
        df = pd.read_excel('/Users/equipo2laboratorio/Documents/Api_DUOC_Results/report.xlsx')
        table_data = [list(df.columns)] + df.values.tolist()

        # Estilo para los párrafos para ajustar el texto largo
        style = getSampleStyleSheet()["BodyText"]
        for i in range(1, len(table_data)):
            for j in range(len(table_data[i])):
                table_data[i][j] = Paragraph(str(table_data[i][j]), style)

        # Calcular el ancho de las columnas dinámicamente
        num_columns = len(df.columns)
        col_width = (width - 2 * inch) / num_columns

        # Dividir la tabla en partes de 10 filas cada una
        table_parts = [table_data[i:i + 10] for i in range(0, len(table_data), 10)]

        for part in table_parts:
            part_table = Table(part, colWidths=[col_width] * num_columns)
            part_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 6),  # Tamaño de fuente más pequeño para los encabezados
                ('FONTSIZE', (0, 1), (-1, -1), 6),  # Tamaño de fuente más pequeño para el contenido de la tabla
                ('BOTTOMPADDING', (0, 0), (-1, 0), 6),  # Ajustar el relleno inferior para los encabezados
                ('BOTTOMPADDING', (0, 1), (-1, -1), 6),  # Ajustar el relleno inferior para el contenido
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ]))

            # Calcular la posición vertical de la tabla para centrarla
            part_table_height = len(part) * 0.3 * inch  # Ajustar el tamaño de la tabla
            y_start = (height - part_table_height) / 2  # Centrar verticalmente
            part_table.wrapOn(c, width, height)
            part_table.drawOn(c, inch, y_start)

            c.showPage()
            self.page_num += 1
