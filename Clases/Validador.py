import self


class Validador:
    def __init__(self, campo, tipo_valor, es_obligatorio, id_validacion, nombre_campo, campos_parseados):
        self.campo = campo
        self.tipo_valor = tipo_valor  # "Alfa" o "Num"
        self.es_obligatorio = es_obligatorio  # 0 o 1
        self.id_validacion = id_validacion  # id que indica el tipo de validación por la que pasará
        self.nombre_campo = nombre_campo  # Nombre del campo
        self.campos_parseados = campos_parseados  # todos los campos
        self.errores = []

    def validar_todo(self, valor):
        """
        Ejecuta todas las validaciones sobre el valor.
        :param valor: Valor a validar.
        :return: Tupla (valido, mensaje) donde 'valido' es un booleano y 'mensaje' proporciona detalles sobre la validación.
        """
        # Validar si es obligatorio
        es_valido, mensaje = self.obligatorio(valor)
        if not es_valido:
            return es_valido, mensaje

        # Validar si es numérico
        es_valido, mensaje = self.es_numerico(valor)
        if not es_valido:
            return es_valido, mensaje

        # Validaciones de tipo conjunto Ids empiezan con 2, conjuntos
        if str(self.id_validacion).startswith(("2", "3")):
            es_valido, mensaje = self.validar_valores_permitidos(valor)
            if not es_valido:
                return es_valido, mensaje

        # Validaciones de tipo conjunto Ids empiezan con 3, dependencia con otros campos
        if str(self.id_validacion).startswith("3"):
            es_valido, mensaje = self.validar_dependencia(valor)
            if not es_valido:
                return es_valido, mensaje

        # Validaciones de tipo conjunto Ids empiezan con 4, no puede contener
        if str(self.id_validacion).startswith("4"):
            es_valido, mensaje = self.validar_no_contiene(valor)
            if not es_valido:
                return es_valido, mensaje

        # Validaciones de tipo conjunto Ids empiezan con 5, primeros numeros de tarjeta
        if str(self.id_validacion).startswith("5"):
            es_valido, mensaje = self.validar_inicio_tarjeta(valor)
            if not es_valido:
                return es_valido, mensaje

        # Puedes agregar más validaciones aquí si es necesario

        return True, "ok"

    def obligatorio(self, valor):
        """
        Verifica si el valor es obligatorio (es_obligatorio = 1) y si está vacío.
        :param valor: Valor a verificar.
        :return: Tupla (valido, mensaje) donde 'valido' es un booleano y 'mensaje' proporciona detalles sobre la validación.
        """
        if self.es_obligatorio == 1:
            if not valor.strip():  # Verifica si el valor está vacío o solo contiene espacios
                error = {
                    'nombre_campo': self.nombre_campo,
                    'valor': valor,  # Asegúrate de pasar el valor correcto aquí
                    'tipo_error': 'Campo Obligatorio',
                    'detalle_error': 'El campo no puede ser nulo o vacío.'
                }
                self.errores.append(error)
                return False, error  # Solo devuelve el error actual, no la lista completa de errores
        return True, "ok"

    def es_numerico(self, valor):
        """
        Verifica si el valor es numérico según el tipo especificado (Num).
        :param valor: Valor a verificar.
        :return: True si el valor es numérico, False en caso contrario.
        """
        # Si el valor está vacío, no se considera error para un campo numérico
        if valor.strip() == "":
            return True, "ok"

        if self.tipo_valor == "Num":
            if not valor.isdigit():
                error = {
                    'nombre_campo': self.nombre_campo,
                    'valor': self.campo,
                    'tipo_error': 'Campo Numérico',
                    'detalle_error': f'El campo no puede contener letras u otros caracteres no numéricos.'
                }
                self.errores.append(error)
                return False, error  # Solo devuelve el error actual, no la lista completa de errores
        return True, "ok"

    def validar_valores_permitidos(self, valor):
        """
        Verifica si el valor está dentro de un conjunto predefinido de valores permitidos cuando dependiendo del id_validacion.
        :param valor: Valor a verificar.
        :return: Tupla (valido, mensaje) donde 'valido' es un booleano y 'mensaje' proporciona detalles sobre la validación.
        """
        # Limpiar espacios en blanco al inicio y al final del valor
        valor_limpio = valor.strip()
        mensaje = f'Los valores permitidos para el campo son:'
        valores_permitidos = {}
        if self.id_validacion == 21:
            valores_permitidos = {"0100", "0200", "0120", "0220", "0400", "0420", "0302"}
        elif self.id_validacion == 22:
            valores_permitidos = {"1000", "2000", "3000", "4000"}
        elif self.id_validacion == 23:
            valores_permitidos = {"N"}
            mensaje = f'El valor permitido para el campo es:'
        elif self.id_validacion == 24:
            valores_permitidos = {"D", "V", "C", "B", "E", "N"}
        elif self.id_validacion == 25:
            valores_permitidos = {"ATM", "POS"}
        elif self.id_validacion == 26:
            valores_permitidos = {"Cr", "De", "Pp"}
        elif self.id_validacion == 27:
            valores_permitidos = {"K", "C", "M"}
        elif self.id_validacion == 28:
            valores_permitidos = {"VIA", "BAA", "TBA", "TVT", "TMT"}
        elif self.id_validacion == 29:
            valores_permitidos = {"G", "D", "F", "P"}
        elif self.id_validacion == 210:
            valores_permitidos = {"620", "700", "710", "720", "750"}
        elif self.id_validacion == 214:
            valores_permitidos = {"ByteIN"}
            mensaje = f'El valor permitido para el campo es:'
        elif self.id_validacion == 213:
            valores_permitidos = {"08650"}
            mensaje = f'El valor permitido para el campo es:'
        elif self.id_validacion == 215:
            valores_permitidos = {"  "}
            mensaje = f'El valor permitido para el campo es:'
        elif self.id_validacion == 216:
            valores_permitidos = {"  ", "R "}
        elif self.id_validacion == 217:
            valores_permitidos = {"   ", "1  "}
        elif self.id_validacion == 220:
            valores_permitidos = {"S", "N"}
        elif self.id_validacion == 221:
            valores_permitidos = {"S", "M"}
        elif self.id_validacion == 31:
            valores_permitidos = {"S", "N"}
        elif self.id_validacion == 32:
            valores_permitidos = {"000002", "000001"}

        # Permitir valores vacíos
        if valor_limpio == "":
            return True, "ok"

        if valor not in valores_permitidos:
            error = {
                'nombre_campo': self.nombre_campo,
                'valor': self.campo,
                'tipo_error': 'Valores permitidos',
                'detalle_error': f'{mensaje} {valores_permitidos}.'
            }
            self.errores.append(error)
            return False, error  # Solo devuelve el error actual, no la lista completa de errores

        return True, "ok"

    def validar_dependencia(self, valor):
        """
        Valida el campo_b basado en el contenido del campo_a.

        - Si campo_a es "0400" o "0420", campo_b debe ser "S".
        - En cualquier otro caso, campo_b debe ser "N".
        """
        # Limpiar espacios en blanco al inicio y al final del valor
        valor_limpio = valor.strip()

        # Inicializar pos como None
        pos = None

        # Configuración de validación basada en id_validacion
        if self.id_validacion == 31:
            valores_validos_campo_a = {"0400", "0420"}
            campo_bsi = "S"
            campo_bno = "N"
            pos = 5  # Índice del campo en campos_parseados
        elif self.id_validacion == 32:
            valores_validos_campo_a = {"0220", "0120"}
            campo_bsi = "000002"
            campo_bno = "000001"
            pos = 5  # Ajusta el índice según sea necesario

        # Verificar si pos sigue siendo None
        if pos is None:
            return False, "Error: id_validacion no es válido o no se asignó el índice correctamente."

        # Permitir valores vacíos
        if valor_limpio == "":
            return True, "ok"

        # Verificar que el índice pos exista en campos_parseados
        if pos < len(self.campos_parseados):
            campo_a = self.campos_parseados[pos]
        else:
            return False, f"Error: El índice {pos} está fuera de rango en campos_parseados."

        # Realiza la validación
        if campo_a in valores_validos_campo_a:
            if valor == campo_bno:
                error = {
                    'nombre_campo': self.nombre_campo,
                    'valor': self.campo,
                    'tipo_error': 'Dependencia de campos',
                    'detalle_error': f"El campo debe ser '{campo_bsi}' si el campo en la posición {pos+1}: '{campo_a}' es {valores_validos_campo_a}."
                }
                self.errores.append(error)
                return False, error  # Solo devuelve el error actual, no la lista completa de errores
            return True, "ok"
        elif valor == campo_bsi:
            error = {
                'nombre_campo': self.nombre_campo,
                'valor': self.campo,
                'tipo_error': 'Dependencia de campos',
                'detalle_error': f"El campo debe ser '{campo_bno}' si el campo en la posición {pos+1}: '{campo_a}' no es {valores_validos_campo_a}."
            }
            self.errores.append(error)
            return False, error  # Solo devuelve el error actual, no la lista completa de errores

        return True, "ok"

    def validar_no_contiene(self, valor):
        """
        Verifica que el valor no contenga las subcadenas 'ByteIN' o 'ByteF'.

        :param valor: La cadena que se desea validar.
        :return: (bool, str) - True si la cadena es válida, False si contiene subcadenas no permitidas, y el mensaje de validación.
        """
        subcadenas_no_permitidas = ["ByteIN", "ByteF"]

        for subcadena in subcadenas_no_permitidas:
            if subcadena in valor:
                error = {
                    'nombre_campo': self.nombre_campo,
                    'valor': self.campo,
                    'tipo_error': 'Excluir Valor',
                    'detalle_error': f"El campo no puede contener la subcadena: '{subcadena}'."
                }
                self.errores.append(error)
                return False, error  # Solo devuelve el error actual, no la lista completa de errores

        return True, "ok"

    def validar_inicio_tarjeta(self, valor):
        # Verificar si los primeros 8 caracteres de campo2 son iguales a campo1
        cant = None
        campo = None
        if self.id_validacion == 51:
            cant = 8
            campo = 6
            mensaje = f"El campo tiene que ser igual a los primeros '{cant}' dígitos de la tarjeta campo {campo}: '{self.campos_parseados[campo][:cant]}'."
        elif self.id_validacion == 52:
            cant = 6
            campo = 6
            mensaje = f"El campo tiene que ser igual a los primeros '{cant}' dígitos de la tarjeta campo {campo}: '{self.campos_parseados[campo][:cant]}'."
        elif self.id_validacion == 53:
            cant = 3
            campo = 24
            mensaje = f"El campo tiene que ser igual al campo {campo}: '{self.campos_parseados[campo][:cant]}'."
        if valor == self.campos_parseados[campo][:cant]:
            return True, "ok"
        else:
            error = {
                'nombre_campo': self.nombre_campo,
                'valor': self.campo,
                'tipo_error': 'Dependencia de campos',
                'detalle_error': mensaje
            }
            self.errores.append(error)
            return False, error  # Solo devuelve el error actual, no la lista completa de errores







