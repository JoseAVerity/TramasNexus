import configparser


class Config:
    def __init__(self):
        """Inicializa la configuración leyendo el archivo config.ini"""
        self.config = configparser.ConfigParser()
        self.config.read(r'C:/Users/Usuario/PycharmProjects/TramasNexus/config.ini')

        # Verifica si el archivo fue leído correctamente
        if not self.config.sections():
            raise FileNotFoundError(f"El archivo de configuración '{r'C:/Users/Usuario/PycharmProjects/TramasNexus/config.ini'}' no se encontró o está vacío.")

    def obtener(self, section, key):
        """Obtiene el valor correspondiente a una clave dentro de una sección"""
        try:
            return self.config[section][key]
        except KeyError as e:
            raise KeyError(f"No se encontró la clave '{key}' en la sección '{section}'. Error: {e}")

    def obtener_path(self, key):
        """Obtiene un valor de la sección 'DEFAULT'"""
        try:
            return self.config['PATH'][key]
        except KeyError as e:
            raise KeyError(f"No se encontró la clave '{key}' en la sección 'DEFAULT'. Error: {e}")

    def obtener_seccion(self, section):
        """Obtiene todas las claves y valores dentro de una sección específica"""
        try:
            return dict(self.config.items(section))
        except KeyError as e:
            raise KeyError(f"La sección '{section}' no existe en el archivo de configuración. Error: {e}")






