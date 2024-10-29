import os
import re

# Ruta de la carpeta que contiene los archivos
folder_path = 'C://Users//fbourse//Downloads//OrcaFiles'

# Expresión regular para eliminar caracteres no ASCII
non_ascii_pattern = re.compile(r'[^\x00-\x7F]+')

# Recorre todos los archivos en la carpeta
for filename in os.listdir(folder_path):
    # Crea la ruta completa del archivo
    file_path = os.path.join(folder_path, filename)

    # Verifica si es un archivo (no una carpeta)
    if os.path.isfile(file_path):
        # Elimina caracteres no ASCII del nombre del archivo
        new_filename = non_ascii_pattern.sub('', filename)

        # Crea la nueva ruta del archivo
        new_file_path = os.path.join(folder_path, new_filename)
        
        # Renombra el archivo
        os.rename(file_path, new_file_path)

print("Renombrado completado.")
