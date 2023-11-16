import openpyxl

# Especifica la ruta completa del archivo Excel
ruta_archivo = r'C:\Users\atorrez\Desktop\TRABAJOS\MADE\scrpits-configuraciones\mi_tabla.xlsx'

# Abrir un archivo Excel existente
archivo_excel = openpyxl.load_workbook(ruta_archivo)

# Seleccionar una hoja de cálculo específica
hoja = archivo_excel['TABLA1']

# Especifica la ruta completa para el archivo de texto
ruta_salida_txt = r'C:\Users\atorrez\Desktop\TRABAJOS\MADE\scrpits-configuraciones\scripts_tabla1.txt'

# Contador
count = 1

# Abre un archivo de texto en modo escritura en la nueva ubicación
with open(ruta_salida_txt, 'w') as archivo_txt:
    for fila in range(3, 98):
        
        dato1= hoja[f'B{fila}'].value #fila B
        dato2= hoja[f'C{fila}'].value #fila C
        dato3= hoja[f'D{fila}'].value #fila D
        dato4= hoja[f'E{fila}'].value #fila E
        dato5= hoja[f'F{fila}'].value #fila F
        dato6= hoja[f'G{fila}'].value #fila G

        # IMPRIMIR FILA EN CONSOLA
        # print(f'FILA {contador}')
        # print(f'    1: {dato1}')
        # print(f'    2: {dato2}')
        # print(f'    3: {dato3}')
        # print(f'    4: {dato4}')
        # print(f'    5: {dato5}')
        # print(f'    6: {dato6}')
        
        # IMPRIMIR FILA EN TXT
        # archivo_txt.write(f'FILA {count}\n')
        # archivo_txt.write(f'    1: {dato1}\n')
        # archivo_txt.write(f'    2: {dato2}\n')
        # archivo_txt.write(f'    3: {dato3}\n')
        # archivo_txt.write(f'    4: {dato4}\n')
        # archivo_txt.write(f'    5: {dato5}\n')
        # archivo_txt.write(f'    6: {dato6}\n')
        # archivo_txt.write('\n')

        # IMPRIMIR DISEÑO FINAL EN TXT
        archivo_txt.write(f'edit {dato1}\n')
        archivo_txt.write(f'set interface "ENTEL_A"\n')
        archivo_txt.write(f'set ike-version 1\n')
        archivo_txt.write(f'set keylife {dato2}\n')
        archivo_txt.write(f'set peertype any\n')
        archivo_txt.write(f'set net-device disable\n')
        archivo_txt.write(f'set proposal {dato3}\n')
        archivo_txt.write(f'set dhgrp {dato4}\n')
        archivo_txt.write(f'set remote-gw {dato5}\n')
        archivo_txt.write(f'set psksecret {dato6}\n')
        archivo_txt.write(f'next\n')
        archivo_txt.write('\n')

        count += 1


# Cerrar el archivo Excel
archivo_excel.close()
print(f"TXT guardado con exito! - '{ruta_salida_txt}'")