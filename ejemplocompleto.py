import openpyxl
from openpyxl.styles import Font, Alignment

# Crear un nuevo archivo de Excel
workbook = openpyxl.Workbook() #woorkbook es el libro de trabajo, representan todo el archivo de excel

#Seleccionar la hoja activa
hoja = workbook.active
hoja.title = "DatosPersonas"

#Agregar encabezados
hoja['A1'] = "NOMBRE"
hoja['B1'] = "EDAD"

#Agregar nombres y edades 
datos = [
    ["miguel", 23],
    ["daniel", 30],
    ["emely", 28],
    ["pedro", 35],
    ["marta", 25]
]

# Agregar los datos a las filas
for i, persona in enumerate(datos, start=2):
    hoja[f'A{i}'] = persona[0]  # Nombre 
    hoja[f'B{i}'] = persona[1]  # Edad

#Insertar la fórmula para calcular el promedio de las edades
hoja['A7'] = "PROMEDIO"
hoja['B7'] = "=AVERAGE(B2:B6)"  # Fórmula para calcular el promedio

#Modificar el estilo: tamaño de letra, negrita, y centrar
for i in range(2, 8):
    hoja[f'A{i}'].font = Font(name="Arial", size=9, bold=True)
    hoja[f'A{i}'].alignment = Alignment(horizontal="center", vertical="center")
    hoja[f'B{i}'].alignment = Alignment(horizontal="center", vertical="center")

# 7. Guardar el archivo Excel
workbook.save("ejemplo.completo.xlsx")
print("Archivo Excel creado y guardado exitosamente.")

#Leer los datos y el promedio desde el archivo para mostrarlos en la terminal

# Cargar el archivo para leerlo
workbook = openpyxl.load_workbook("ejemplo.completo.xlsx")
hoja = workbook["DatosPersonas"]

# Imprimir los nombres, edades y promedio
print("\nDatos en el archivo Excel:")
for i in range(2, 7):
    nombre = hoja[f'A{i}'].value
    edad = hoja[f'B{i}'].value
    print(f"Nombre: {nombre}, Edad: {edad}")

