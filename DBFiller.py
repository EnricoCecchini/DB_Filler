'''
Python program to generate SQL code from EXCEL file to fill tables
'''

import openpyxl

path = input('File Path: ')
print(path)
wb = openpyxl.load_workbook(path)
sheet = wb.active
print()

n = 0
atributos = []
tabla = input("Nombre de Tabla: ")
print()
at = []

# Receive column number and attributes names for table
while True:
    try:
        n = int(input('Columna: '))
        at.append(input('Nombre Atributo: '))
        atributos.append(n)
    except ValueError:
        break

datos = []
col = []
# Amount of ROWS in spreadsheet
filas = 46
# Iterate rows and store data in list of lists
for i in atributos:
    for row_cells in sheet.iter_rows(min_row=2, max_row=filas):
        try:
            col.append(int(row_cells[i].value))
        except ValueError:
            col.append(str(row_cells[i].value))
        except TypeError:
            col.append(str(row_cells[i].value))
    datos.append(col)
    col = []
#print(len(datos))

# Generate SQL template with table name and attributes
sql = f'INSERT INTO {tabla} ('
for i in range(len(at)):
    if i < len(at)-1:
        sql += at[i] + ', '
    else:
        sql += at[i]
sql += ')\nValues ('
#print(sql)

val = []
val2 = []
c = 0

# Read values from list and store them in a string in SQL code format
for i in range(filas-1):
    for j in range(len(datos)):
        try:
            if isinstance(datos[j][i], str):
                sql += f'\'{datos[j][i]}\'' 
            else:
                sql += f'{datos[j][i]}'

            if j < len(datos)-1:
                sql += ', '
        except IndexError:
            break
    sql += ')'
    if i < filas-1:
        sql += ', \n('

print(sql)