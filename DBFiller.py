'''
Python program to generate SQL code from EXCEL file to fill tables
'''
import openpyxl

# open Excel file
path = input('File Path: ')
print(path)
wb = openpyxl.load_workbook(path)
sheet = wb.active
print()

# Table class to store data
class Table:

    def __init__(self, name, attributes, columns, data, rows):
        self.name = name
        self.attributes = attributes
        self.columns = columns
        self.data = data
        self.rows = rows
    
    # Function to generate SQL code for the table
    def generateSQL(self):
        sql = f'INSERT INTO {self.name} ('
        for i in range(len(self.attributes)):
            if i < len(self.attributes)-1:
                sql += self.attributes[i] + ', '
            else:
                sql += self.attributes[i]
        sql += ')\nValues ('

        # Read values from list and store them in a string in SQL code format
        for i in range(self.rows):
            for j in range(len(self.data)):
                try:
                    if isinstance(self.data[j][i], str):
                        sql += f'\'{self.data[j][i]}\'' 
                    else:
                        sql += f'{self.data[j][i]}'

                    if j < len(self.data)-1:
                        sql += ', '
                except IndexError:
                    break

            if i < self.rows-2:
                sql += '), \n('
            elif i < self.rows-1:
                sql += ')'
        
        return sql


tables = []
rows = int(input('Amount of Rows in columns: '))
print()

while True:
    print('*'*5, 'New Table', '*'*5)

    name = input("Table Name: ")
    attributes = []
    columns = []
    data = []

    if name == '':
        break
    
    print()
    print('*'*5, 'Attributes', '*'*5)
    # Receive column number and attributes names for table
    while True:
        try:
            n = int(input('Column Number: '))
            attributes.append(input('Attribute Name: '))
            columns.append(n)
        except ValueError:
            break

    col = []
    # Read data from the column and stores it in a list
    for i in columns:
        for row_cells in sheet.iter_rows(min_row=2, max_row=rows):
            try:
                col.append(float(row_cells[i].value))
            except ValueError:
                col.append(str(row_cells[i].value))
            except TypeError:
                col.append(str(row_cells[i].value))
            
        data.append(col)
        col = []
    tables.append(Table(name, attributes, columns, data, rows))
    print()

sql = ''
for table in tables:
    sql += table.generateSQL()
    sql += '\n\n'

print()
print('*'*5, 'SQL Code', '*'*5)
print(sql)