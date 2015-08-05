from docx import Document
file = Document('July 31 Results.docx')

#get the tables in the file
tables = file.tables
#how many tables?
number_of_tables = len(tables)
tables_with_data = []
n=0
while n < number_of_tables:
    n += 1
    tables_with_data.append(n)
    n += 1
print tables_with_data
for table_number in tables_with_data:
    current_table = tables[table_number]
    for row in current_table.rows:
        for cell in row.cells:
            
            print cell.text
