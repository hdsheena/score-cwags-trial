from docx import Document

file = Document('July 31 Results TESTFORPYTHON.docx')

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

def print_cells_in_row():
    print "Rows"
    print row.cells[0].text
    print row.cells[1].text
    print row.cells[2].text
    print row.cells[3].text
    print row.cells[4].text
    print row.cells[5].text
    print row.cells[6].text
    print row.cells[7].text
    #print row.cells[8].text
    print "End Rows"

for table_number in tables_with_data:
    current_table = tables[table_number]
    for row in current_table.rows[6:]:
        print_cells_in_row()

#file.save('results.docx')