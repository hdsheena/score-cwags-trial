from docx import Document
import json

file = Document('July 31 Results.docx')

#open json reference of numbers to dog names
referenceFile = open('cwagsnumberref.json', 'r')
parsedFile = {}
parsedrefFile = json.loads(referenceFile.read())
#print parsedrefFile
m = 1
for i in parsedrefFile:
    print parsedrefFile[i]
    if len(i) > 1:
        parsedrefFile[i][unicode('Team Number')] = unicode(m)
        m +=1
print parsedrefFile

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


def print_cells_in_row():
    print "Rows"
    print row.cells[0].text
    print row.cells[1].text
    print row.cells[2].text
    print row.cells[3].text
    print row.cells[4].text
    print row.cells[5].text
    print row.cells[6].text
    print "End Rows"

# Calling the columns makes this work. No idea why. Bug in the program? 
#it works!
for table_number in tables_with_data:
    current_table = tables[table_number]
    for row in current_table.rows[6:]:
       
        #this is the dog name in the table!
        dogname = row.cells[2].text
        score = row.cells[7].text
        handler = row.cells[3].text
        team_number = row.cells[6].text
        dogData = parsedrefFile[dogname]
        row.cells[1].text = dogData['Registration Number']
        #print_cells_in_row()
        #print dogData['Handler']
        print handler
        row.cells[3].text = dogData['Handler']
        #row.cells[3].text = 'Handler'
        #print_cells_in_row()
        print handler
        print score
        print row.cells[4].text
        if len(dogname) >1:
            row.cells[6].text = dogData['Team Number']
        #row.cells[6].text = 'Team number'
        #print_cells_in_row()
       # for cell in row.cells:
         #   print cell.text

#Give me the number of Qs in a round:
qs_per_round = {}
for table_number in tables_with_data:
    current_table = tables[table_number]
    passed_rounds = 0
    entered_rounds = 0
    for row in current_table.rows[6:]:
       
        #this is the dog name in the table!
        dogname = row.cells[2].text
        score = row.cells[7].text
        if len(score) >0:
            entered_rounds += 1
        if score == "P":
            passed_rounds += 1
    header_table = tables[table_number-1]
    header_table.row_cells(1)[3].text = "# Entered: " + str(entered_rounds)
    print header_table.row_cells(0)[0].text
    print passed_rounds
    print "Entered"
    print entered_rounds
    print "_____"
            

file.save('results.docx')