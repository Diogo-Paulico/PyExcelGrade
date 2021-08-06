import xlsxwriter
from xlsxwriter.utility import xl_cell_to_rowcol, xl_rowcol_to_cell

school = 'Agrupamento de Escolas Santa Iria de Azoia'
school_year = 'Ano Letivo de 2020/2021'
numberStudents = 28

def getTermString(term):
    return 'Inglês - Avaliação {}º Período'.format(term)

workbook = xlsxwriter.Workbook('grade.xlsx')
worksheet = workbook.add_worksheet("Avaliação 1.º Período")

school_format = workbook.add_format({'font_size': 14, 'bold':True})

titles_format = workbook.add_format({'font_size': 12, 'bold':True})


worksheet.write('A1', school, school_format)

worksheet.write('A3', school_year,titles_format)

worksheet.write('A5', getTermString('1'),titles_format)

# merge_format = workbook.add_format({'align': 'center'})
merge_format = workbook.add_format({'align': 'center'})
merge_format.set_bold()
merge_format.set_border(2)
merge_format.set_align('center')
merge_format.set_align('vcenter')
merge_format.set_font_size(10)

# (sr,sc) = xl_cell_to_rowcol('A9')
# (er,ec) = xl_cell_to_rowcol('A10')

worksheet.merge_range('A9:A10','Nº', merge_format)
# worksheet.merge_range(sr,sc,er,ec,'Nº')
# worksheet.write('A', school_year)
num_format = workbook.add_format()
num_format.set_border(2)
num_format.set_align('center')
num_format.set_bold()
num_format.set_font_size(10)


worksheet.set_column_pixels(0, 0, 25)

student_format = workbook.add_format()
student_format.set_border(2)
student_format.set_right(6)
student_format.set_bold()
student_format.set_font_size(10)


worksheet.merge_range('B9:B10', 'Nome', merge_format)


percent_fmt = workbook.add_format({'num_format': '0%'})
# worksheet.merge_range('C9:C10', 0.10, percent_fmt);


for i in range(1,numberStudents+1):
    worksheet.write_number((10+i),0,i,num_format)
    worksheet.write_string((10+i),1,'Aluno/a',student_format)
    
groupsDict = {
    "Capacidades e Conhecimentos": {
        1: {
            "DIAG" : 0,
            "Teste 1": 0.25,
            "Teste 2": 0.25
        },
        2: {
            "Dim. Pratica": 0.1,
            "Leit": 0.05,
            "Oral": 0.05,
            "Escr": 0.05
        }
    }
}


startGroup = {
    "row": 10,
    "col": 2 
}

currentGroup = {
    "row": 10,
    "col": 2
}

groupTot = 0



for i in groupsDict:
    for j in groupsDict[i]:
        for k in groupsDict[i][j]:
            worksheet.write(currentGroup["row"], currentGroup["col"], groupsDict[i][j][k],percent_fmt)
            worksheet.write(currentGroup["row"] - 1, currentGroup["col"], k)
            currentGroup['col'] += 1;
            groupTot += groupsDict[i][j][k]
        worksheet.write(currentGroup["row"], currentGroup["col"], groupTot,percent_fmt)
        currentGroup['col'] += 1;
        groupTot = 0
    worksheet.merge_range(startGroup['row'] - 2, startGroup['col'], currentGroup["row"] - 2, currentGroup["col"] - 1, i)


        















# cell_format = workbook.add_format()
# cell_format.set_shrink()

# worksheet.write('A1', 'Hello world',cell_format)

workbook.close()

