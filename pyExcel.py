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
    },
    "Atitudes e Valores":{
        1:{
            "R/RESP": 0.08,
            "I/EM": 0.05,
            "CTAR": 0.08,
            "AUT": 0.02,
            "PONT": 0.02
        }
    }
}

group_titles_format = workbook.add_format()
group_titles_format.set_right(6)
group_titles_format.set_left(6)
group_titles_format.set_top(2)
group_titles_format.set_bottom(1)
group_titles_format.set_align('center')
group_titles_format.set_align('vcenter')
group_titles_format.set_bold()

subgroupSumFormat = workbook.add_format()
subgroupSumFormat.set_bold()
subgroupSumFormat.set_border(1)
subgroupSumFormat.set_right(6)

subGroupMaxFormat = workbook.add_format()
subGroupMaxFormat.set_bold()
subGroupMaxFormat.set_border(1)
subGroupMaxFormat.set_right(6)
subGroupMaxFormat.set_num_format('0%')


startGroup = {
    "row": 10,
    "col": 2 
}

startSubGroup = {
    "row": 10,
    "col" : 2
}

currentGroup = {
    "row": 10,
    "col": 2
}

groupTot = 0
numSubGroupFields = 0
formulaString = ""
weightCell = ""

subGroupsValueCols = []

for i in groupsDict:
    for j in groupsDict[i]:
        for k in groupsDict[i][j]:
            worksheet.write(currentGroup["row"], currentGroup["col"], groupsDict[i][j][k],percent_fmt)
            worksheet.write(currentGroup["row"] - 1, currentGroup["col"], k)
            currentGroup['col'] += 1
            numSubGroupFields += 1
            # groupTot += groupsDict[i][j][k]
        worksheet.write_formula(currentGroup["row"], currentGroup["col"],'=SUM({}:{})'.format(xl_rowcol_to_cell(startSubGroup["row"], startSubGroup["col"]), xl_rowcol_to_cell(currentGroup["row"],currentGroup['col'] - 1)),subGroupMaxFormat)
        subGroupsValueCols.append(currentGroup["col"])
        for h in range(1,numberStudents+1):
            for s in range (1, numSubGroupFields + 1):
                weightCell = xl_rowcol_to_cell(currentGroup['row'], currentGroup['col'] - s)
                formulaString += '+({}*{})'.format(weightCell, xl_rowcol_to_cell(currentGroup['row'] + h, currentGroup['col'] - s)) ##move rows to get aliiggn, currentGroup row moves without saving
            worksheet.write_formula(currentGroup['row'] + h, currentGroup['col'], '=' + formulaString, subgroupSumFormat)
            formulaString = ""
        
        numSubGroupFields = 0
        currentGroup['col'] += 1;
        startSubGroup['col'] = currentGroup["col"]
        # groupTot = 0
    worksheet.merge_range(startGroup['row'] - 2, startGroup['col'], currentGroup["row"] - 2, currentGroup["col"] - 1, i, group_titles_format)
    startGroup['col'] = currentGroup['col']
    startGroup['row'] = currentGroup['row']


for i in range(0, numberStudents + 1):
    formulaString = ""
    for j in subGroupsValueCols:
        formulaString += "+{}".format(xl_rowcol_to_cell(currentGroup["row"] + i ,j))
    worksheet.write_formula(currentGroup['row'] + i, currentGroup["col"], formulaString, percent_fmt if (i == 0) else None)
    formulaString = ""



currentGroup['col'] += 1



        















# cell_format = workbook.add_format()
# cell_format.set_shrink()

# worksheet.write('A1', 'Hello world',cell_format)

workbook.close()

