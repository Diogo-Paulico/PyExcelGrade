import xlsxwriter
from xlsxwriter.utility import xl_cell_to_rowcol, xl_rowcol_to_cell

school = 'Agrupamento de Escolas Santa Iria de Azoia'
school_year = 'Ano Letivo de 2020/2021'
numberStudents = 28

def getTermString(term):
    return 'Inglês - Avaliação {}º Período'.format(term)

workbook = xlsxwriter.Workbook('grade.xlsx')
worksheet = workbook.add_worksheet("Avaliação 1.º Período")

worksheet.write('A1', school)

worksheet.write('A3', school_year)

worksheet.write('A5', getTermString('1'))

# merge_format = workbook.add_format({'align': 'center'})
merge_format = workbook.add_format({'align': 'center'})
merge_format.set_bold()

# (sr,sc) = xl_cell_to_rowcol('A9')
# (er,ec) = xl_cell_to_rowcol('A10')

worksheet.merge_range('A9:A10','Nº', merge_format)
# worksheet.merge_range(sr,sc,er,ec,'Nº')
# worksheet.write('A', school_year)



for i in range(1,numberStudents+1):
    worksheet.write_number((10+i),0,i)
    
















# cell_format = workbook.add_format()
# cell_format.set_shrink()

# worksheet.write('A1', 'Hello world',cell_format)

workbook.close()

