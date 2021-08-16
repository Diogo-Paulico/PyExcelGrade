import xlsxwriter, json
from xlsxwriter.utility import xl_cell_to_rowcol, xl_cell_to_rowcol_abs, xl_rowcol_to_cell

f = open('data.json')

data = json.load(f)


school = data['school']
school_year = data['school_year']
className = data['className']
numberStudents = len(data['students'])

def getTermString(term):
    return 'Inglês - Avaliação {}º Período'.format(term)

workbook = xlsxwriter.Workbook(data['filename'])

grades = data['grades']



gradeSheet = workbook.add_worksheet(data['gradeSheetName'])

gradeSheet.write_string(1,1, "Mínimo")
gradeSheet.write_string(1,2, "Nota")

valuesRange = xl_rowcol_to_cell(2, 1) + ":" + xl_rowcol_to_cell(1 + len(grades), 1)
gradesRange = xl_rowcol_to_cell(2, 2) + ":" + xl_rowcol_to_cell(1 + len(grades), 2)

currentRow = 2

for y in grades:
    gradeSheet.write_number(currentRow,1,grades[y])
    gradeSheet.write_string(currentRow,2,y)
    currentRow+=1

gradeSheet.protect()



school_format = workbook.add_format({'font_size': 14, 'bold':True})
school_format.set_align('center')
school_format.set_align('vcenter')


titles_format = workbook.add_format({'font_size': 12, 'bold':True})
titles_format.set_align('center')
titles_format.set_align('vcenter')

classFormat = workbook.add_format({'font_size': 12, 'bold':True})
classFormat.set_align('vcenter')

prevEvalMergeFormat = workbook.add_format()
prevEvalMergeFormat.set_bold()
prevEvalMergeFormat.set_border(1)
prevEvalMergeFormat.set_align('center')
prevEvalMergeFormat.set_align('vcenter')
prevEvalMergeFormat.set_font_size(10)



merge_format = workbook.add_format({'align': 'center'})
merge_format.set_bold()
merge_format.set_border(2)
merge_format.set_align('center')
merge_format.set_align('vcenter')
merge_format.set_font_size(10)


num_format = workbook.add_format()
num_format.set_border(2)
num_format.set_align('center')
num_format.set_bold()
num_format.set_font_size(10)


student_format = workbook.add_format()
student_format.set_border(2)
student_format.set_right(6)
student_format.set_bold()
student_format.set_font_size(10)


nameFormat = workbook.add_format({'align': 'center'})
nameFormat.set_bold()
nameFormat.set_border(2)
nameFormat.set_align('center')
nameFormat.set_align('vcenter')
nameFormat.set_font_size(10)
nameFormat.set_right(6)


greyFormat = workbook.add_format()
greyFormat.set_border(1)
greyFormat.set_bg_color('#D8D8D8')


percent_fmt = workbook.add_format({'num_format': '0%'})


percent_center_fmt = workbook.add_format({'num_format': '0%'})
percent_center_fmt.set_align('center')
percent_center_fmt.set_border(1)
percent_center_fmt.set_font_size(10)


field_name = workbook.add_format()
field_name.set_align('center')
field_name.set_right(1)
field_name.set_bottom(1)
field_name.set_top(1)
field_name.set_bold()
field_name.set_font_size(9)


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
subgroupSumFormat.set_num_format('0.00')
subgroupSumFormat.set_font_size(10)
subgroupSumFormat.set_align('center')

subGroupMaxFormat = workbook.add_format()
subGroupMaxFormat.set_bold()
subGroupMaxFormat.set_border(1)
subGroupMaxFormat.set_right(6)
subGroupMaxFormat.set_num_format('0%')
subGroupMaxFormat.set_align('center')
subGroupMaxFormat.set_font_size(10)


fieldValueFormat = workbook.add_format()
fieldValueFormat.set_top(1)
fieldValueFormat.set_right(1)
fieldValueFormat.set_bottom(1)
fieldValueFormat.set_align('center')
fieldValueFormat.set_font_size(10)

def getFormating(i, num):
    format = workbook.add_format()
    if i == 0:
        format.set_num_format('0%')
    else:
        format.set_num_format('0.00')
    format.set_top(1)
    format.set_bottom(1)
    format.set_bold()
    format.set_font_size(10)
    format.set_align('center')
    if num == 1:
        format.set_right(1)
    else:
        format.set_right(2)
    if num == -1:
        format.set_right(1)
        format.set_left(1)
    return format

gradeFormat = workbook.add_format()
gradeFormat.set_top(1)
gradeFormat.set_right(2)
gradeFormat.set_bottom(1)
gradeFormat.set_bold()
gradeFormat.set_align('center')
gradeFormat.set_font_size(10)

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



def buildSheet(groups, num, prevEval, sheetName):
    worksheet = workbook.add_worksheet(sheetName)
    worksheet.set_column_pixels(0, 0, 25)
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
    worksheet.merge_range('A9:A10','Nº', merge_format)

    for i in groups:
        for j in groups[i]:
            for k in groups[i][j]:
                worksheet.write(currentGroup["row"], currentGroup["col"], groups[i][j][k],percent_center_fmt)
                worksheet.write(currentGroup["row"] - 1, currentGroup["col"], k, field_name)
                currentGroup['col'] += 1
                numSubGroupFields += 1
                # groupTot += groups[i][j][k]
            worksheet.write_blank(currentGroup["row"] - 1, currentGroup["col"],'none', subGroupMaxFormat)
            worksheet.write_formula(currentGroup["row"], currentGroup["col"],'=SUM({}:{})'.format(xl_rowcol_to_cell(startSubGroup["row"], startSubGroup["col"]), xl_rowcol_to_cell(currentGroup["row"],currentGroup['col'] - 1)),subGroupMaxFormat)
            subGroupsValueCols.append(currentGroup["col"])
            for h in range(1,numberStudents+1):
                for s in range (1, numSubGroupFields + 1):
                    weightCell = xl_rowcol_to_cell(currentGroup['row'], currentGroup['col'] - s)
                    worksheet.write_blank(currentGroup['row'] + h, currentGroup['col'] - s, 'none', fieldValueFormat)
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


    startGroup['col'] = currentGroup["col"]
    startGroup['row'] = currentGroup['row'] - 1

    prevEval[sheetName] = {
        "row": currentGroup['row'], # not the first student grade, but where the 100% is in final
        "col": currentGroup['col']
    }

    for i in range(0, numberStudents + 1):
        formulaString = ""
        for j in subGroupsValueCols:
            formulaString += "+{}".format(xl_rowcol_to_cell(currentGroup["row"] + i ,j))
        worksheet.write_formula(currentGroup['row'] + i, currentGroup["col"], formulaString, getFormating(i,num))
        formulaString = ""

    if num == 1:
        currentGroup['col'] += 1

        worksheet.write_string(currentGroup['row'], currentGroup['col'], 'N', gradeFormat)

        for k in range(1,numberStudents + 1):
            worksheet.write_formula(currentGroup['row'] + k, currentGroup['col'], '=IF({}=0,"",LOOKUP({},{}!{},{}!{}))'.format(xl_rowcol_to_cell(currentGroup['row'] + k, currentGroup['col'] - 1),xl_rowcol_to_cell(currentGroup['row'] + k, currentGroup['col'] - 1), data['gradeSheetName'],valuesRange, data['gradeSheetName'],gradesRange), gradeFormat)

        merge_format.set_bottom(1)
        worksheet.merge_range(startGroup['row'] - 1, startGroup['col'], currentGroup['row'] - 1, currentGroup['col'], 'Final', merge_format)
    else:
        merge_format.set_bottom(1)
        worksheet.merge_range(startGroup['row'] - 1, startGroup['col'], currentGroup['row'] - 1, currentGroup['col'], 'Final', merge_format)
        currentGroup['col'] += 2
        
        for l in prevEval:
            worksheet.merge_range(currentGroup['row'] - 2, currentGroup['col'], currentGroup['row'] - 1, currentGroup['col'], l, prevEvalMergeFormat)
            worksheet.write_blank(currentGroup['row'], currentGroup['col'],'', greyFormat)
            for i in range(1, numberStudents + 1): #getFormating(1,1)
                if l == sheetName:
                    worksheet.write_formula(currentGroup['row'] + i, currentGroup['col'], "={}".format(xl_rowcol_to_cell(prevEval[l]['row'] + i, prevEval[l]['col'])), getFormating(1,-1))
                else:
                    worksheet.write_formula(currentGroup['row'] + i, currentGroup['col'], "='{}'!{}".format(l, xl_rowcol_to_cell(prevEval[l]['row'] + i, prevEval[l]['col'])), getFormating(1,-1))
            currentGroup['col'] += 1
            
        worksheet.write_blank(currentGroup['row'], currentGroup['col'],'', greyFormat)

        startGroup['col'] = currentGroup["col"]
        startGroup['row'] = currentGroup['row'] - 2

        for i in range(1, numberStudents + 1):
            formulaString = '=('
            for p in range(1, len(prevEval) + 1):
                formulaString += '+' + xl_rowcol_to_cell(currentGroup['row'] + i, currentGroup['col'] - p)
            formulaString += ')/' + str(len(prevEval))
            worksheet.write_formula(currentGroup['row'] + i, currentGroup['col'], formulaString, getFormating(1,-1))

        currentGroup['col'] += 1
        worksheet.write_string(currentGroup['row'], currentGroup['col'], 'N', getFormating(1,-1))

        for k in range(1, numberStudents + 1):
            worksheet.write_formula(currentGroup['row'] + k, currentGroup['col'], '=IF({}=0,"",LOOKUP({},Notas!{},Notas!{}))'.format(xl_rowcol_to_cell(currentGroup['row'] + k, currentGroup['col'] - 1),xl_rowcol_to_cell(currentGroup['row'] + k, currentGroup['col'] - 1), valuesRange, gradesRange), getFormating(1,-1))

        worksheet.merge_range(startGroup['row'] , startGroup['col'], currentGroup['row'] - 1, currentGroup['col'], 'Final', prevEvalMergeFormat)

    worksheet.merge_range(0,0,1,currentGroup['col'],school, school_format)
    worksheet.merge_range(2,0,3,currentGroup['col'], school_year, titles_format)
    worksheet.merge_range(4,0,5,currentGroup['col'], getTermString(num), titles_format)
    worksheet.merge_range(6,0,7,currentGroup['col'], className, classFormat)

    worksheet.write_blank(10,0,'',num_format)
    worksheet.write_blank(10,1,'',student_format)

    for i in range(1,numberStudents+1):
        worksheet.write_number((10+i),0,i,num_format)
        worksheet.write_string((10+i),1, data['students'][i-1],student_format)
        

    worksheet.merge_range('B9:B10', 'Nome', nameFormat)
    return prevEval


i = 1
prevEval = {}
for sheet in data['sheets']:
    # currSheet = data['sheets'][sheet]
    prevEval = buildSheet(sheet['groups'], i, prevEval, sheet['sheetName'])
    i += 1



# prevEval = buildSheet('',1,{}, "Avaliação 1.º Período")
# prevEval = buildSheet('',2,prevEval,'Avaliação 2.º Período')
# prevEval = buildSheet('',3,prevEval,'Avaliação 3.º Período')

workbook.close()