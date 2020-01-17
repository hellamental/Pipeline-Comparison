import csv


#import_milestone(csv_file)

def import_opportunities(csv_filename):
	f = open(csv_filename)
	csv_dictf = csv.DictReader(f)

	opportunity_matrix = []

	for row in csv_dictf:
		opportunity_matrix.append(row)

	#del opportunity_matrix[0]

	return opportunity_matrix	



def matrix_creator(matrixheight):
    w,h = 21, matrixheight;
    excel_matrix = [[0 for x in range(w)] for y in range(h)]

    xcount = 0
    ycount = 0
    Headings = ['ID','Account Name','Opportunity / Project Name','Close Date','Stage','Project Value($)','Probability (%)','Weighted Value ($)','Status'," ",'Account Name 2','Opportunity / Project Name 2','Close Date 2','Stage 2','Project Value($) 2','Probability (%) 2','Weighted Value ($) 2','Stage Changed To','Project Value Change','Weighted Value Change','Close Date Change']
    for i in Headings:
        excel_matrix[ycount][xcount]=i
        xcount += 1

    return excel_matrix


def write_to_excel(excel_matrix,excel_offset_col,excel_offset_row,worksheet,workbook):
    plain_format = workbook.add_format()
    plain_format.set_font_size(9)

    numrows = len(excel_matrix)
    print numrows
    numcols = len(excel_matrix[0])
    print numcols

    xref = 0
    yref = 0

    while yref < numrows:
        xref = 0
        while xref < numcols:
            value = excel_matrix[yref][xref]
            worksheet.write(yref+excel_offset_row,xref+excel_offset_col,value,plain_format)
            #print xref
            xref += 1
        #print yref    
        yref += 1
        