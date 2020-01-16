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
    w,h = 22, matrixheight;
    excel_matrix = [[0 for x in range(w)] for y in range(h)]

    xcount = 0
    ycount = 0
    Headings = ['ID','Account Name','Opportunity / Project Name','Close Date','Stage','Project Value($)','Probability (%)','Weighted Value ($)','Status'," ",'Account Name','Opportunity / Project Name','Close Date','Stage','Project Value($)','Probability (%)','Weighted Value ($)','Status','Stage Changed To','Project Value Change','Weighted Value Change','Close Date Change']
    for i in Headings:
        x=i
        excel_matrix[ycount][xcount]=x
        xcount += 1

    return excel_matrix