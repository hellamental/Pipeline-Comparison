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



def matrix_creator():
    w,h = len(22), len()