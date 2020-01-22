
#Lead Component of Comparison


#import leads
#write leads matrix to new spreadsheet
#create new spreadsheet

worksheet2 = workbook.add_worksheet('Excel - Leads')
lead_matrix = import_object2(leadextract)

write_to_excel(lead_matrix,0,0,worksheet2,workbook)
        
#creating dashboard
worksheet3 = workbook.add_worksheet('Comparison')
worksheet3.write('V2','State')
worksheet3.write('V3','Source')
worksheet3.write('V4','Business')
worksheet3.write('W2','<>')
worksheet3.write('W3','<>')
worksheet3.write('W4','Advisory')

worksheet3.write('V11','Compare Date 1')
worksheet3.write('V12','Compare Date 2')
worksheet3.write('V13','Report Date')
worksheet3.write('W2','1/10/2017')
worksheet3.write('W3','31/03/2018')
worksheet3.write('W4','16/04/2018')

worksheet3.write('Y2','=CONCATENATE("<=",W11)')
worksheet3.write('Y3','=CONCATENATE(">=",W11)')
worksheet3.write('Y6','=CONCATENATE("<",W11)')
worksheet3.write('Y7','=CONCATENATE(">",W11)')
worksheet3.write('Z7','=CONCATENATE("<=",W12)')

worksheet3.write('AB2','General Business')
worksheet3.write('AB3','<>Partner')

worksheet3.write('B10','Created Prior')
worksheet3.write('C10',"=SUMIFS('Excel - Leads'!AF:AF,Lead_Record_Type,AB2,Lead_Source,AB3,'Excel - Leads'!O:O,Comparision!Y2,'Excel - Leads'!M:M,Comparision!W2,'Excel - Leads'!E:E,Comparision!W3)")
worksheet3.write('B11','Converted Prior')
worksheet3.write('C11',"=SUMIFS('Excel - Leads'!AF:AF,Lead_Record_Type,AB2,Lead_Source,AB3,'Excel - Leads'!O:O,Comparision!Y2,'Excel - Leads'!Q:Q,Y2,'Excel - Leads'!M:M,Comparision!W2,'Excel - Leads'!E:E,Comparision!W3)")
worksheet3.write('B12','Created Prior')
worksheet3.write('C12',"=SUMIFS('Excel - Leads'!AF:AF,Lead_Record_Type,AB2,Lead_Source,AB3,'Excel - Leads'!O:O,Comparision!Y2,'Excel - Leads'!R:R,Comparision!Y2,Lead_Status,\"Closed\",'Excel - Leads'!M:M,Comparision!W2,'Excel - Leads'!E:E,Comparision!W3)")
worksheet3.write('W11','1/10/2017')
worksheet3.write('B13','=CONCATENATE("As of ",TEXT(W11,"DD-MMM-YY"))')




#copy both snapshots into new spreadsheet on individual tabs 
