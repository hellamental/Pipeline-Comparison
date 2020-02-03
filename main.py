from sys import argv
from datetime import datetime, date
from dateutil.parser import *
import xlsxwriter
import os
import pandas as pd
import sys
from functions import *
reload(sys)
sys.setdefaultencoding('utf-8')

script, oppextract1, oppextract2, accextract, leadextract = argv

path = r"D:\Verdia Pty Ltd\Systems Tools and IT - Salesforce\Data Extracts\Opportunities"
os.chdir(path)
print os.getcwd()


#import both pipeline snapshots to matrix
pipeline_snapshot1 = import_object(oppextract1)
pipeline_snapshot2 = import_object(oppextract2)
account_matrix = import_object(accextract)


print len(pipeline_snapshot1)
print len(pipeline_snapshot2)


#create a list of project ID's in pipeline_snapshot1
opp_matrix1 = []
for row in pipeline_snapshot1:
    if(row['RECORDTYPEID']=='0120K000000yfFWQAY' and (row['SALES_TEAM__C']=='Advisory' or row['SALES_TEAM__C']=='Large Programs') and (row['STAGENAME']=='Qualify' or row['STAGENAME']=='Validate/ Propose' or row['STAGENAME']=='Decide' or row['STAGENAME']=='Paid Feasibility')):#or row['STAGENAME']=='Closed Lost' or row['STAGENAME']=='Won' or row['STAGENAME']=='Hold')):
        opp_matrix1.append(row['ID'])
    else:
        pass
print str(len(opp_matrix1)) + " Snapshot1"

#create a list of project ID's in pipeline_snapshot2
opp_matrix2 = []
for row in pipeline_snapshot2:
    if(row['RECORDTYPEID']=='0120K000000yfFWQAY' and (row['SALES_TEAM__C']=='Advisory' or row['SALES_TEAM__C']=='Large Programs') and (row['STAGENAME']=='Qualify' or row['STAGENAME']=='Validate/ Propose' or row['STAGENAME']=='Decide' or row['STAGENAME']=='Paid Feasibility')):# or row['STAGENAME']=='Closed Lost' or row['STAGENAME']=='Won' or row['STAGENAME']=='Hold' )):
        opp_matrix2.append(row['ID'])
    else:
        pass
print str(len(opp_matrix2)) + " Snapshot2"

#identify any Project ID's that were created since the previous snapshot - created date > first nominated date
#identifying new vs existing opportunities
existing_opps = []
new_opps = []
for i in opp_matrix2:
    if i in opp_matrix1:
        existing_opps.append(i)
    else:
        new_opps.append(i)

print str(len(existing_opps)) + " existing opps"
print str(len(new_opps)) + " new opps"

for i in new_opps:
    for j in pipeline_snapshot2:
        if i == j['ID']:
            print j['ID'] + " - " + j['NAME'] + " - " + j['CREATEDDATE']
        else:
            pass

print "....."

#identify any Projects ID's in older snapshot not appearing in more recent snapshot (indicating deleted opportunities)
#identifying deleted opportunities
deleted_opps = []
for i in opp_matrix1:
    if i not in opp_matrix2:
        deleted_opps.append(i)
    else:
        pass

for i in deleted_opps:
    for j in pipeline_snapshot1:
        if i == j['ID']:
            print j['ID'] + " - " + j['NAME'] + " - " + j['CREATEDDATE']
        else:
            pass

print str(len(deleted_opps)) + " deleted opps"


#create matrix to house comparison to push to excel document
matrixheight = len(deleted_opps) + len(new_opps)+ len(existing_opps) + 1
print matrixheight

excel_matrix = matrix_creator(matrixheight)

xcount = 0
ycount = 1
for i in new_opps:
    excel_matrix[ycount][xcount]=i
    excel_matrix[ycount][8]='New Opportunity'
    ycount += 1

for i in existing_opps:
    excel_matrix[ycount][xcount]=i
    excel_matrix[ycount][8]='Existing Opportunity'
    ycount += 1

for i in deleted_opps:
    excel_matrix[ycount][xcount]=i
    excel_matrix[ycount][8]='Removed from Pipeline'
    ycount += 1
    
#declaring variables for excel matrix to make things easier

# 'ID' = 0
# 'Account Name' = 1
# 'Opportunity / Project Name' = 2
# 'Close Date' = 3
# 'Stage' = 4 
# 'Project Value($)' = 5
# 'Probability (%)' = 6
# 'Weighted Value ($)' = 7
# 'Status' = 8
# ' ' = 9
# 'Account Name 2' = 10
# 'Opportunity / Project Name 2' = 11
# 'Close Date 2' = 12
# 'Stage 2' = 13
# 'Project Value($) 2' = 14
# 'Probability (%) 2' = 15
# 'Weighted Value ($) 2' = 16
# 'Stage Changed To' = 17
# 'Project Value Change' = 18
# 'Weighted Value Change' = 19
# 'Close Date Change' = 20

# populating excel matrix details for new opps
for i in excel_matrix:

    for k in pipeline_snapshot1:
        if i[0] == k['ID']:
            i[1] = k['ACCOUNTID']
            i[2] = k['NAME']
            i[3] = k['CLOSEDATE']
            i[4] = k['STAGENAME']
            i[5] = float(k['VD_PROJECT_VALUE__C'])
            i[6] = float(k['PROBABILITY'])
            i[7] = float(k['WEIGHTED_PROJECT_VALUE__C'])


    for j in pipeline_snapshot2:
        if i[0] == j['ID']:
            i[10] = j['ACCOUNTID']
            i[11] = j['NAME']
            i[12] = j['CLOSEDATE']
            i[13] = j['STAGENAME']
            i[14] = float(j['VD_PROJECT_VALUE__C'])
            i[15] = float(j['PROBABILITY'])
            i[16] = float(j['WEIGHTED_PROJECT_VALUE__C'])
        else:
            pass

for i in excel_matrix:
    for j in account_matrix:
        if i[1] == j['ID']:
            i[1] = j['NAME']
        else:
            pass
        
        if i[10] == j['ID']:
            i[10] = j['NAME']
        else:
            pass

    if i[8] == 'Status': #i[8] == 'Removed from Pipeline' or
        pass
    elif i[8] == 'Removed from Pipeline':
        i[9] = ""
        i[14] = ""
        i[15] = ""
        i[16] = ""
        i[17] = ""
        i[18] = -i[5]
        i[19] = ""
        i[20] = ""

    elif i[8] == 'New Opportunity':
        if i[3] == 0:
            i[1] = ""
            i[2] = ""
            i[3] = ""
            i[4] = ""
            i[5] = ""
            i[6] = ""
            i[7] = ""
            i[9] = ""
            i[18] = i[14]
            i[17] = ""
            i[19] = ""
            i[20] = "" 
        else:
            i[5] = ""
            i[6] = ""
            i[7] = ""
            i[9] = ""
            i[17] = ""
            i[18] = i[14]
            i[19] = ""
            i[20] = ""
    #existing opportunity
    else:
        i[9] = ""
        #Stage Changed To - ok
        if i[4] != i[13]:
            i[17] = i[13]
        else:
            i[17] = ""

        #Project Value Change - ok
        if i[5] != i[14]:
            i[18] = i[14] - i[5]
        else:
            i[18] = ""
        
        #Weighted Value Change - ok
        if i[6] != i[15]:
            i[19] = i[15]-i[6]
        else:
            i[19] = ""
               
        #Close Date Change - ok
        if i[3] != i[12]:
            date1 = parse(i[3])
            date2 = parse(i[12])
            i[20] = date2 - date1
        else:
            i[20] = ""

#print len(excel_matrix)
#print excel_matrix

formula1 = "=SUBTOTAL(9,F3:F" + str(len(excel_matrix)+1) + ")"
formula2 = "=SUBTOTAL(9,O3:O" + str(len(excel_matrix)+1) + ")"
formula3 = "=SUBTOTAL(9,S3:S" + str(len(excel_matrix)+1) + ")"
#identify projects that have exited the pipeline because they were moved to Closed lost or Closed Won
#identify any Projects thats have been modified between snapshots


#identify what changes have occurred since last snapshot e.g. stage change, close date change, dollar value change etc. 

#create a new report spreadsheet - titles pipeline comparison report (Date 1 v Date 2)

path = "C:\Users\Mitchell.Dawson\Desktop"
os.chdir(path)
workbook_filename = "Pipeline_Comparison.xlsx"
workbook = xlsxwriter.Workbook(workbook_filename)
worksheet1 = workbook.add_worksheet('Pipeline Comparison')

#excel formats
plain_format = workbook.add_format()
plain_format.set_font_size(9)
money_format = workbook.add_format({'num_format': '$#,##0'})
money_format.set_font_size(9)


write_to_excel(excel_matrix,0,1,worksheet1,workbook)
worksheet1.write('F1',formula1,money_format)
worksheet1.write('O1',formula2,money_format)
worksheet1.write('S1',formula3,money_format)

worksheet1.write('X2','Movement',plain_format)
worksheet1.write('Y2','Count',plain_format)
worksheet1.write('Z2','Value',plain_format)
worksheet1.write('X3','New Opportunity',plain_format)
worksheet1.write('Y3','=COUNTIF(I:I,"=New Opportunity")',plain_format)
worksheet1.write('Z3','=SUMIF(I:I,"=New Opportunity",S:S)',money_format)
worksheet1.write('X4','Existing Opportunity',plain_format)
worksheet1.write('Y4','=COUNTIF(I:I,"=Existing Opportunity")',plain_format)
worksheet1.write('Z4','=SUMIF(I:I,"=Existing Opportunity",S:S)',money_format)
worksheet1.write('X5','Removed From Pipeline',plain_format)
worksheet1.write('Y5','=COUNTIF(I:I,"=Removed From Pipeline")',plain_format)
worksheet1.write('Z5','=SUMIF(I:I,"=Removed From Pipeline",S:S)',money_format)

worksheet1.write('X6','Hold',plain_format)
worksheet1.write('Y6','=COUNTIF(N:N,"=Hold")',plain_format)
worksheet1.write('Z6','=SUMIF(N:N,"=Hold",S:S)',money_format)
worksheet1.write('X7','Closed Lost',plain_format)
worksheet1.write('Y7','=COUNTIF(N:N,"=Closed Lost")',plain_format)
worksheet1.write('Z7','=SUMIF(N:N,"=Closed Lost",S:S)',money_format)
worksheet1.write('X8','Won',plain_format)
worksheet1.write('Y8','=COUNTIF(N:N,"=Won")',plain_format)
worksheet1.write('Z8','=SUMIF(N:N,"=Won",S:S)',money_format)
worksheet1.write('X9','Other',plain_format)
worksheet1.write('Y9','=COUNTIFS(I:I,"=Removed from Pipeline",N:N,"<>Won",N:N,"<>Closed Lost",N:N,"<>Hold")',plain_format)
worksheet1.write('Z9','=SUMIFS(S:S,I:I,"=Removed from Pipeline",N:N,"<>Won",N:N,"<>Closed Lost",N:N,"<>Hold")',money_format)

worksheet1.write('X11','Pipeline Movement Summary',plain_format)
worksheet1.write('X12','Stage',plain_format)
worksheet1.write('X13','Previous Pipeline (Count)',plain_format)
worksheet1.write('X14','Previous Pipeline (Value)',plain_format)

worksheet1.write('X16','Entering Stage (Count)',plain_format)
worksheet1.write('X17','Entering Stage (Value)',plain_format)
worksheet1.write('X18','Exiting Stage (Count)',plain_format)
worksheet1.write('X19','Exiting Stage (Value)',plain_format)
worksheet1.write('X20','Value Changes (No Stage Movement)',plain_format)

worksheet1.write('X22','Current Pipeline (Count)',plain_format)
worksheet1.write('X23','Current Pipeline (Value)',plain_format)

######Qualify
worksheet1.write('Y12','Qualify',plain_format)
worksheet1.write('Y13','=COUNTIF($E:$E,Y$12)',plain_format)
worksheet1.write('Y14','=SUMIF($E:$E,Y$12,$F:$F)',money_format)

worksheet1.write('Y16','=COUNTIFS($N:$N,Y$12,$E:$E,"<>"&Y$12)',plain_format)
worksheet1.write('Y17','=SUMIFS($F:$F,$N:$N,Y$12,$E:$E,"<>"&Y$12)+SUMIFS($O:$O,$N:$N,Y$12,$E:$E,"<>"&Y$12,$I:$I,$X$3)+SUMIFS($S:$S,$N:$N,Y$12,$E:$E,"<>"&Y$12,$I:$I,$X$4)',money_format)
worksheet1.write('Y18','=-COUNTIFS($E:$E,Y$12,$N:$N,"<>" & Y$12)',plain_format)
worksheet1.write('Y19','=-SUMIFS($F:$F,$E:$E,Y$12,$N:$N,"<>"&Y$12)',money_format)
worksheet1.write('Y20','=SUMIFS($S:$S,$N:$N,Y$12,$E:$E,Y$12,$I:$I,$X$4)',money_format)

worksheet1.write('Y22','=COUNTIF($N:$N,Y$12)',plain_format)
worksheet1.write('Y23','=SUMIF($N:$N,Y$12,$O:$O)',money_format)

###### Validate/ Propose
worksheet1.write('Z12','Validate/ Propose',plain_format)
worksheet1.write('Z13','=COUNTIF($E:$E,Z$12)',plain_format)
worksheet1.write('Z14','=SUMIF($E:$E,Z$12,$F:$F)',money_format)

worksheet1.write('Z16','=COUNTIFS($N:$N,Z$12,$E:$E,"<>" & Z$12)',plain_format)
worksheet1.write('Z17','=SUMIFS($F:$F,$N:$N,Z$12,$E:$E,"<>"&Z$12)+SUMIFS($O:$O,$N:$N,Z$12,$E:$E,"<>"&Z$12,$I:$I,$X$3)+SUMIFS($S:$S,$N:$N,Z$12,$E:$E,"<>"&Z$12,$I:$I,$X$4)',money_format)
worksheet1.write('Z18','=-COUNTIFS($E:$E,Z$12,$N:$N,"<>" & Z$12)',plain_format)
worksheet1.write('Z19','=-SUMIFS($F:$F,$E:$E,Z$12,$N:$N,"<>"&Z$12)',money_format)
worksheet1.write('Z20','=SUMIFS($S:$S,$N:$N,Z$12,$E:$E,Z$12,$I:$I,$X$4)',money_format)

worksheet1.write('Z22','=COUNTIF($N:$N,Z$12)',plain_format)
worksheet1.write('Z23','=SUMIF($N:$N,Z$12,$O:$O)',money_format)

###### Paid Feasibility
worksheet1.write('AA12','Paid Feasibility',plain_format)
worksheet1.write('AA13','=COUNTIF($E:$E,AA$12)',plain_format)
worksheet1.write('AA14','=SUMIF($E:$E,AA$12,$F:$F)',money_format)

worksheet1.write('AA16','=COUNTIFS($N:$N,AA$12,$E:$E,"<>" & AA$12)',plain_format)
worksheet1.write('AA17','=SUMIFS($F:$F,$N:$N,AA$12,$E:$E,"<>"&AA$12)+SUMIFS($O:$O,$N:$N,AA$12,$E:$E,"<>"&AA$12,$I:$I,$X$3)+SUMIFS($S:$S,$N:$N,AA$12,$E:$E,"<>"&AA$12,$I:$I,$X$4)',money_format)
worksheet1.write('AA18','=-COUNTIFS($E:$E,AA$12,$N:$N,"<>" & AA$12)',plain_format)
worksheet1.write('AA19','=-SUMIFS($F:$F,$E:$E,AA$12,$N:$N,"<>"&AA$12)',money_format)
worksheet1.write('AA20','=SUMIFS($S:$S,$N:$N,AA$12,$E:$E,AA$12,$I:$I,$X$4)',money_format)

worksheet1.write('AA22','=COUNTIF($N:$N,AA$12)',plain_format)
worksheet1.write('AA23','=SUMIF($N:$N,AA$12,$O:$O)',money_format)

###### Decide
worksheet1.write('AB12','Decide',plain_format)
worksheet1.write('AB13','=COUNTIF($E:$E,AB$12)',plain_format)
worksheet1.write('AB14','=SUMIF($E:$E,AB$12,$F:$F)',money_format)

worksheet1.write('AB16','=COUNTIFS($N:$N,AB$12,$E:$E,"<>" & AB$12)',plain_format)
worksheet1.write('AB17','=SUMIFS($F:$F,$N:$N,AB$12,$E:$E,"<>"&AB$12)+SUMIFS($O:$O,$N:$N,AB$12,$E:$E,"<>"&AB$12,$I:$I,$X$3)+SUMIFS($S:$S,$N:$N,AB$12,$E:$E,"<>"&AB$12,$I:$I,$X$4)',money_format)
worksheet1.write('AB18','=-COUNTIFS($E:$E,AB$12,$N:$N,"<>" & AB$12)',plain_format)
worksheet1.write('AB19','=-SUMIFS($F:$F,$E:$E,AB$12,$N:$N,"<>"&AB$12)',money_format)
worksheet1.write('AB20','=SUMIFS($S:$S,$N:$N,AB$12,$E:$E,AB$12,$I:$I,$X$4)',money_format)

worksheet1.write('AB22','=COUNTIF($N:$N,AB$12)',plain_format)
worksheet1.write('AB23','=SUMIF($N:$N,AB$12,$O:$O)',money_format)

##### Total
worksheet1.write('AC12','Total',plain_format)
worksheet1.write('AC13','=SUM(Y13:AB13)',plain_format)
worksheet1.write('AC14','=SUM(Y14:AB14)',money_format)

worksheet1.write('AC16','=SUM(Y16:AB16)',plain_format)
worksheet1.write('AC17','=SUM(Y17:AB17)',money_format)
worksheet1.write('AC18','=SUM(Y18:AB18)',plain_format)
worksheet1.write('AC19','=SUM(Y19:AB19)',money_format)
worksheet1.write('AC20','=SUM(Y20:AB20)',money_format)

worksheet1.write('AC22','=SUM(Y22:AB22)',plain_format)
worksheet1.write('AC23','=SUM(Y23:AB23)',money_format)




workbook.close()






