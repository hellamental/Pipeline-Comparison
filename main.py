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

script, oppextract1, oppextract2 = argv

path = r"D:\Verdia Pty Ltd\Systems Tools and IT - Salesforce\Data Extracts\Opportunities"
os.chdir(path)
print os.getcwd()


#import both pipeline snapshots to matrix
pipeline_snapshot1 = import_opportunities(oppextract1)
pipeline_snapshot2 = import_opportunities(oppextract2)

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
            i[2] = k['NAME']
            i[3] = k['CLOSEDATE']
            i[4] = k['STAGENAME']
            i[5] = float(k['VD_PROJECT_VALUE__C'])
            i[6] = float(k['PROBABILITY'])
            i[7] = float(k['WEIGHTED_PROJECT_VALUE__C'])


    for j in pipeline_snapshot2:
        if i[0] == j['ID']:
            i[11] = j['NAME']
            i[12] = j['CLOSEDATE']
            i[13] = j['STAGENAME']
            i[14] = float(j['VD_PROJECT_VALUE__C'])
            i[15] = float(j['PROBABILITY'])
            i[16] = float(j['WEIGHTED_PROJECT_VALUE__C'])
        else:
            pass

for i in excel_matrix:
    if i[8] == 'Status': #i[8] == 'Removed from Pipeline' or
        pass
    elif i[8] == 'Removed from Pipeline':
        i[18] = -i[5]
        i[14] = ""
        i[17] = ""
        i[19] = ""
        i[20] = ""
    else:
        if i[4] != i[13]:
            i[17] = i[13]
        else:
            i[17] = ""
        if i[5] != i[14]:
            i[18] = i[14] - i[5]
        else:
            i[18] = ""
        if i[6] != i[15]:
            i[19] = i[15]
        else:
            i[19] = ""
        if i[7] != i[13]:
            i[17] = i[13]
        else:
            i[17] = ""
        if i[3] != i[12]:
            i[20] = "Date Changed"
        else:
            i[20] = ""

#print len(excel_matrix)
#print excel_matrix



#identify projects that have exited the pipeline because they were moved to Closed lost or Closed Won
#identify any Projects thats have been modified between snapshots


#identify what changes have occurred since last snapshot e.g. stage change, close date change, dollar value change etc. 

#create a new report spreadsheet - titles pipeline comparison report (Date 1 v Date 2)

path = "C:\Users\Mitchell.Dawson\Desktop"
os.chdir(path)
workbook_filename = "Pipeline_Comparison.xlsx"
workbook = xlsxwriter.Workbook(workbook_filename)
worksheet = workbook.add_worksheet('Pipeline Comparison')


write_to_excel(excel_matrix,0,0,worksheet,workbook)


workbook.close()



#copy both snapshots into new spreadsheet on individual tabs 
















