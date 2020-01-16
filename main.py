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
matrixheight = len(deleted_opps) + len(new_opps)+ len(existing_opps)
print matrixheight

excel_matrix = matrix_creator(matrixheight)
print excel_matrix


#identify projects that have exited the pipeline because they were moved to Closed lost or Closed Won
#identify any Projects thats have been modified between snapshots


#identify what changes have occurred since last snapshot e.g. stage change, close date change, dollar value change etc. 

#create a new report spreadsheet - titles pipeline comparison report (Date 1 v Date 2)
"""
path = "C:\Users\Mitchell.Dawson\Desktop"
os.chdir(path)
workbook_filename = "Pipeline_Comparison.xlsx"
workbook = xlsxwriter.Workbook(workbook_filename)

workbook.close()

"""
#copy both snapshots into new spreadsheet on individual tabs 
















