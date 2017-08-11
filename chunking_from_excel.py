# -*- coding: utf-8 -*-
"""
Created on Thu Aug  3 10:25:05 2017

@author: Jose Luna

Reads from 1 excel file, 'Chunking Version 5.xlsx', with 3 sheets:
    'SFDC.prjCapitalLaborSynced', 'SFDC.resourcePoolSynced', 'Rate.Synced'
Returns a table called cas_costs with 3 columns:
    CPR-CAS Name, CAS Owner, Total

To produce results for any fiscal year, change current_year variable to
a string of the form 'year1, year2', ie. '1516', '1617', '1718', etc.

If finance is set to 'y', resources with CCX as their Cost Center
will not be taken into account in the calculations.
If finance is set to 'n', all resources will be counted.

Any changes from the CSV version are commented out.
"""

import pandas as pd
from datetime import datetime as dt
from datetime import date
from timeit import default_timer

#Not necessary; just for testing efficiency
start = default_timer()

raw_projects = pd.read_excel('Chunking Version 5.xlsx', sheetname='SFDC.prjCapitalLaborSynced')
#raw_projects = pd.read_csv('projects.csv')

raw_resources = pd.read_excel('Chunking Version 5.xlsx', sheetname='SFDC.resourcePoolSynced')
#raw_resources = pd.read_csv('resources.csv')

rates = pd.read_excel('Chunking Version 5.xlsx', sheetname='Rate.Synced')
#rates = pd.read_csv('rates.csv')

#====================== VARIABLES =============================================

monthly_rate = rates['Fully Burdened Monthly Cost.amount'][0]

#FINANCE VS TOTAL
#If finance is set to 'y', resources with CCX as their Cost Center
#will not be taken into account in the calculations.
#If finance is set to 'n', all resources will be counted.
finance = 'y'

#DEFINE FISCAL YEAR
#To produce results for any fiscal year, change current_year variable to
#a string of the form 'year1, year2', ie. '1516', '1617', '1718', etc.
current_year = '1718'
year1, year2 = current_year[:2], current_year[2:]

#DICTIONARY FOR MONTH DELIMITERS
#Key is month number, value is end day of month

month_ends = {1:31, 2:28, 3:31, 4:30, 5:31, 6:30, 7:31, 8:31, 9:30,
    10:31, 11:30, 12:31}

#====================== GENERAL UTILITY FUNCTIONS =============================
    
def transformColumn(table, column, function, *args):
    '''General-use. Takes a table, a column name in that table, and a function
    and applies that function destructively to every element in the column
    table: dataframe; column: series'''
    if len(args) > 0:
        table[column] = table[column].apply(function, axis=args[0])
    else:
        table[column] = table[column].apply(function)
    
def transformMultipleColumns(table, column_list, function, *args):
    '''General-use. Takes a table, a list of column names in that table,
    and afunction and applies that function destructively to every elemnt
    in each of the columns.'''
    for col in column_list:
        transformColumn(table, col, function, *args)
    
def addEmptyColumn(table, column_name):
    '''Given a table and column name, adds an empty column with that name'''
    table['%s' % column_name] = ""
    
def addMultipleColumns(table, column_list):
    '''Given a table and list of column names, adds an empty column for
    every name in the column list'''
    for col in column_list:
        addEmptyColumn(table, col)

def applyRowTransform(table, column_list, function):
    '''For each column name string in column_list, applies a row-wise function
    to populate the column. Used to generate any columns that depend on previous columns.
    Function must take 2 parameters; 1 for row and 1 for column.'''
    for col in column_list:
        table[col] = table.apply(lambda row: function(row, col), axis = 1)

#====================== DATE FUNCTIONS ========================================

def convertDates(date):
    '''Takes a date in string format (MM/DD/YYYY) and converts it to
    a date object with format YYYY/MM/DD. Checks for NaN values and sets to None'''
    #if date != date:
        #return None
    if isinstance(date, str):
        return dt.strptime(date, "%m/%d/%Y").date()
    else:
        #return date
        return date.date()

def convertColumnDates(table, column):
    '''Specific-use. Takes a table and colum name and applies the convertDates
    function to every element in the column
    table: dataframe; column: series'''
    transformColumn(table, column, convertDates)

def createMonths(year):
    '''Returns a list of 12 months, formatted as Month/Year strings (
    ie. 7/17, 8/17, ... 6/18) depending on current_year'''
    months = []
    for i in range(7, 13):
        months.append('{0}/{1}'.format(i, year1))
    for j in range(1, 7):
        months.append('{0}/{1}'.format(j, year2))
    return months

def parseMonth(mon_year):
    '''Receives a string of the form Month/Year and returns
    an integer for the month number (1 to 12)'''
    spl = mon_year.split('/')
    month = spl[0]
    month = month.lstrip("0")
    return int(month)

def parseYear(mon_year):
    '''Receives a string of the form Month/Year and returns
    an integer for the year number (ie. 2012)'''
    spl = mon_year.split('/')
    year = '20' + spl[1]
    return int(year)

def monthStartEnd(mon_year, start_end):
    '''Receives a string of the form Month/Year and a string which
    is either 's' or 'e'and returns a date object for either
    the first or last day of that month, based on the start_end string.'''
    month = parseMonth(mon_year)
    year = parseYear(mon_year)
    if start_end == 's':
        day = 1
    elif start_end == 'e':
        day = month_ends[month]
    return date(year, month, day)

def rangeOverlap(start1, start2, end1, end2):
    '''General Use. Finds overlapping range (in days) between 2 date ranges:
    (start1, end1), (start2, end2). Returns 0 if no overlap.
    All parameters are date objects'''
    max_start = max(start1, start2)
    min_end = min(end1, end2)
    delta = min_end - max_start
    days = delta.days + 1
    if days <= 0:
        return 0
    return abs(days)

#====================== SPECIFIC PROJECT FUNCTIONS ============================

def projWeights(proj_start, proj_end, mon_year):
    '''Finds overlapping range between project and month
    on a scale from 0 (no overlap) to 1 (full overlap).
    If there is no proj_start, it will be set to Jan 1st 1900
    If there is no proj_end, it will be set to Dec 31st 2099
    proj_start & proj_end: datetime; mon_year: string'''
    if (proj_start == None):
        proj_start = date(1900, 1, 1)
    if (proj_end == None):
        proj_end = date(2099, 12, 31)
    month_start = monthStartEnd(mon_year,'s')
    month_end = monthStartEnd(mon_year, 'e')
    return rangeOverlap(proj_start, month_start,
                        proj_end, month_end)/month_end.day
    
def projRowWeights(row, mon):
    '''Given a row Series and month string, returns the overlap weight between
    that row's proj_start, proj_end and the month'''
    proj_start = row['Start']
    proj_end = row['End']
    return projWeights(proj_start, proj_end, mon)
        
def generateProjectWeights():
    '''Populates the month columns in project_resources with the
    overlap-weight of each cell. See overlap() for more detail'''
    applyRowTransform(project_resources, months, projRowWeights)
        
    
def weightCostMultiplier(row, mon):
    proj_month_weight = row[mon]
    res_month_cost = row[mon + '_r']
    return proj_month_weight * res_month_cost
    
def generatePRCosts():
    costs = pd.merge(project_resources, resources,
                          how='left', on='Resource ID', suffixes = ('','_r'))
    applyRowTransform(costs, months, weightCostMultiplier)
    return costs

#====================== SPECIFIC RESOURCE FUNCTIONS ===========================
        
def resourceActive(res_start, res_end, mon_year):
    '''Given date obects: res_start, res_end and a mon_year string
    of the form Month/Year; returns true if the resource is active
    during that month, and false otherwise'''
    year = parseYear(mon_year)
    month = parseMonth(mon_year)
    month_middle = date(year, month, 15)
    return ((res_start <= month_middle) & (res_end >= month_middle))
        
def resourceMonthRate(row, mon):
    '''For each resource row and month, returns the monthly project rate
    by dividing the average monthly rate by the sum of project weights
    (+1 for Opex-Capex) for the corresponding resource ID in project_resources
    If there is no res_start, it will be set to Jan 1st 1900
    If there is no res_end, it will be set to Dec 31st 2099'''
    res_ID = row['Resource ID']
    res_start = row['Start']
    res_end = row['End']
    res_cc = row['Cost Center']
    res_alloc = row['Cost Allocation']
    if res_start == None:
        res_start == date(1900, 1, 1)
    if res_end == None:
        res_end = date(2099, 12, 31)
    if not (resourceActive(res_start, res_end, mon)):
        return 0
    if (res_cc == 'CCX') & (finance == 'y'):
        return 0
    proj_count = project_resources[project_resources['Resource ID'] == res_ID][mon].sum()
    if proj_count == 0:
        return 0
    if res_alloc == 'Opex-Capex':
        return monthly_rate/(proj_count + 1)
    if (res_alloc == '100% Capex') | (res_alloc == '-'):
        return monthly_rate/proj_count
    return 0

def generateResourceMonthRates():
    '''Populates the month columns in resources with the average cost per
    project-month for this resource'''
    applyRowTransform(resources, months, resourceMonthRate)

#================== DATA PREPROCESSING / SETTING UP TABLES ====================
    
#CREATE MONTH COLUMNS
months = createMonths(current_year)
    
#CONVERT PROJECT + RESOURCE START AND END DATES
for col in ['Project: Execution Start', 'Project: In Service/Actual End Date']:
    convertColumnDates(raw_projects, col)

for col in ['Start MM-DD-YYYY', 'End MM-DD-YYYY']:
    convertColumnDates(raw_resources, col)
    
#CREATE TABLES
project_resources = pd.DataFrame(
        data = raw_projects[[
                'Project: Project ID',
                'Project: Project Name',
                'Project: Execution Start',
                'Project: In Service/Actual End Date',
                'Project: CPR+Name',
                'Project: CAS Owner',
                'Contact-ID',
                'Full Name'
                ]].copy())
    
resources = pd.DataFrame(
        data = raw_resources[[
                'Full Name',
                'Contact-ID',
                'Cost Center',
                'Cost Allocation',
                'Start MM-DD-YYYY',
                'End MM-DD-YYYY'
                ]].copy())
    
#ADD MONTH COLUMNS
addMultipleColumns(project_resources, months)
addMultipleColumns(resources, months)

#RENAME COLUMNS
project_renaming = {'Project: Project ID':'ID',
             'Project: Project Name':'Name',
             'Project: Execution Start':'Start',
             'Project: In Service/Actual End Date':'End',
             'Project: CPR+Name': 'CAS',
             'Project: CAS Owner': 'CAS Owner',
             'Contact-ID':'Resource ID',
             'Full Name':'Resource Name'}
resource_renaming = {'Contact-ID':'Resource ID',
             'Start MM-DD-YYYY':'Start',
             'End MM-DD-YYYY':'End'}

project_resources.rename(columns=project_renaming, inplace=True)
resources.rename(columns = resource_renaming, inplace=True)

#============================ CALCULATIONS ====================================

#POPULATE PROJECT-RESOURCES AND RESOURCES TABLES
generateProjectWeights()
generateResourceMonthRates()

#CREATE PROJECT-RESOURCE COST TABLE / CLEAN UP / ADD TOTAL COLUMN
pr_costs = generatePRCosts()
pr_costs = pr_costs.drop([m + '_r' for m in months], axis = 1)
pr_costs = pr_costs.drop(['Full Name', 'Cost Center', 'Cost Allocation',
         'Start_r', 'End_r'], axis = 1)
    
#ADD TOTAL COLUMN
pr_costs['Total'] = pr_costs[months].sum(axis=1)

#AGGREGATE INTO TOTAL PROJECT COST TABLE
proj_costs = pr_costs.groupby(['Name', 'ID', 'CAS', 'CAS Owner'], as_index = False)['Total'].agg('sum')

#AGGREGATE INTO TOTAL CAS COST TABLE
cas_costs = proj_costs.groupby(['CAS', 'CAS Owner'], as_index = False)['Total'].agg('sum')
cas_costs.loc[len(cas_costs)] = ['Grand Total', '—', cas_costs['Total'].sum()] 
cas_costs.set_index('CAS', inplace = True)

#EXPORT TO CSV -- for later.
#cas_costs.to_csv('CAS Costs.csv')

#PRINT TIME TO RUN
duration = default_timer() - start
print('Time to Run: {0:.2f} seconds'.format(duration))

