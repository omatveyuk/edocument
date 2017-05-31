##############################################################################
#
# A simple program to write some data to an Excel file using the XlsxWriter
# Python module.
#
# This program is shown, with explanations, in Tutorial 1 of the XlsxWriter
# documentation.
#
# Copyright 2013-2016, John McNamara, jmcnamara@cpan.org
#
import calendar

import xlsxwriter
import json
import os
from dateutil import parser
# from textblob.classifiers import NaiveBayesClassifier

from datetime import datetime
from titlecase import titlecase
from collections import OrderedDict
import unicodedata

def cell(l):
    start_month = 4
    start_year = 2016

    end_month = 5
    end_year = 2015


    m = l[0]

    start_month - m
    y = l[1]




    stats = [[] for i in range(12)]
    y = start_year
    m = start_month
    for i in range(12):
        m = start_month - i

        if m == 0:
            m = 12
            y = start_year - 1

        stats[i] = y

def xcapitalize(s):
    not_these = ['a', 'the', 'of', 'and', 'or', 'is', 'an']
    print (' '.join(word
                   if word in not_these
                   else word.title()
                   for word in s.capitalize().split(' ')))


def format_url(color):
    return {
        'font_color': 'blue',
        'fg_color': color,
        'underline': 1,
        'border': 1
    }


def format_date(color):
    return {'num_format': 'yyyy-mm-dd', 'fg_color': color, 'border': 1}

def extract_year(p1):
    tmp = ''
    parts = p1["Parts"]
    for p in parts[1:]:
        key = p["Key"]
        val = p["Value"]


        if key == "i":
            #tmp += " year:"
            tmp += val
    tmp = tmp.strip()

    return tmp

def format_issue(p1):
    tmp = ''
    parts = p1["Parts"]
    for p in parts[1:]:
        key = p["Key"]
        val = p["Value"]

        if key == 'a':
            tmp += " v."
        elif key == 'b':
            tmp += " no:"
        elif key == "i":
            tmp += " year:"
        elif key == "j":
            tmp += " month:"
        elif key == "k":
            tmp += " day:"

        tmp += val
    tmp = tmp.strip()

    return tmp

def statistic_last_12_months(issues):
    current_month = datetime.now().month
    current_year = datetime.now().year
    # create list of statistic where item[0] - statistic current month and year
    # item[1] - statistic previous month of current date
    # item[2] - statistic previous previous month of current day
    # item[3] - etc.
    # item[11] - 11 months ago
    statistic = [0]*12
    for issue in issues:
        date_issue = parser.parse(issue["Timestamp"]).date()
        month_issue = date_issue.month
        year_issue =  date_issue.year
        if month_issue <= current_month and year_issue == current_year:
            statistic[current_month-month_issue] += 1
        else:
            if month_issue > current_month and year_issue == current_year-1:
                statistic[current_month-month_issue+12] += 1
    return statistic


rootdir = os.path.dirname(os.path.realpath(__file__))



words = 'The quick brown fox jumps over the lazy dog'.split()
print words
stuff = [[w.upper(), w.lower(), len(w)] for w in words if w.startswith('t') or w.startswith('T')]
print stuff


stuff = map(lambda w : w.startswith('T'), words)
print stuff

list = []
max_date = datetime.min.date()
files = [each for each in os.listdir(rootdir + "/latest/") if each.endswith('.json')]

for file in files:
    #if not file.startswith("3573"):
    #    continue

    data = json.load(open(rootdir + "/latest/" + file))
    # print file
    issues = data["Issues"]
    pubFrequency = data["PublicationFrequency"]
    issues = sorted(issues, key=lambda k: k['Timestamp'], reverse=True)

    # this will give me a list of all issues
    l = [[parser.parse(i["Timestamp"]).date().month, parser.parse(i["Timestamp"]).date().year] for i in issues]
    parts = [ i["Parts"] for i in issues if i["Parts"][0].get("Key") == "8" ]

    index = [i for i in range(len(issues))]
    max = [0,0]
    latest_index = 0

    for i in index:
        parts = [p for p in range(len(issues[i]["Parts"])) ]
        for p in parts:
            if (issues[i]["Parts"][p]).get("Key") == "8":
                #numbers = int((issues[i]["Parts"][p]).get("Value").split(".") #[1])
                numbers = (issues[i]["Parts"][p]).get("Value").split(".")

                #val = "".join(str(x) for x in numbers)
                #val = val.ljust(5,'0')
                #val = int(val)
                val = numbers


                #print i, p, val
                if (int(max[0]) == int(val[0]) and int(max[1]) < int(val[1])) or (int(max[0]) < int(val[0])):
                    max = val
                    latest_index = i

    #print "result", latest_index, len(issues), max
    #print issues[latest_index]

    #print issues[0]

    #if (issues[latest_index] == issues[0]):
    #    print "stop: Problem detected in " + file

    # part = max([i[0].get("Value") for i in parts if i[0]].get("Key") == "8"])


    for i in l:
        pos = cell(i)
    countIssues = 0

    date_object = "n/a"
    formatted_issue = "n/a"
    string_date = date_object
    issues_last_12_months = [ 0 for i in range(12)]

    if issues:
        string_date = issues[0]["Timestamp"]
        date_object = parser.parse(string_date).date()
        if max_date < date_object:
            max_date = date_object
        formatted_issue = format_issue(issues[latest_index]) #not sure what this should be? 0 or something else
        countIssues = len(issues)
        latest_issue_year = extract_year(issues[latest_index])
        issues_last_12_months = statistic_last_12_months(issues)
#        print("issues for " + file + " " +  str(date_object))

    else:

        print ("---> info missing: " + file + " " + date_object)

    joe = (titlecase(data["CatalogTitle"]), date_object, data["StanfordLibraryId"], data["CallNumber"], formatted_issue,
           pubFrequency, countIssues, string_date, latest_issue_year, issues_last_12_months)

    list.append(joe)

# sort list by last issue date
list = sorted(list, key=lambda item: item[7], reverse=True)

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('Expenses01.xlsx')
worksheet = workbook.add_worksheet()

# Some data we want to write to the worksheet.
expenses = (
    ['Rent', 1000],
    ['Gas', 100],
    ['Food', 300],
    ['Gym', 50],
)

# Start from the first cell. Rows and columns are zero indexed.
row = 0
col = 0

grey = 'F2F7FF'

# Write name of columns and their width
header = workbook.add_format({'bold': True,
                              'align': 'center',
                              'fg_color': '#D7E4BC',
                              'border': 1})  # Implicit format.
grey_row = workbook.add_format({'fg_color': grey})
worksheet.set_row(0, 20, header)
worksheet.freeze_panes(1, 0)
worksheet.write(row, col, "URL")
worksheet.set_column(col, col, 10)
worksheet.write(row, col + 1, "ID")
worksheet.write(row, col + 2, "Title")
worksheet.set_column(col + 2, col + 2, 50)
worksheet.write(row, col + 3, "Latest issue year")
worksheet.set_column(col + 3, col + 3, 20)


worksheet.write(row, col + 4, "Last issue")
worksheet.set_column(col + 4, col + 4, 30)
worksheet.write(row, col + 5, "Received on")
worksheet.set_column(col + 5, col + 5, 10)
worksheet.write(row, col + 6, "CALL Number")
worksheet.write(row, col + 7, "Frequency")
worksheet.set_column(col + 7, col + 7, 20)
worksheet.write(row, col + 8, "Total")
current_month = datetime.now().month
current_year = datetime.now().year
for i in range(0, 12):
    if current_month-i > 0:
        month_year = calendar.month_abbr[current_month-i] + "-" + str(current_year)
    else:
        month_year = calendar.month_abbr[current_month-i+12] + "-" + str(current_year-1)
    worksheet.write(row, col + 8 + i + 1, month_year)
row += 1

# Iterate over the data and write it out row by row.
for item in (list):
    cell_color = "#FFFFFF"
    if item[1] == max_date:
        cell_color = grey
    worksheet.set_row(row, None, workbook.add_format({'fg_color': cell_color,
                                                      'border': 1}))

    worksheet.write_url(row, col, 'https://searchworks.stanford.edu/catalog/{0}/librarian_view'.format(item[2]),
                        workbook.add_format(format_url(cell_color)), "SearchWorks")

    worksheet.write(row, col + 1, item[2])
    worksheet.write(row, col + 2, item[0])
    worksheet.write(row, col + 3, item[8])
    worksheet.write(row, col + 4, item[4])
    worksheet.write(row, col + 5, item[1], workbook.add_format(format_date(cell_color)))
    worksheet.write(row, col + 6, item[3], )
    worksheet.write(row, col + 7, "".join(item[5]))
    worksheet.write(row, col + 8, item[6])

    for i in range(0,12):
        worksheet.write(row, col + 8 + i + 1, item[9][i])
    row += 1

workbook.close()
