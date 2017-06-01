"""Daily report periodicals inventory"""
""" some comment """
import calendar
import sys
import xlsxwriter
import json
import os, getopt
from dateutil import parser
from datetime import datetime
from titlecase import titlecase
from collections import OrderedDict
import unicodedata
from Timer import Timer

def format_url(color):
    """Format cell 'url' using color
       color: cell color (grey means most recent)
    """
    return {
        'font_color': 'blue',
        'fg_color': color,
        'underline': 1,
        'border': 1
    }


def format_date(color):
    """Format cell 'date' using color
       color: cell color (grey means most recent)
    """
    return {'num_format': 'yyyy-mm-dd', 'fg_color': color, 'border': 1}


def extract_year(parsed_part):
    """Extract year from parsed data structure
       parsed_part: item from data["Issues"] list
       Return year of latest issue
    """

    parts = parsed_part["Parts"]
    # for p in parts[1:]:
    for p in parts:
        if p["Key"] == "i":         # year
            return p["Value"].strip()


def format_issue(parsed_part):
    """Format issue's data.
       parsed_part: item from data["Issues"] list
       Return formatted string
    """
    str_issue = ''
    parts = parsed_part["Parts"]

    for p in parts[1:]:
        if p["Key"] == 'a':
            str_issue += " v."
        elif p["Key"] == 'b':
            str_issue += " no:"
        elif p["Key"] == "i":
            str_issue += " year:"
        elif p["Key"] == "j":
            str_issue += " month:"
        elif p["Key"] == "k":
            str_issue += " day:"
        str_issue += p["Value"]

    return str_issue.strip()


def statistic_last_12_months(issues):
    """Return  list of statistic data for last 12 months where
       [0] - statistic current month and year
       [1] - statistic previous month of current date
       [2] - statistic previous previous month of current day
       [3] - etc.
       [11] - 11 months ago
    """
    current_month = datetime.now().month
    current_year = datetime.now().year
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


def create_report_periodicals(path):
    """Return list with information for each periodical and data of latest update"""
    if path == '':
        path = os.path.dirname(os.path.realpath(__file__)) + "/latest/"

    report_periodicals = []
    latest_update = datetime.min.date()
    files = [each for each in os.listdir(path) if each.endswith('.json')]

    for f in files:
        with open(path + f) as json_file:
            with Timer(True) as t:
                t.__enter__()
                data = json.load(json_file)
                t.__exit__()

            issues = data["Issues"]
            pubFrequency = data["PublicationFrequency"]

            with Timer(True) as t:
                t.__enter__()
                issues = sorted(issues, key=lambda k: k['Timestamp'], reverse=True)
                # Find latest index
                max_val=[0,0]
                latest_index = 0
                for i in xrange(len(issues)):
                    for p in xrange(len(issues[i]["Parts"])):
                        if issues[i]["Parts"][p]["Key"] == "8":
                            numbers = (issues[i]["Parts"][p]).get("Value").split(".")
                            val = issues[i]["Parts"][p]["Value"].split(".")
                            val = numbers
                            if (int(max_val[0]) == int(val[0]) and int(max_val[1]) < int(val[1])) or (int(max_val[0]) < int(val[0])):
                                max_val = val
                                latest_index = i

                date_object = "n/a"
                formatted_issue = "n/a"
                string_date = date_object
                issues_last_12_months = [ 0 for i in range(12)]
                number_issues = len(issues)

                if issues:
                    string_date = issues[0]["Timestamp"]
                    date_object = parser.parse(string_date).date()
                    if latest_update < date_object:
                        latest_update = date_object
                    formatted_issue = format_issue(issues[latest_index])
                    number_issues = len(issues)
                    latest_issue_year = extract_year(issues[latest_index])
                    issues_last_12_months = statistic_last_12_months(issues)
                else:
                    print ("---> info missing: " + f + " " + date_object)

                row = (titlecase(data["CatalogTitle"]), date_object, data["StanfordLibraryId"], data["CallNumber"], formatted_issue,
                       pubFrequency, number_issues, string_date, latest_issue_year, issues_last_12_months)

                report_periodicals.append(row)
                t.__exit__()

    # sort list by last issue date
    with Timer(True) as t:
        t.__enter__()
        report_periodicals = sorted(report_periodicals, key=lambda item: item[7], reverse=True)
        t.__exit__()
    return [report_periodicals, latest_update]

def create_xls(outputdir, report_periodicals, latest_update):
    """Create a workbook and add a worksheet."""
    if outputdir == '':
        workbook = xlsxwriter.Workbook('stanford_daily_report.xlsx')
    else:
        workbook = xlsxwriter.Workbook(outputdir + 'stanford_daily_report.xlsx')
    worksheet = workbook.add_worksheet()

    # Start from the first cell. Rows and columns are zero indexed.
    row = 0
    col = 0

    grey = '#F2F7FF'

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
    for item in (report_periodicals):
        cell_color = "#FFFFFF"
        if item[1] == latest_update:
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


def create_html(outputdir,report_periodicals, latest_update):
    if outputdir == '':
        html_file = 'stanford_daily_report.html'
    else:
        html_file = outputdir + 'stanford_daily_report.html'
    
    with open(html_file, 'w') as outfile:
        with open('header_report.html') as infile:
                outfile.write(infile.read())

    with open(html_file, 'a+') as outfile:
        current_month = datetime.now().month
        current_year = datetime.now().year
        html = ''
        for i in range(0, 12):
            if current_month-i > 0:
                month_year = calendar.month_abbr[current_month-i] + "-" + str(current_year)
            else:
                month_year = calendar.month_abbr[current_month-i+12] + "-" + str(current_year-1)

            html += '<th>' + month_year + '</th>'

        html += '</tr></thead><tbody>'
        outfile.write(html)

        # Iterate over the data and write it out row by row.
        for item in (report_periodicals):
            if item[1] == latest_update:
                html = '<tr style="background: #FCE9E9">'
            else:
                html = '<tr>'
            html += '<td><a href="https://searchworks.stanford.edu/catalog/{0}/librarian_view"'.format(item[2])
            html += '>' + str(item[2]) + '</a></td>'
            html += '<td>' + item[0] + '</td>'
            if item[8] is None:
                html += '<td> </td>'
            else:
                html += '<td>' + item[8] + '</td>'
            html += '<td>' + item[4] + '</td>'
            html += '<td>' + str(item[1]) + '</td>'
            html += '<td>' + item[3] + '</td>'
            html += '<td>' + "".join(item[5]) + '</td>'
            html += '<td>' + str(item[6]) + '</td>'

            for i in range(0,12):
                html += '<td>' + str(item[9][i]) + '</td>'
            
            html += '</tr>'
            outfile.write(html.encode('utf-8'))


        with open('footer_report.html') as infile:
                outfile.write(infile.read())


def main(argv):
    total = len(sys.argv)
    cmdargs = str(sys.argv)

    # Read command line args
    inputdir = ''
    outputdir = ''
    try:
        opts, args = getopt.getopt(argv, "hi:o:", ["idir=", "odir="])
    except getopt.GetoptError:
        print
        'test.py -i <inputdir> -o <outputdir>'
        sys.exit(2)
    for opt, arg in opts:
        if opt == '-h':
            print 'test.py -i <inputdir> -o <outputdir>'
            sys.exit()
        elif opt in ("-i", "--idir"):
            inputdir = arg
        elif opt in ("-o", "--odir"):
            outputdir = arg
    print 'Input directory is "', inputdir
    print 'Output directory is "', outputdir

    with Timer(True) as t:
        t.__enter__()
        report_periodicals, latest_update = create_report_periodicals(inputdir)
        create_xls(outputdir, report_periodicals, latest_update)
        create_html(outputdir, report_periodicals, latest_update)
        t.__exit__();

if __name__ == "__main__":
    main(sys.argv[1:])
