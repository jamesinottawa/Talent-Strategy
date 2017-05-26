# This file loads in the excel spreadsheets
#
# Author: James Chapman
# Creation Date : 20715-05-16

# load the libraries
import pandas as pd
from glob import glob
import re


# this is the directory to look into
# applications = glob("O:/Talent Strategy - To process/Conference Applications/2017Q2/*.xlsm")
applications = glob("C:/Users/chja/Desktop/2017Q2/*.xlsm")
quarter = "2017Q2"
save_file = "C:/Users/chja/Desktop/temp_ts.xlsx"

def read_sheet(file_name,quarter):
    # load the data into memory
    # Now remove the non-conforming entries
    z = pd.read_excel(file_name,sheetname="questionnaire",skiprows=7,header=None)
    conference = z.ix[9:14, [1, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17]]
    # This tests if the conference is null
    x = pd.isnull(conference[[1, 4, 5, 6, 7]]).all(axis=1)
    # This tests if it accidentalliy added other lines
    y = conference[1].str.contains("Explain") | conference[1].str.contains("Rational")
    conference = conference.ix[~(x|y)]
    conference.columns = ["Conf_name", "start_date", "end_date", "location", "conf_link", "accepted", "departure_date",
                          "return_date", "days", "transport", "hotel", "meals", "registration", "other", "total"]
    conference['authors'] = z.ix[0, 4]
    last_name = z.ix[1,4].strip().split(" ")[1]
    conference['applicant'] = z.ix[1, 4]
    conference['applicant_email'] = z.ix[2, 4]
    conference['submission_date'] = z.ix[0, 13]
    conference['dept'] = z.ix[1, 13]
    conference['manager'] = z.ix[2, 13]
    conference['paper_title'] = z.ix[4,4]
    # Now create links to the papers
    paper = z.ix[5,4]
    if ":" not in paper:
        paper = '=HYPERLINK("O:/Talent Strategy - Supporting Files/'+quarter+"/"+last_name+"/"+paper+'","link")'
    else:
        paper = '=HYPERLINK("'+paper+'","link")'
    conference['paper_link'] = paper
    conference['times_approves'] = z.ix[6,16]
    conference['recent_quarter'] = z.ix[6,16]
    return(conference)

# Now find all the files and loop over them
app_data = []
for file in applications:
    app_data.append(read_sheet(file,quarter))
app_data = pd.concat(app_data)

# Now reorder the columns and rename them to a friendly version
app_data_order = ['dept', 'authors','paper_title', 'Conf_name','conf_link', 'location', 'departure_date',
                  'start_date', 'end_date',  'return_date', 'days',
                  'accepted', 'total', 'transport', 'hotel', 'meals',
                  'registration', 'other', 'applicant',
                  'applicant_email', 'submission_date',  'manager',
                   'paper_link', 'times_approves',
                  'recent_quarter']
header_names = ['Dept', 'Authors','Paper', 'Conference', 'Conference Link','Location', 'Departure',
                  'Start', 'End',  'Return',  'Days',
                  'Accepted', 'Total', 'Transport', 'Hotel', 'Meals',
                  'Registration', 'Other', 'Applicant',
                  'applicant_email', 'submission_date',  'manager',
                   'paper_link', 'times_approves',
                  'recent_quarter']
app_data = app_data[app_data_order]

# Now start an XlsxWriter instance
writer = pd.ExcelWriter(save_file,engine="xlsxwriter")
app_data.to_excel(writer,sheet_name="applications",header=header_names)
workbook = writer.book
writer.save()

