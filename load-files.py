# This file loads in the excel spreadsheets
#
# Author: James Chapman
# Creation Date : 20715-05-16

# load the libraries
import pandas as pd
from glob import glob

# this is the directory to look into
applications = glob("O:/Talent Strategy - To process/Conference Applications/2017Q2/*.xlsm")
save_file = "C:/Users/chja/Desktop/temp_ts.xlsx"

def read_sheet(file_name):
    # load the data into memory
    # Now remove the non-conforming entries
    z = pd.read_excel(file_name,sheetname="questionnaire",skiprows=7,header=None)
    conference = z.ix[9:14, [1, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17]]
    x = pd.isnull(conference[[1, 4, 5, 6, 7]]).all(axis=1)
    conference = conference.ix[~x]
    conference.columns = ["Conf_name", "start_date", "end_date", "location", "conf_link", "accepted", "departure_date",
                          "return_date", "days", "transport", "hotel", "meals", "registration", "other", "total"]
    conference['authors'] = z.ix[0, 4]
    conference['applicant'] = z.ix[1, 4]
    conference['applicant_email'] = z.ix[2, 4]
    conference['submission_date'] = z.ix[0, 13]
    conference['dept'] = z.ix[1, 13]
    conference['manager'] = z.ix[2, 13]
    conference['paper_title'] = z.ix[4,4]
    conference['paper_link'] = z.ix[5,4]
    conference['times_approves'] = z.ix[6,16]
    conference['recent_quarter'] = z.ix[6,16]
    return(conference)

# Now find all the files and loop over them
app_data = []
for file in applications:
    app_data.append(read_sheet(file))
app_data = pd.concat(app_data)

# Now reorder the columns and rename them to a friendly version
app_data_order = ['dept', 'authors','paper_title', 'Conf_name', 'location', 'departure_date',
                  'start_date', 'end_date',  'return_date', 'conf_link',
                  'accepted', 'total', 'days', 'transport', 'hotel', 'meals',
                  'registration', 'other', 'applicant',
                  'applicant_email', 'submission_date',  'manager',
                   'paper_link', 'times_approves',
                  'recent_quarter']
header_names = ['Dept', 'Authors','Paper', 'Conference', 'Location', 'Departure',
                  'Start', 'End',  'Return', 'Conference Link',
                  'Accepted', 'Total', 'Days', 'Transport', 'Hotel', 'Meals',
                  'Registration', 'Other', 'Applicant',
                  'applicant_email', 'submission_date',  'manager',
                   'paper_link', 'times_approves',
                  'recent_quarter']
app_data = app_data[app_data_order]


app_data.to_excel(save_file,"applications",header=header_names)
