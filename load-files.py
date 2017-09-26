# This file loads in the excel spreadsheets
#
# Author: James Chapman
# Creation Date : 2017-05-16

# load the libraries
import pandas as pd
from glob import glob

# this is the directory to look into
# applications = glob("O:/Talent Strategy - To process/Conference Applications/2017Q2/*.xlsm")
applications = glob("O:\Talent Strategy - To process\Conference Applications/2017Q3/*.xlsm")
dsas = glob("O:\Talent Strategy - To process\DSA Applications/2017Q3/*.xlsm")
quarter = "2017Q3"
save_file = "C:/Users/chja/Desktop/new_temp_ts_2017Q3.xlsx"


def read_dsa_sheet(file_name,quarter):
    # load the data into memory
    # Now remove the non-conforming entries
    z = pd.read_excel(file_name, sheetname="questionnaire",
                      skiprows=7, header=None)
    visits = z.loc[11:15,1:13]
    # This tests if the visit is null
    visits.columns = ["visit_number","visit_start","visit_end","work_days",
                      "working_transit","per_diem","total_honorarium",
                      "transport","hotel","meal","hospitality","other","total"]
    x = visits[["visit_start", "visit_end"]].isnull().all(axis=1)
    visits = visits.loc[~x]
    visits["applicant"] = z.loc[0,3]
    visits["department"] = z.loc[1,10]
    visits["DSA_name"] = z.loc[4,3]
    return(visits)

def read_conference_sheet(file_name,quarter):
    # load the data into memory
    # Now remove the non-conforming entries
    z = pd.read_excel(file_name,sheetname="questionnaire",
                      skiprows=7,header=None)
    conference = z.loc[9:14, [1, 4, 5, 6, 7, 
                              8, 9, 10, 11,
                              12, 13, 14, 15, 
                              16, 17]]
    # This tests if the conference is null
    x = pd.isnull(conference[[1, 4, 5, 6, 7]]).all(axis=1)
    # This tests if it accidentalliy added other lines below the trips
    y = conference[1].str.contains("Explain") | conference[1].str.contains("Rational")
    conference = conference.loc[~(x|y)]
    conference.columns = ["Conf_name", "start_date", "end_date", "location",
                          "conf_link", "accepted", "departure_date",
                          "return_date", "days", "transport", "hotel",
                          "meals", "registration", "other", "total"]
    conference['authors'] = z.loc[0, 4]
    last_name = z.loc[1,4].strip().split(" ")[1]
    conference['applicant'] = z.loc[1, 4]
    conference['applicant_email'] = z.loc[2, 4]
    conference['submission_date'] = z.loc[0, 13]
    conference['dept'] = z.loc[1, 13]
    conference['manager'] = z.loc[2, 13]
    conference['paper_title'] = z.loc[4,4]
    # Now create links to the papers
    paper = z.loc[5,4]
    if ":" not in paper:
        paper = '=HYPERLINK("O:/Talent Strategy - Supporting Files/'+quarter+"/"+last_name+"/"+paper+'","link")'
    else:
        paper = '=HYPERLINK("'+paper+'","link")'
    conference['paper_link'] = paper
    conference['times_approves'] = z.loc[6,16]
    conference['recent_quarter'] = z.loc[6,16]
    return(conference)

# Now find all the files and loop over them
app_data = []
for file in applications:
    app_data.append(read_conference_sheet(file,quarter))
app_data = pd.concat(app_data)
# Reset the index
app_data.sort_values(by=["dept","authors","paper_title","departure_date"],
                     inplace=True)
app_data.reset_index(drop=True,inplace=True)

dsa_data = []
for file in dsas:
    dsa_data.append(read_dsa_sheet(file,quarter))
dsa_data = pd.concat(dsa_data)
dsa_data.sort_values(by=['department','DSA_name'],inplace=True)
dsa_data.reset_index(drop=True,inplace=True)

# Now reorder the columns and rename them to a friendly version
# trips
app_data_order = ['dept', 'authors','paper_title', 'Conf_name',
                  'conf_link', 'location', 'departure_date', 'start_date', 
                  'end_date',  'return_date', 'days', 'accepted', 
                  'total', 'transport', 'hotel', 'meals',
                  'registration', 'other', 'applicant', 'applicant_email', 
                  'submission_date', 'manager', 'paper_link', 'times_approves',
                  'recent_quarter']
header_names = ['Dept', 'Authors','Paper', 'Conference', 
                'Conference Link','Location', 'Departure','Start', 
                'End',  'Return',  'Days', 'Accepted', 
                'Total', 'Transport', 'Hotel', 'Meals',
                  'Registration', 'Other', 'Applicant', 'applicant_email',
                  'submission_date', 'manager', 'paper_link', 'times_approves',
                  'recent_quarter']

app_data = app_data[app_data_order]
# dsas
dsa_data_order = ['DSA_name','applicant', 'department','total',
                  'visit_number', 'visit_start', 'visit_end', 'work_days',
                'working_transit', 'per_diem', 'total_honorarium', 'transport',
                'hotel', 'meal', 'hospitality', 'other']
dsa_data = dsa_data[dsa_data_order]

# Now start an XlsxWriter instance
with pd.ExcelWriter(save_file,engine="xlsxwriter") as writer:
    app_data.to_excel(writer,sheet_name="Research Trips",header=header_names)
    dsa_data.to_excel(writer,sheet_name="DSAs",header=dsa_data_order)



