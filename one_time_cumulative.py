# This file is to be run once to strip out all the old cumulative reports for data
# Author : James T.E. Chapman
# Creation Date : 29-05-2017

import os
import pandas as pd

# This is the file to load
old_cumulative_file = "O:/Talent Strategy - To process/Master List - cumulative as of 2016Q4.xlsm"
save_file = "C:/Users/chja/Desktop/temp_cum_ts.xlsx"

total_file = pd.read_excel(old_cumulative_file,header=None,skiprows=2)
total_file.columns = ["applicants","dept","paper","quarterl_approve","total_times","conference","location",
                      "start","end","departure","return","cost","approved","forecast","comments","accepted",
                      "lastQ","approved","comments","email","manager","status","decision","T","T1"]

# fill in the missing entries by a rolling fill
total_file.total_times = total_file.total_times.fillna(0)
total_file.accepted = total_file.total_times.fillna(0)
total_file.ix[:,0:3] = total_file.ix[:,0:3].fillna(method="ffill")
# drop the entries that start with
totals = total_file.applicants.str.lower().str.contains("total")
cleaned_file = total_file[~totals]

# Now start an XlsxWriter instance
writer = pd.ExcelWriter(save_file,engine="xlsxwriter")
cleaned_file.to_excel(writer,sheet_name="applications")
writer.save()

