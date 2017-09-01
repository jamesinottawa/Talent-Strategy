import pandas as pd

# These are the cumulative files to merge
master_file_2016Q3 = "O:/Talent Strategy - To process/Master List - cumulative as of 2016Q4.xlsm"
master_file_2016Q4 = "O:/Talent Strategy - To process/Master List - 2016Q4.xlsm"
master_file_2017Q1 = "O:/Talent Strategy - To process/Master List - 2017Q1.xlsm"
# This is the file to save it too
save_file_to_2017Q1 = "O:/Talent Strategy - To process/Master List - cumulative as of 2017Q2.xlsx"
#These are the header names
header_names = ["authors","dept", "paper_title",'times_approves', 'times_accepted','Conf_name', 'location',
'start_date', 'end_date','departure_date',  'return_date','total',"approved","forecast","comments",
                  "accepted","recent_quarter","approved","comments_to_applicant",'applicant_email',
'manager',"emailed_status","emailed_decision"]

master_2016Q3 = pd.read_excel(master_file_2016Q3,header=None, skiprows=3)
master_2016Q3 = master_2016Q3.loc[:,:22]
master_2016Q3.columns = header_names
master_2016Q4 = pd.read_excel(master_file_2016Q4,header=None, names=header_names,skiprows=3)
master_2017Q1 = pd.read_excel(master_file_2017Q1,header=None, names=header_names,skiprows=3)

master = pd.concat([master_2016Q3,master_2016Q4,master_2017Q1])
master[["authors","dept","paper_title"]] = master[["authors","dept","paper_title"]].fillna(method="pad")
totals = master["authors"].str.lower().str.contains("total")
master = master[~totals]

# Now start an XlsxWriter instance
with pd.ExcelWriter(save_file_to_2017Q1,engine="xlsxwriter") as writer:
    master.to_excel(writer,sheet_name="Research Trips",header=header_names)


