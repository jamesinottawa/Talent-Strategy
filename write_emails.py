# This file takes the decision file from CC and creates a directory full of "eml" files to send to
# applicants
# Author : James Chapman
# Creation date : 2017-09-13

# load the needed libraries
import pandas as pd

# this is the CC directory where the decision file should be kept and the file name
cc_dir = "C:/Users/chja/Desktop"
file_name = "Decisions TS 2017Q3.xlsx"
# THi sis the directory where eml files will be saved
save_dir = "C:/Users/chja/Desktop" #TODO fix save directory

# This part is the peices of the email
email_start = "Dear "
email_txt_intro = (',\n'
                   '\n'
                   'Thank you for your application to the TS budget. The committee has carefully reviewed your application(s) and come to \n'
                   'the following decisions highlighted below.\n'
                   '\n'
                   '\n')

email_txt_end = ("\n"
                 "\n"
                 "Good luck in your research at the Bank,\n"
                 "\n"
                 "Talent Strategy Committee")

# load the decision file
ts_decisions =pd.read_excel(cc_dir + "/" + file_name)

# now find the unique emails to begin the loop
emails = ts_decisions.applicant_email.unique()
drafts = {}


for n in range(len(emails)):
    email = emails[n]
    applicant_subset = ts_decisions.loc[ts_decisions.applicant_email == email,]
    applicant_name = applicant_subset.Applicant.unique()
    applicant_manager = applicant_subset.manager.unique()
    drafts[email] = 10*"-" + "\nTo: " + email + "\nCC : " + applicant_manager[0] + "\nSubject: Talent Strategy Decision\n\n"
    tmp_draft = email_start + applicant_name[0] + email_txt_intro
#     Now loop over papers
    projects = applicant_subset.Paper.unique()
    for m in range(len(projects)):
        project = applicant_subset.loc[applicant_subset.Paper == projects[m]]
        tmp_draft += "\nfor the paper '" + projects[m] + "':\n"
        for l in range(project.shape[0]) :
            conference = project.iloc[l,]
            approved = "not approved"
            if conference.Approve == "Yes" :
                approved = "approved"
            conference_line = "\t" + conference.Conference + " in " + conference.Location + " travelling on " + conference.Start.strftime('%Y-%m-%d') + " is " + approved + "\n"
            tmp_draft += conference_line
    tmp_draft += email_txt_end + "\n" + 10*"-"
    drafts[email] += tmp_draft

for v in drafts:
    print(drafts[v])









