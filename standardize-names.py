# This file loops through a given directory
# loads each file and then renames the file
# according to the last name of the applicant
# and the quarter.
#
# Author : James T.E. Chapman
# Creation Date : 29-05-2017

import os, re, os.path
from glob import glob
import pandas as pd

# this is the directory to look in
application_directory = "O:/Talent Strategy - To process/Conference Applications"
quarter = "2017Q2"
# grab the files to be renamed
files = glob(application_directory+"/"+quarter+"/*.xlsm")

for file in files:
    z = pd.read_excel(file, sheetname="questionnaire", skiprows=7, header=None)
    author = z.ix[1, 4].strip().split(" ")[1]
    new_name_1 = "TS_" + quarter + "_" + author
    full_name = "{0}/{1}/{2}.xlsm".format(application_directory, quarter, new_name_1)
    # Check to see if this exists
    tmp = 1
    while os.path.isfile(full_name):
        new_name = new_name_1 + "_" + str(tmp)
        tmp += 1
        full_name = "{0}/{1}/{2}.xlsm".format(application_directory, quarter, new_name)
    # Now rename the file
    os.rename(file,full_name)

print("Done")



