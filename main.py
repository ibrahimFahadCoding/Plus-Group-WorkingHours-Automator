# This program calculates weekly employee compensation.
# Output: Generates an easy-to-read text file with the computed information
# Version: 2.1
# Authors: Ibrahim Fahad
# (C) Copyright 2022

# Import pandas for parsing sheet
import pandas as pd
from datetime import date
import os

# Input the directory path
directory = os.path.expanduser("~/Downloads/Work")

# Filter Excel files in the directory
excelFiles = [file for file in os.listdir(directory) if file.endswith(".xlsx") and not file.startswith("~$")]

print("Scanning Work Directory...")

# Proceed if there are Excel files found
if excelFiles:
    # Select the first Excel file for processing
    fileConvert = excelFiles[0]
    
    # Ask for confirmation
    isit = input("File found: " + fileConvert + ". Is this it (y/n): ")
    
    if isit.lower() == "y":
        print("Calculating " + fileConvert)

        # Read the Excel file
        df = pd.read_excel(os.path.join(directory, fileConvert))

        # Extract filename for result file
        filename = os.path.splitext(fileConvert)[0]

    else:
        fileConvert = input("Enter file: ")
        df = pd.read_excel(os.path.join(directory, fileConvert))

        # Extract filename for result file
        filename = os.path.splitext(fileConvert)[0]

    # Print the path for the result file
    print("\n-----------------------------------------------------------------")
    print("| Go to Downloads > Work > Results > " + filename + "_result.txt |")
    print("-----------------------------------------------------------------\n")

    # Filter the excel sheet
    df.drop(['Personnel number', 'Payroll Area', 'Vendor Key', 'Vendor Text', 'Sending PO',
             'Activity Type', 'For-period for payroll', 'In-period for payroll', 'FP Start date',
             'FP End date', 'IP Start date', 'IP End date', 'Date', 'Wage Type', 'Regular Amounts',
             'Retro Amounts', 'Total Hours', 'Total Amounts'], axis=1, inplace=True)

    # Get the number of samples
    numRows = len(df)

    # Create all necessary variables
    name = df.iloc[0, 0]
    regular = 0
    OT = 0
    retro = 0
    retroOT = 0

    # Create the file for computed data
    with open(os.path.join(directory, "Results", filename + "_result.txt"), "w") as f:
        # Logic to compute the weekly data
        today = date.today()
        d = today.strftime("%m/%d/%y")

        print("******************************************", file=f)
        print("Here are the results. Computed on " + str(d), file=f)
        print("******************************************", file=f)
        print("", file=f)

        for i in range(0, numRows):
            readName = df.iloc[i, 0]
            readWageType = df.iloc[i, 1]

            if readName != name:
                print("Name: " + str(name), file=f)
                print("Regular: " + str(regular), file=f)
                print("Overtime: " + str(round(OT, 2)), file=f)
                print("Retro: " + str(retro), file=f)
                print("Retro Overtime: " + str(retroOT), file=f)
                print("", file=f)
                print("", file=f)
                name = df.iloc[i, 0]
                regular = 0
                OT = 0
                retro = 0
                retroOT = 0

            if readWageType == "Regular":
                regular += df.iloc[i, 2]
                retro += df.iloc[i, 3]
            elif readWageType == "O.T STRAIGHT":
                OT += df.iloc[i, 2]
                retroOT += df.iloc[i, 3]

            if i == numRows - 1:
                print("Name: " + str(name), file=f)
                print("Regular: " + str(regular), file=f)
                print("Overtime: " + str(OT), file=f)
                print("Retro: " + str(retro), file=f)
                print("Retro Overtime: " + str(retroOT), file=f)
                # Close file

    print("Results have been calculated and saved.")

else:
    print("No '.xlsx' file was detected in the directory.")
