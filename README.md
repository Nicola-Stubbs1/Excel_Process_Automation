## Excel Process Automation
❗Note: this repository only contains dummy data and has no reflection on real life NHS providers performance.

This repo is designed to create a sample excel file using the openxlsx package.

Description
The repo consists of:
Inputs folder - Includes 12 csv's - one per month and 2 xlsx files. This data is dummy aggregate data based on NHS England's statistic publication extracts.
* This data is fictional
Outputs folder - This is a folder to save the final excel output to. 
src folder - This file contains functions needed in the process.
Create_report.R file - This is the main script, everything is self contained and will import a file based on current date, then transforms the data into  summary and adds data and formatting to an excel object.
Which is finally written out to the Outputs folder.

Getting Started
Clone the Repo and connect via IDE.
This repo was built in R Studio but can also run in VS Code. 
openxlsx is also availble in Databricks - however This hasn't been tested to check if this repo can run with Databricks.
In R Studio - create new Project > Git > add in URL to this repo, choose a folder path to save repo to.
This will download the repo ready to use in R Studio.

Run the Create_report.R script
Please note - the dates within the csv files aren't reflective of the current date.
I created a function to extract the month from a file name, which doesn't always correlate the current year to the year within the file.
I just created the function to demonstrate some dynamic elements which can be added to the script to automate processes.
