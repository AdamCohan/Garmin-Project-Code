# Garmin-Project-Code
 
Project for Dr. Dustin Joubert meant to help automate the summary of metrics from wearables and package it into an Excel workbook.
I've been working on it for about two weeks so it's gone thorugh some iterations before this but better late than never to make a github repo.

TO USE:
Install pandas and garmin-fit-sdk
Call the following from the command line:
python3 fit2trialsummary.py [.fit filepath] [minute timestamps for the end of the intervals for data summary]
ex: python3 fit2trialsummary.py filename.fit 10 20 30 40 50

OUTPUT:
Writes an excel workbook with the same name (but .xlsx extension (duh)) as the .fit file with three sheets. The first two represent the summary of the relevant metrics over the last 2min and 5min of each interval respectively. The third sheet represents the continuous raw output of data from the .fit file for debugging purposes.

NOTE:
Much of this can be made more modular, such as the metrics which are included in the summary and the 2min/5min interval length in the summary. These are not modular (yet) because this project is made for a specific purpose and at the moment it's okay that they're 'magic numbers'.