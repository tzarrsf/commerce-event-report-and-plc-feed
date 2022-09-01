# Commerce Event Report and PLC Feed (initially for events without completion / grading aspects)

> Currently this template supports "Series-type" events such as Office Hours, How to Sell, Coming Soon, etc. where completion and grading aspects are not needed. There are plans to expand this functionality to support "Workshop-type" events with completion and grading aspects.

## Configuration:

Run as TZ, Notify Immediately on Error, onFormSubmit throughout

## Relevant Quips
None at this time

## Relevant Forms
None at this time

## Relevant Sheet / Tab
https://docs.google.com/spreadsheets/d/1tZFxXg3Uu6atsbHc-4VZBDnVmE4ASODn2MPRJt3kDGI/edit#gid=443946138

## Deployment

Code.gs is loaded into Extensions > Apps Script (from the top menu of the sheet). Presence of the code will create a Commerce Team Tools menu with two options:
1. Generate 'Registered' report for PLC
2. Generate 'Attended' report for PLC

# How to use
1. This is a template in the FY22 event tracking sheet (Lena's dashboard numbers based on our event promotion and hosting).
2. Follow the copy pasta instructions in the sheet using your zoom webinar attendance report as the only input.
3. Regions, email domains and other aspects should resolve to much cleaner data on the right.
4. The summary report at the top should fill in across SF, Partners and broken out by region.
5. Use the *Commerce Team Tools* menu to generate your Registered and Attended data and you can upload it to the PLC for your event. 
6. @-mention to Lena that your new numbers are ready and have been uploaded to PLC.

# TODOs:
1. Expand functionality to allow for workshop-type events which involve completion / grading aspects.
