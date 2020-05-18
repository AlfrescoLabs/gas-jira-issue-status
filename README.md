# Google Apps Script JIRA Issue Status Tracking

This project uses JIRA's REST API to periodically pull data on the status
of collections of tickets identified via JQL queries.

Google Sheets native functions can then be used to process and summarise
the status data trends over time, and charts can be published to Google
Drive or Confluence for visibility across the organisation.

## Quick Setup

Create a new Google Sheet in which you will pull down the data, and add
the following headings in row 1, columns A-C:

    Issue Type | Issue Key | Summary

In cell D1, add in the first date that you want to start tracking,
usually this will be today's date.

Open up the code editor by clicking *Tools* > *Script editor* and make a
note of the project ID in the URL.

Install [Google Clasp](https://github.com/google/clasp) if you do not
have it installed and use `clasp login` to authenticate the tool.

Create a script project associated with your new spreadsheet by passing
the ID of the spreadsheet that you created in Step 1.

    clasp create --parentId "1D_Gxyv*****************************NXO7o" --rootDir src
    clasp push

Reload the spreadsheet in the browser and use the JIRA Issue Tracking
menu to access the functionity of the project.
