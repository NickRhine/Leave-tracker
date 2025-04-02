# Rhine Leave Tracker

This program provides the ability to automatically track and update total number of leave days for each employee and leave requests made by employees.

## How it works

Users are registered by being added to the rhinemechatronics.com group on Microsoft Azure Entra. Admin are made through the same method: if you have an admin role on entra then you will have the admin role on the Leave Tracker.

The database for the website is an excel file hosted on the Rhine sharepoint. This spreadsheet is updated daily through the use of a Python script paired with github actions. (Temporary) As of the moment users will have to be added manually to the excel sheet to start their tracking, this may change in the future. Other than the addition of new users, all operations are able to be performed through the webpage.

## File Structure

More to be added closer to project completetion.
