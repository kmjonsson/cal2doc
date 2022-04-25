# Google Calendar to Google Document Table

Converts a set of google calendar events into google document table rows.

NOTE! The script will modify the document.

# Setup

## Create a Google Sheet.

Only the first sheet is used.

First row is ignored as a header.

Insert the following columns and make them look good: Description, Calendar, Google Doc, Table #, From, To, Start time, End Time, Table Content, Time Zone, Auto

## Fill in the information

### Description

This is just a comment to easy know what the sync is for, it is never used.

### Calendar

This is tha calendar "ID" found in the calendar settings. It has the form of an email adress.

### Google Doc

This is the "ID" of the document found in the URI of an open document.

### Table #

This is the reference to to table in the document. The first table has "0" and so on.

### From

Start Date from where to fetch events.

### To

End Date from where to fetch events.

### Start Time

If start time of an event is the same as "Start Time" it will not show in the table.

### End Time

If end time of an event is the same as "End Time" it will not show in the table.

### Table Content

A "," separated lite of content from the event to insert into the table.

Available data:

- dateStr
- dateISOStr (ISO format date)
- title
- location (First part until ",")
- longLocation
- description
- empty ("")
- skip (Do not edit)

Default: dateStr,title,location

### Time Zone

Time Zone code use. 

Default: CET

### Auto

If set to "yes" or "true" a trigger will be created that runs the sync on changes in the calender.

## Add App Script

Press the "Add on" meny and then "Apps Script".

Copy the code from "Code.gs" into the editor.

## Add Services

Add Calendar and DocumentApp service to the Apps Script.


# Run sync

Do the sync...
