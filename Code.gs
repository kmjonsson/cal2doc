/*
Cal2Doc - Converts Google Calender Events into Tables in Google Docs.
Copyright (C) 2022  Magnus Jonsson

This program is free software; you can redistribute it and/or
modify it under the terms of the GNU General Public License
as published by the Free Software Foundation; either version 2
of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program; if not, write to the Free Software
Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.
*/

function sync(e) {
  let syncData = getSyncData();
  // Sync calenders
  for(let data of syncData) {
    if(!e || e['calendarId'] && e.calendarId == data.calendarId) {
      let rows = getEvents(data);
      updateTable(data,rows);
    }
  }
  // Don't update triggers on trigger runs
  if(!e) {
    updateTriggers(syncData);
  }
}

function onOpen() {
  let ui = SpreadsheetApp.getUi(); // Or DocumentApp or SlidesApp or FormApp.
  let menu = ui.createMenu('Cal2Doc');
  menu.addItem('Sync', 'sync');
  menu.addToUi();
}

function updateTriggers(data) {
  let allTriggers = ScriptApp.getProjectTriggers();

  // Add
  for(let d of data) {
    if(!d.auto) {
      continue;
    }
    let found = false;
    for (let trigger of allTriggers) {
      if(trigger.getTriggerSourceId() == d.calendarId) {
        found = true;
        break;
      }
    }
    if(!found) {
      Logger.log("Add trigger for: %s", d.calendarId);
      ScriptApp.newTrigger('sync')
        .forUserCalendar(d.calendarId)
        .onEventUpdated()
        .create();
      // Update
      allTriggers = ScriptApp.getProjectTriggers();
    }
  }

  // Remove unused..
  for (let trigger of allTriggers) {
    //Logger.log(trigger.getEventType());
    //Logger.log(trigger.getTriggerSource());
    if(trigger.getEventType() != "ON_EVENT_UPDATED") {
      continue;
    }
    if(trigger.getTriggerSource() != "CALENDAR") {
      continue;
    }
    let found = false;
    for(let d of data) {
      if(!d.auto) {
        continue;
      }
      if(trigger.getTriggerSourceId() == d.calendarId) {
        found = true;
        break;
      }
    }
    //Logger.log(trigger.getTriggerSourceId());
    if(found) {
      continue;
    }
    Logger.log("Remove trigger for: %s",trigger.getTriggerSourceId());
    ScriptApp.deleteTrigger(trigger);
  }

}

function getSyncData() {
  let ss = SpreadsheetApp.getActive();
  let sheets = ss.getSheets();
  let sheet = sheets[0];

  let syncData = []
  var rows = sheet.getDataRange().getValues();
  rows.shift();
  for(let row of rows) {
    if(row[0]) {
      syncData.push({
        "description": row[0],     // Just for debug, not used.
        "calendarId": row[1],      // calenderId
        "documentId": row[2],      // documentId
        "table": parseInt(row[3]), // Table number in document starting @ 0
        "fromDate": row[4],        // Start Date to extract from calendar
        "toDate": row[5],          // End Date to extract from calendar
        "start_time": row[6],      // Default start time of event
        "end_time": row[7],        // Default end time of event
        "table_content": row[8],   // What to insert into calendar
        'timeZone': (row[9] && row[9] != "") ? row[9] : "CET",
        'auto': ['yes','true','YES','ja'].includes(row[10]) ? true : false,
      })
    }
  }
  return syncData;
}

function updateTable(data,rows) {
  //Logger.log(data);
  var body = DocumentApp.openById(data.documentId).getBody();
  var tables = body.getTables()
  var table = tables[data.table]
  var tableRow = table.getRow(1);

  let len = rows.length;

  if(len <= 0) {
    Logger.log("Refusing to update 0 rows.");
    return;
  }

  // Remove extras
  if(table.getNumRows()-1 > len) {
    for(var i=table.getNumRows()-1;i>=len;i--) {
      table.removeRow(i);
    }
  }
  // Create missing
  if(table.getNumRows()-1 < len) {
    let n = len-table.getNumRows()+1;
    for(var i=0;i<n;i++) {
      table.appendTableRow(tableRow.copy());
    }
  }

  let rowKeys = data.table_content.split(",");
  if(rowKeys.length == 0 || rowKeys == "") {
    rowKeys = ["dateStr","title","location"];
  }

  let rp = 1;
  for(const [n, row] of rows.entries()) {
    let t = table.getRow(rp);
    for (const [i, key] of rowKeys.entries()) {
      if(key == 'skip') {
        continue;
      }
      if(key in row) {
        let ctext = t.getCell(i).getText();
        if(ctext != row[key]) {
          t.getCell(i).setText(row[key]);
        }
      } else {
        Logger.log("Can't update cell: %d on row: %d, no key: %s", i, n, key);
      }
    }
    rp++;
  }
}

/**
 * Fetch events
 */
function getEvents(data) {
  //Logger.log(data);
  var events = Calendar.Events.list(data.calendarId, {
    timeMin: data.fromDate.toISOString(),
    timeMax: data.toDate.toISOString(),
    singleEvents: true,
    orderBy: 'startTime',
  });
  var rows = [];
  if (events.items && events.items.length > 0) {
    for (var event of events.items) {
      let start;
      let start_time = "";
      let end;
      let end_time = "";

      let timeZone = data.timeZone;

      if (event.start.date) {
        // All-day event.
        start = new Date(Date.parse(event.start.date));
        end = new Date(Date.parse(event.end.date)-1000); // ends @ next day 00:00:00, sub -1s
        timeZone = "UTC"; // start, end is in UTC
      } else {
        start = new Date(event.start.dateTime);
        end = new Date(event.end.dateTime);
        // Remove seconds..
        start_time = start.toLocaleTimeString("sv-SE", {timeZone}).replace(/^(\d+:\d+):\d\d$/,'$1');
        end_time = end.toLocaleTimeString("sv-SE", {timeZone}).replace(/^(\d+:\d+):\d\d$/,'$1');
        // Remove if default start/end_time
        if(start_time == data.start_time && end_time == data.end_time) {
          start_time = "";
          end_time = "";
        }
      }
      // Convert to "D/M"-format.
      let start_date = start.toLocaleDateString("sv-SE", {timeZone}).replace(/(\d+)-0?(\d+)-0?(\d+)/, '$3/$2');
      let end_date = end.toLocaleDateString("sv-SE", {timeZone}).replace(/(\d+)-0?(\d+)-0?(\d+)/, '$3/$2');
      let dateStr = start_date;
      if(start_date != end_date) { // add end_date if start_date != end_date
        dateStr = start_date + " " + start_time + " - " + end_date + " " + end_time;
      } else if(start_time != "") {
        dateStr = start_date + " " + start_time + " - " + end_time;
      }
      // Convert to ISO-format
      let iso_start_date = start.toLocaleDateString("sv-SE", {timeZone});
      let iso_end_date = end.toLocaleDateString("sv-SE", {timeZone});
      let dateISOStr = iso_start_date;
      if(iso_start_date != iso_end_date) { // add iso_end_date if iso_start_date != iso_end_date
        dateISOStr = iso_start_date + " " + start_time + " - " + iso_end_date + " " + end_time;
      } else if(start_time != "") {
        dateISOStr = iso_start_date + " " + start_time + " - " + end_time;
      }

      rows.push({
        'dateStr': dateStr.replace("  "," "),
        'dateISOStr': dateISOStr.replace("  "," "),
        'title': event.summary,
        'location': (event.location ? event.location : "").split(",")[0],
        'longLocation': event.location,
        'description': event.description ? event.description : "",
        'empty': "",
      });
    }
  } else {
    Logger.log('No events found.');
  }
  return rows;
}
