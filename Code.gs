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
    if(!e || e['calendarId'] && data.calendarId.includes(e.calendarId)) {
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
  let allCalenders = new Set()

  for(let d of data) {
    if(d.auto) {
      for(let c of d.calendarId) {
        allCalenders.add(c);
      }
    }
  }

  // Add
  for(let c of allCalenders) {
    let found = false;
    for (let trigger of allTriggers) {
      if(trigger.getTriggerSourceId() == c) {
        found = true;
        break;
      }
    }
    if(!found) {
      Logger.log("Add trigger for: %s", c);
      ScriptApp.newTrigger('sync')
        .forUserCalendar(c)
        .onEventUpdated()
        .create();
      // Update
      allTriggers = ScriptApp.getProjectTriggers();
    }
  }

  // Remove unused..
  for (let trigger of allTriggers) {
    if(trigger.getEventType() != "ON_EVENT_UPDATED") {
      continue;
    }
    if(trigger.getTriggerSource() != "CALENDAR") {
      continue;
    }
    let found = false;
    for(let c of allCalenders) {
      if(trigger.getTriggerSourceId() == c) {
        found = true;
        break;
      }
    }
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
    if(row[0] && row[1] && row[2] && row[4] && row[5] && row[6] && row[7] && row[8]) {
      syncData.push({
        "description": row[0],     // Just for debug, not used.
        "calendarId": row[1].split(","),      // calenderId
        "documentId": row[2],      // documentId
        "table": row[3] ? parseInt(row[3]) : 0, // Table number in document starting @ 0
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
      let ctext = t.getCell(i).getText();
      let ntext = "";
      if(key in row) {
        ntext = row[key];
      } else {
        Logger.log("Can't update cell: %d on row: %d, no key: %s", i, n, key);
        Logger.log(row);
      }
      if(ctext != ntext) {
        t.getCell(i).setText(ntext);
      }
    }
    rp++;
  }
}

/**
 * Sort function for sorting events
 */
function compareEvent(a,b) {
  if(a.start.date) {
    if(b.start.date) {
      return new Date(a.start.date) - new Date(b.start.date);
    } else {
      return new Date(a.start.date) - new Date(b.start.dateTime);
    }
  } else {
    if(b.start.date) {
      return new Date(a.start.dateTime) - new Date(b.start.date);
    } else {
      return new Date(a.start.dateTime) - new Date(b.start.dateTime);
    }
  }
  return 0;
}

/**
 * Fetch events
 */
function getEvents(data) {
  var events = []
  for(id of data.calendarId) {
    var e = Calendar.Events.list(id, {
        timeMin: data.fromDate.toISOString(),
        timeMax: data.toDate.toISOString(),
        singleEvents: true,
        orderBy: 'startTime',
      });
    events = [...events,...e.items];
  }
  events.sort(compareEvent);
  var rows = [];
  if (events && events.length > 0) {
    for (var event of events) {
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

      let rdata = {
        'dateStr': dateStr.replace("  "," "),
        'dateISOStr': dateISOStr.replace("  "," "),
        'title': event.summary,
        'location': (event.location ? event.location : "").split(",")[0],
        'longLocation': event.location,
        'description': event.description ? event.description.replace(/<br>|<\/li>/g, "\n").replace(/<.+?>/g, "").trim() : "",
        'empty': "",
      };
      let i = 0;
      for(var line of rdata.description.split("\n")) {
        line = line.replace("\n","");
        rdata['description[' + i++ + ']'] = line
      }
      rows.push(rdata);
    }
  } else {
    Logger.log('No events found.');
  }
  return rows;
}
