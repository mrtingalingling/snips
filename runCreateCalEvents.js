function runCreateCalEvents() {
  // identifyCal
  var spreadsheet = SpreadsheetApp.getActiveSheet();
  // var calendarId = spreadsheet.getRange("B6").getValue();
  // var eventCal = CalendarApp.getCalendarById(calendarId);
  const mainCal = CalendarApp.getDefaultCalendar();
  // Logger.log(cal)

  // getEventInfo
  // var eventInfo = spreadsheet.getRange("D:J").getValues();
  let eventInfo = spreadsheet.getRange(2, 4, spreadsheet.getLastRow(), 10).getValues();
  let eventName, day, startDate, startDate_startTime, startDate_endTime, endDate, startTime, endTime, roomLoc;
  // Logger.log('Number of Events: ' + eventInfo.length);

  for (let i in eventInfo){
    // for (let j in eventInfo[i]){
    //   Logger.log(j + ' - ' + eventInfo[i][j]);
    // }
    // startTime = new Date(eventInfo[i][1]);
    // endTime = new Date(eventInfo[i][2]);
    // startDate = new Date(eventInfo[i][3]);
    endDate = new Date(eventInfo[i][4]);
    endDate = new Date(endDate.getTime() + (24 * 60 * 60 * 1000)); // Add a day to the endDate
    startDate_startTime = new Date(eventInfo[i][8]);
    startDate_endTime = new Date(eventInfo[i][9]);
    roomLoc = eventInfo[i][5];
    eventName = eventInfo[i][7];
    // Logger.log('startDate_startTime: ' + startDate_startTime);
    // Logger.log('startDate_endTime: ' + startDate_endTime);
    // Logger.log('endDate: ' + endDate);
    // Logger.log('EventName: ' + eventName);

    // startDate_startTime = startDate.setTime(startTime.getHours, startTime.getMinutes);
    // startDate_endTime = startDate.setTime(endTime.getHours,endTime.getMinutes);

    switch (eventInfo[i][0]) {
      case "Su":
        day = CalendarApp.Weekday.SUNDAY;
        break;
      case "Mo":
        day = CalendarApp.Weekday.MONDAY;
        break;
      case "Tu":
        day = CalendarApp.Weekday.TUESDAY;
        break;
      case "We":
        day = CalendarApp.Weekday.WEDNESDAY;
        break;
      case "Th":
        day = CalendarApp.Weekday.THURSDAY;
        break;
      case "Fr":
        day = CalendarApp.Weekday.FRIDAY;
        break;
      case "Sa":
        day = CalendarApp.Weekday.SATURDAY;
    }

    // checkCalEvent
    var roomLocCal = calMap(roomLoc) || mainCal;
    // Logger.log('Calendar: ' + roomLocCal.getName() + ' - ' + roomLoc);
    if (eventName) {
      let eventCheckLocal = checkCalEvent(roomLocCal, eventName, startDate_startTime, endDate)
      let eventCheckMain = checkCalEvent(mainCal, eventName, startDate_startTime, endDate)
      if (eventCheckLocal.length == 0 && eventCheckMain.length == 0) {
        // createCalendarEvents
        createCalendarEvent(mainCal, eventName, startDate_startTime, startDate_endTime, endDate, day, roomLoc);
        Logger.log(eventName + ' Created!');
      }
    }
  }
}

function deleteEvents(cal, eventName, startDate, endDate) {
  let events = checkCalEvent(cal, eventName, startDate, endDate);
  for (let k in events) {
    events[k].deleteEvent();
    Logger.log(eventName + ' Deleted!');
  }
}

function createCalendarEvent(cal, eventName, startDate_startTime, startDate_endTime, endDate, day, roomLoc) {
    var eventSeries = cal.createEventSeries(eventName, startDate_startTime, startDate_endTime, 
        CalendarApp.newRecurrence().addWeeklyRule()
            .onlyOnWeekdays([day])
            .until(endDate),
        {location: roomLoc}
    );
    Logger.log('Event Series ID: ' + eventSeries.getId());
}

function checkCalEvent(cal, eventName, startDate, endDate) {
  // Determines how many events are happening in the next two hours that contain eventName
  var events = cal.getEvents(startDate, endDate, {search: eventName});
  Logger.log('Number of matching ' + eventName + ' occurances: ' + events.length);
  return events
}

function calMap(roomLoc) {
  // Get Calendar ID from matching room location
  let calendarId; 

  switch (true) {
    case roomLoc.includes("____"):
      calendarId = "____";
      break;
  }
  let localCal = CalendarApp.getCalendarById(calendarId); 

  return localCal
}

// https://workspace.google.com/blog/productivity-collaboration/g-suite-pro-tip-how-to-automatically-add-a-schedule-from-google-sheets-into-calendar
// https://developers.google.com/apps-script/reference/calendar/calendar-app
// https://www.youtube.com/watch?v=Z0fWGKausWA
