function doGet(e) 
{
  read_calendar();
  read_calendar2();
  return HtmlService.createTemplateFromFile("main").evaluate();  
}

function include(file_name)
{
  return HtmlService.createHtmlOutputFromFile(file_name).getContent();
}

function saveUser(userInfo)
{
    var url_ss = "https://docs.google.com/spreadsheets/d/1Om1kYwsVAISmAS8LnI8S2_INkpf0Q33-35GLhbY_jp0/edit#gid=0";
    var ss = SpreadsheetApp.openByUrl(url_ss);
    var sheet = ss.getSheetByName("Data");
    var cell_status = "C" + (parseInt(userInfo.row)+1).toString();
    var cell_note = "D" + (parseInt(userInfo.row)+1).toString();
    sheet.getRange(cell_status).setValue(userInfo.status);
    sheet.getRange(cell_note).setValue(userInfo.note);
}

function saveBooking(bookingInfo)
{ 
  var url_ss = "https://docs.google.com/spreadsheets/d/1Om1kYwsVAISmAS8LnI8S2_INkpf0Q33-35GLhbY_jp0/edit#gid=0";
  var ss = SpreadsheetApp.openByUrl(url_ss);
  var sheet = ss.getSheetByName("Data");
  var cell_booking = "F" + (parseInt(bookingInfo.row)+1).toString();
  var cell_bookingId = "G" + (parseInt(bookingInfo.row)+1).toString();
  sheet.getRange(cell_booking).setValue(bookingInfo.date + " " + bookingInfo.startTime + " - " + bookingInfo.finishTime);
  
  var calendar_name = '会議室の予約';
  //var calendar_name = 'アルバイト';
  var calendars=CalendarApp.subscribeToCalendar(calendar_name);
  //var calendars = CalendarApp.getCalendarsByName(calendar_name);
  var start_time = new Date(bookingInfo.date + " " + bookingInfo.startTime);
  var end_time = new Date(bookingInfo.date + " " + bookingInfo.finishTime);
  Logger.log('right here')
  const eventsToday = calendars[0].createEvent(calendar_name +" (" + bookingInfo.name + ")", new Date(start_time.getTime()-1000 * 60 * 60 * 14), new Date(end_time.getTime()-1000 * 60 * 60 * 14));
  sheet.getRange(cell_bookingId).setValue(eventsToday.getId());
  calendars.unsubscribeFromCalendar()
  
}

function deleteBooking(row)
{
  var url_ss = "https://docs.google.com/spreadsheets/d/1Om1kYwsVAISmAS8LnI8S2_INkpf0Q33-35GLhbY_jp0/edit#gid=0";
  var ss = SpreadsheetApp.openByUrl(url_ss);
  var sheet = ss.getSheetByName("Data");
  var cell_booking = "F" + (parseInt(row)+1).toString();
  var cell_bookingId = "G" + (parseInt(row)+1).toString();
  var event_id = sheet.getRange(cell_bookingId).getValue();
  sheet.getRange(cell_booking).setValue("");
  sheet.getRange(cell_bookingId).setValue("");

  //var calendar_name = '会議室の予約';
  var calendar_name = 'アルバイト';
  var event = CalendarApp.getCalendarsByName(calendar_name)[0].getEventById(event_id);
  event.deleteEvent();
  Logger.log(event);
}

function data_from_ss()
{
  var url_ss = "https://docs.google.com/spreadsheets/d/1Om1kYwsVAISmAS8LnI8S2_INkpf0Q33-35GLhbY_jp0/edit#gid=0";
    var ss = SpreadsheetApp.openByUrl(url_ss);
    var sheet = ss.getSheetByName("Data");
    
    r = sheet.getLastRow();
    c = sheet.getLastColumn();
    
    datass = sheet.getSheetValues(2, 2, r-1, 4);
    for (var i = 0 ; i<datass.length ; i++)
    {
      if (datass[i][3]!=="")
      {
        datass[i][3] = datass[i][3].replace(/\n/g,"<br>"); // /g all matches
      } 
    }
    return datass;
}

function read_calendar()
{
  var calendar_name = 'アルバイト';  
  var today = new Date();
  var calendar = CalendarApp.getCalendarsByName(calendar_name)[0].getEventsForDay(today);
  
  if (calendar.length == 0)
  {
    return;
  }

  var url_ss = "https://docs.google.com/spreadsheets/d/1Om1kYwsVAISmAS8LnI8S2_INkpf0Q33-35GLhbY_jp0/edit#gid=0";
  var ss = SpreadsheetApp.openByUrl(url_ss);
  var sheet = ss.getSheetByName("Data");

  list = sheet.getSheetValues(2, 2, sheet.getLastRow()-1, 1);

  for (var i = 0 ; i < calendar.length ; i++ )
  {
    var st = Utilities.formatDate(calendar[i].getStartTime(), "GMT+9", "HH:mm");
    var et = Utilities.formatDate(calendar[i].getEndTime(), "GMT+9", "HH:mm");
    var cell_id = "E"+find_name(list, calendar[i].getTitle());
    sheet.getRange(cell_id).setValue(calendar_name + " " + st + " - " + et);
    
  }
}

function read_calendar2()
{
  var url_ss = "https://docs.google.com/spreadsheets/d/1Om1kYwsVAISmAS8LnI8S2_INkpf0Q33-35GLhbY_jp0/edit#gid=0";
  var ss = SpreadsheetApp.openByUrl(url_ss);
  var sheet = ss.getSheetByName("Data");
  var calendar_names=sheet.getRange(2,8,sheet.getLastRow()-1).getValues();
  var today = new Date();
  var email = Session.getActiveUser().getEmail();

  //Logger.log(email)
  for(var i=0 ; i<calendar_names.length ; i++)
  {
    if (calendar_names[i].toString()!=="")
    {
      var calendar_name = calendar_names[i].toString();
      
      var calendar  =  Calendar.Calendars.get(calendar_name);
      if(typeof calendar !== 'undefined')
      {
        if (calendar_name !== email)
        {
          var aCal=CalendarApp.subscribeToCalendar(calendar.id);
          var eventToday=CalendarApp.getCalendarsByName(aCal.getName())[0].getEventsForDay(today);
        }
        else
        {
          var eventToday = CalendarApp.getCalendarById(calendar.id).getEventsForDay(today);
        }

        if(eventToday.length>0)
        { 
          var cell_id = "E"+(i+2);
          var string_events = sheet.getRange(cell_id).getValues()+"\n";  
          for (var j = 0 ; j<eventToday.length ; j++)
          {
            // Logger.log(eventToday.length);
            var theEvent = eventToday[j];
            var st = Utilities.formatDate(theEvent.getStartTime(), "GMT+9", "HH:mm");
            var et = Utilities.formatDate(theEvent.getEndTime(), "GMT+9", "HH:mm");
            
            string_events = string_events + theEvent.getTitle() + " " + st + " - " + et + "\n";
            // Logger.log(string_events);

          }
          sheet.getRange(cell_id).setValue(string_events); 
        }
        if (calendar_name !== email)
        {
          aCal.unsubscribeFromCalendar();
        }
          
      }
      else
      {
      }     
    }
  }
}  
  
function find_name(list, name)
{
  for (var i = 0 ; i < list.length ; i++)
  {
    if (list[i] == name)
    {
      return (i+2).toString(); // it adds 2 to match with the cell
    }
  }
}

// function test()
// {   
//   var email = Session.getActiveUser().getEmail();
//   Logger.log(email);  
// }

