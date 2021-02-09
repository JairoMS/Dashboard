var meeting_calendar_id='c_k8u5g1urpvnqvh0tpovu43cr3k@group.calendar.google.com';

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
  

  //会議室のカレンダーに登録してアクセスする / Subscribe to the meeting room calendar and access it
  var calendars = CalendarApp.getCalendarsByName(CalendarApp.subscribeToCalendar(meeting_calendar_id).getName());
  var start_time = new Date(bookingInfo.date + " " + bookingInfo.startTime);
  var end_time = new Date(bookingInfo.date + " " + bookingInfo.finishTime);
  const eventsToday = calendars[0].createEvent(subCal.getName() +" (" + bookingInfo.name + ")", new Date(start_time.getTime()-1000 * 60 * 60 * 14), new Date(end_time.getTime()-1000 * 60 * 60 * 14));//イベントを作成する / Create Event
  sheet.getRange(cell_bookingId).setValue(eventsToday.getId());
  calendars[0].unsubscribeFromCalendar() //会議室のカレンダーの登録を削除する / Unsubscribe from meeting room calendar
  
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

    //会議室のカレンダーに登録してアクセスする / Subscribe to the meeting room calendar and access it
  var meeting_calendar=CalendarApp.getCalendarsByName(CalendarApp.subscribeToCalendar(meeting_calendar_id).getName());
  var event = meeting_calendar[0].getEventById(event_id);
  event.deleteEvent();
  //会議室のカレンダーの登録を削除する / Unsubscribe from meeting room calendar
  meeting_calendar[0].unsubscribeFromCalendar();
  Logger.log(event);
}

function data_from_ss()
{
  var url_ss = "https://docs.google.com/spreadsheets/d/1Om1kYwsVAISmAS8LnI8S2_INkpf0Q33-35GLhbY_jp0/edit#gid=0";
    var ss = SpreadsheetApp.openByUrl(url_ss);
    var sheet = ss.getSheetByName("Data");
    
    r = sheet.getLastRow();
    c = sheet.getLastColumn();
    return sheet.getSheetValues(2, 2, r-1, 4);
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
    var cell_id = "E"+find_name(list, calendar[i].getTitle())
    sheet.getRange(cell_id).setValue(calendar_name + " " + st + " - " + et);
    
  }
}

function read_calendar2(){
  var url_ss = "https://docs.google.com/spreadsheets/d/1Om1kYwsVAISmAS8LnI8S2_INkpf0Q33-35GLhbY_jp0/edit#gid=0";
  var ss = SpreadsheetApp.openByUrl(url_ss);
  var sheet = ss.getSheetByName("Data");
  var calendar_names=sheet.getRange("H2:H6").getValues();
  for(var i=0;i<calendar_names.length;i++){
    if (calendar_names[i].toString()!==""){
          var calendar_name = calendar_names[i].toString();
          var today = new Date();
          var calendar  =  Calendar.Calendars.get(calendar_name)
          if(typeof calendar !== 'undefined'){
            var aCal=CalendarApp.subscribeToCalendar(calendar.id)
            var eventToday=CalendarApp.getCalendarsByName(aCal.getName())[0].getEventsForDay(today);
            if(eventToday.length>0){
                eventToday.forEach(theEvent=>{
                  Logger.log(theEvent.getTitle())
                  var st = Utilities.formatDate(theEvent.getStartTime(), "GMT+9", "HH:mm");
                  var et = Utilities.formatDate(theEvent.getEndTime(), "GMT+9", "HH:mm");
                  var cell_id = "E"+(i+2)
                  sheet.getRange(cell_id).setValue(aCal.getName() + " " + st + " - " + et);
                })
            }
              
          }else{

          }
          
    }
    
    

  }
  
  /*var ownedCal=CalendarApp.getAllCalendars();
  var today=new Date();
  for(var i=0;i<ownedCal.length;i++){
    var calId=ownedCal[i].getName()
    Logger.log(CalendarApp.getCalendarsByName(calId)[0].getEventsForDay(today))
  }*/
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
// }

