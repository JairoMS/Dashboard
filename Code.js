var meeting_calendar_id='c_k8u5g1urpvnqvh0tpovu43cr3k@group.calendar.google.com';
var arubaito_calendar_id= 'silk.co.jp_p6079i696o6uajq49ijl3mgk94@group.calendar.google.com'
var reserveDate='nada';

function doGet(e) 
{
  // read_calendar();
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

function testing(){
  console.log("ff")
}

function getUserEmails(){
  var url_ss = "https://docs.google.com/spreadsheets/d/1Om1kYwsVAISmAS8LnI8S2_INkpf0Q33-35GLhbY_jp0/edit#gid=0";
  var ss = SpreadsheetApp.openByUrl(url_ss);
  var sheet = ss.getSheetByName("Data");
  var list_email = sheet.getRange(2,8,sheet.getLastRow()-1).getValues();
  return list_email;
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
  start_time = new Date(start_time.getTime()-1000 * 60 * 60 * 14);
  end_time = new Date(end_time.getTime()-1000 * 60 * 60 * 14);

  var eventsOnThatDay= calendars[0].getEventsForDay(start_time);
  Logger.log(eventsOnThatDay)

  // Check if there is an event at that time
  for (var i = 0 ; i<eventsOnThatDay.length ; i++)
  {
    s1 = eventsOnThatDay[i].getStartTime();
    e1 = eventsOnThatDay[i].getEndTime();

    if ( ( (start_time > s1 && start_time < e1) || (end_time > s1 && end_time < e1) ) || ((s1 > start_time && s1 < end_time) && (e1 > start_time && e1 < end_time)))
    { 
      Logger.log('Event exists!')
      return false;
    }
  }

  const eventsToday = calendars[0].createEvent(calendars[0].getName() +" (" + bookingInfo.name + ")", start_time, end_time);//イベントを作�Eする / Create Event
  sheet.getRange(cell_bookingId).setValue(eventsToday.getId());
  eventsToday.addEmailReminder(15);
  
  Logger.log(!calendars[0].isOwnedByMe());
  if (!calendars[0].isOwnedByMe())
  {
    calendars[0].unsubscribeFromCalendar(); //会議室のカレンダーの登録を削除する / Unsubscribe from meeting room calendar
  }
  return true;
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
  if(!meeting_calendar[0].isOwnedByMe())
  {
    meeting_calendar[0].unsubscribeFromCalendar();
  }
  
  // Logger.log(event);
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
  var calendar_name = 'アルバイト'  
  // var today = new Date();
  // Use below function to get today date in JST format
  var today = today_jst();
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
  // var today = new Date();
  // Use below function to get today date in JST format
  var today = today_jst();
  var email = Session.getActiveUser().getEmail();

  Logger.log(today)
  for(var i=0 ; i<calendar_names.length ; i++)
  { 
    var cell_id = "E"+(i+2);
    sheet.getRange(cell_id).clearContent(); 

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
          // var string_events = sheet.getRange(cell_id).getValues()+"\n"; 
          var string_events = ""; 

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

function today_jst()
{ 
  // By default, new Date() function in Google Apps Script gets the standard time from America/Los_Angeles (Pacific time)  
  var today = new Date(); 
  var today_jst = new Date(today.getTime()+1000*60*60*14);
  Logger.log(today_jst);
  return today_jst;
}

function active_user()
{
  var email_user = Session.getActiveUser().getEmail();
  var url_ss = "https://docs.google.com/spreadsheets/d/1Om1kYwsVAISmAS8LnI8S2_INkpf0Q33-35GLhbY_jp0/edit#gid=0";
  var ss = SpreadsheetApp.openByUrl(url_ss);
  var sheet = ss.getSheetByName("Data");
  var list_email = sheet.getRange(2,8,sheet.getLastRow()-1).getValues();
  // Logger.log(find_name(list_email,email_user)-1)
  return find_name(list_email,email_user)-1;
}

function send_reminder() 
{
  var email = "jmsaenz@silk.jp";
  var subject = "Reminder";
  var body = "Reminder";

  // Send yourself an email with a link to the document.
  GmailApp.sendEmail(email, subject, body);
}

// function saving_booking_onEdit()
// {
//   var url_ss = "https://docs.google.com/spreadsheets/d/1Om1kYwsVAISmAS8LnI8S2_INkpf0Q33-35GLhbY_jp0/edit#gid=0";
//   var ss = SpreadsheetApp.openByUrl(url_ss);
//   var sheet = ss.getSheetByName("Data");
//   var list_booking = sheet.getRange(2,6,sheet.getLastRow()-1).getValues();
//   var list_id = sheet.getRange(2,7,sheet.getLastRow()-1).getValues();
//   for (var i = 0 ; i < list_booking.length ; i++)
//   {
//     if (list_booking[i]!=="" || list_id[i]!=="")
//     {
//       Logger.log(list_booking[i])
//       Logger.log(list_id[i])
//     }
//   }
  // Logger.log(list_booking)
  // Logger.log(list_id)
// }

//// Function to create installable trigger onEdit()
// function createSpreadsheetOnEditTrigger() 
// {
// }

// function test()
// {
// }
