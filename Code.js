var meeting_calendar_id = 'c_dmhmoef7aksb1hiu0n2taobjhk@group.calendar.google.com';
var reserveDate='nada';
//var url_ss = "https://docs.google.com/spreadsheets/d/1p7kIRtdrElo-QaKkrnGTjE8DTch23OrmHag8gDJRgqU/edit#gid=0";
  var url_ss = "https://docs.google.com/spreadsheets/d/1Om1kYwsVAISmAS8LnI8S2_INkpf0Q33-35GLhbY_jp0/edit#gid=0";
var email_main_access = "part-timer@silk.co.jp"


function doGet(e) 
{
  
  var x = isNewUser();
  Logger.log(Session.getActiveUser().getEmail())
  
  var x = 0;
  //Logger.log(Session.getActiveUser().getEmail())
  //Logger.log(x)
  if(x==0){
    return HtmlService.createTemplateFromFile("welcome").evaluate(); 
  }
  else if ( x != -1)
  {
    return HtmlService.createTemplateFromFile("new-user").evaluate(); 
  }
  else
  {
    return HtmlService.createTemplateFromFile("main").evaluate();
  }
}

function include(file_name)
{
  return HtmlService.createHtmlOutputFromFile(file_name).getContent();
}

function saveUser(userInfo)
{
    var ss = SpreadsheetApp.openByUrl(url_ss);
    var sheet = ss.getSheetByName("Data");
    var cell_status = "C" + (parseInt(userInfo.row)+1).toString();
    var cell_note = "D" + (parseInt(userInfo.row)+1).toString();
    sheet.getRange(cell_status).setValue(userInfo.status.trim());
    sheet.getRange(cell_note).setValue(userInfo.note);
}

function saveBooking(bookingInfo)
{ 
  var ss = SpreadsheetApp.openByUrl(url_ss);
  var sheet = ss.getSheetByName("Booking");
  var row = parseInt(bookingInfo.row)+1;
  var last_column = sheet.getRange(row,2).getValue()*2+3;
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

    if ( ( (start_time >= s1 && start_time < e1) || (end_time > s1 && end_time < e1) ) || ((s1 > start_time && s1 < end_time) && (e1 > start_time && e1 < end_time)))
    { 
      Logger.log('Event exists!')
      return false;
    }
  }

  const eventsToday = calendars[0].createEvent(calendars[0].getName() +" (" + bookingInfo.name + ")", start_time, end_time);//イベントを作�Eする / Create Event
  sheet.getRange(row,last_column).setValue(bookingInfo.date + " " + bookingInfo.startTime + " - " + bookingInfo.finishTime);
  sheet.getRange(row,last_column+1).setValue(eventsToday.getId());
  //eventsToday.addEmailReminder(15);
  sheet.getRange(row,2).setValue(sheet.getRange(row,2).getValue()+1);

  Logger.log(!calendars[0].isOwnedByMe());
  if (!calendars[0].isOwnedByMe())
  {
    calendars[0].unsubscribeFromCalendar(); //会議室のカレンダーの登録を削除する / Unsubscribe from meeting room calendar
  }
  return true;
}

function deleteBooking(list)
{
  var ss = SpreadsheetApp.openByUrl(url_ss);
  var sheet = ss.getSheetByName("Booking");
  var row = parseInt(list.row)+1;
  var c;
  var event_id = [];
  var num_booking = sheet.getRange(row,2).getValue();

  for (var i = 0 ; i<list.list.length ; i++)
  {
    c = parseInt(list.list[i])*2+2-2*i;
    event_id[i] = sheet.getRange(row,c).getValue();
    sheet.getRange(row,c-1,1,2).deleteCells(SpreadsheetApp.Dimension.COLUMNS);
    // Logger.log(sheet.getRange(row,c-1,1,2).getValues());
    num_booking -- ;
  }

  Logger.log(event_id);
  sheet.getRange(row,2).setValue(num_booking);
  
 //会議室のカレンダーに登録してアクセスする / Subscribe to the meeting room calendar and access it
  var meeting_calendar=CalendarApp.getCalendarsByName(CalendarApp.subscribeToCalendar(meeting_calendar_id).getName());
  
  for (var i = 0 ; i<event_id.length ; i++)
  {
    var event = meeting_calendar[0].getEventById(event_id[i]);
    event.deleteEvent();
  }

  //会議室のカレンダーの登録を削除する / Unsubscribe from meeting room calendar
  if(!meeting_calendar[0].isOwnedByMe())
  {
    meeting_calendar[0].unsubscribeFromCalendar();
  }
  
  // Logger.log(event);
}

function data_from_ss()
{
  var ss = SpreadsheetApp.openByUrl(url_ss);
  var sheet = ss.getSheetByName("Data");
  
  r = sheet.getLastRow();
  datass = {}
  datass.data = sheet.getSheetValues(2, 2, r-1, 3);
  
  datass.index=find_name(sheet.getSheetValues(2, 6, r-1, 1) , Session.getActiveUser().getEmail())-1;

  // Logger.log(sheet.getSheetValues(2, 6, r-1, 2));
  Logger.log(datass.data);
  Logger.log(datass.data[1][0]);
  return datass;
}

// function read_calendar()
// {
//   var calendar_name = 'アルバイト'  
//   var today = new Date();
//   // Use below function to get today date in JST format
//   today = date_jst(today);
//   var calendar = CalendarApp.getCalendarsByName(calendar_name)[0].getEventsForDay(today);
  
//   if (calendar.length == 0)
//   {
//     return;
//   }
//   var ss = SpreadsheetApp.openByUrl(url_ss);
//   var sheet = ss.getSheetByName("Data");

//   list = sheet.getSheetValues(2, 2, sheet.getLastRow()-1, 1);

//   for (var i = 0 ; i < calendar.length ; i++ )
//   {
//     var st = Utilities.formatDate(calendar[i].getStartTime(), "GMT+9", "HH:mm");
//     var et = Utilities.formatDate(calendar[i].getEndTime(), "GMT+9", "HH:mm");
//     var cell_id = "E"+find_name(list, calendar[i].getTitle());
//     sheet.getRange(cell_id).setValue(calendar_name + " " + st + " - " + et);
    
//   }
// }

// function read_calendar_events_today()
// {
//   var ss = SpreadsheetApp.openByUrl(url_ss);
//   var sheet = ss.getSheetByName("Data");
//   var calendar_names=sheet.getRange(2,6,sheet.getLastRow()-1).getValues();
//   var today = new Date();
  
//   // Use below function to get today date in JST format
//   today = date_jst(today);
//   var email = Session.getActiveUser().getEmail();

//   // Logger.log(today)
//   for(var i=0 ; i<calendar_names.length ; i++)
//   { 
//     var cell_id = "E"+(i+2);
//     sheet.getRange(cell_id).clearContent(); 

//     if (calendar_names[i].toString()!=="")
//     {
//       var calendar_name = calendar_names[i].toString();
      
//       var calendar  =  Calendar.Calendars.get(calendar_name);
//       if(typeof calendar !== 'undefined')
//       {
//         if (calendar_name !== email)
//         {
//           var aCal=CalendarApp.subscribeToCalendar(calendar.id);
//           var eventToday=CalendarApp.getCalendarsByName(aCal.getName())[0].getEventsForDay(today);
//         }
//         else
//         {
//           var eventToday = CalendarApp.getCalendarById(calendar.id).getEventsForDay(today);
//         }

//         if(eventToday.length>0)
//         { 
//           // var string_events = sheet.getRange(cell_id).getValues()+"\n"; 
//           var string_events = ""; 

//           for (var j = 0 ; j<eventToday.length ; j++)
//           {
//             // Logger.log(eventToday.length);
//             var theEvent = eventToday[j];
//             var st = Utilities.formatDate(theEvent.getStartTime(), "GMT+9", "HH:mm");
//             var et = Utilities.formatDate(theEvent.getEndTime(), "GMT+9", "HH:mm");
            
//             string_events = string_events + theEvent.getTitle() + " " + st + " - " + et + "\n";
//             // Logger.log(string_events);

//           }
//           sheet.getRange(cell_id).setValue(string_events); 
//         }
//         if (calendar_name !== email)
//         {
//           aCal.unsubscribeFromCalendar();
//         }
          
//       }
//       else
//       {
//       }     
//     }
//   }
// } 

function read_calendar_date()
{ 
  var date = new Date();
  date = date_jst(new Date(date));
  var ss = SpreadsheetApp.openByUrl(url_ss);
  var sheet = ss.getSheetByName("Data");
  var calendar_names=sheet.getRange(2,6,sheet.getLastRow()-1).getValues();
  var email = Session.getActiveUser().getEmail();
  
  var list_events = [];
  var k = 0 ;
 
  list_events[k] = Array(calendar_names.length);
  
  for(var i = 0 ; i<calendar_names.length ; i++)
  { 
    var cell_id = "E"+(i+2);
    list_events[k][i] = "";
    
    if (calendar_names[i].toString()!=="")
    {
      var calendar_name = calendar_names[i].toString();
      
      try
      {
        var calendar  =  Calendar.Calendars.get(calendar_name);
      }
      catch(err)
      {
        Logger.log(calendar_name+" has to change his/her rights to access calendar")
        list_events[k][i] = "?Check permissions"
        continue;
      }
      
      if(typeof calendar !== 'undefined')
      {
        // if (calendar_name !== email)
        // {
        //   var aCal=CalendarApp.subscribeToCalendar(calendar.id);
        //   var eventToday=CalendarApp.getCalendarsByName(aCal.getName())[0].getEventsForDay(date);
        // }
        // else
        // {
        //   var eventToday = CalendarApp.getCalendarById(calendar.id).getEventsForDay(date);
        // }
        // Logger.log(calendar.id)
        var aCal=CalendarApp.subscribeToCalendar(calendar.id);
        
        try
        {
          var eventToday = CalendarApp.getCalendarsByName(aCal.getName())[0].getEventsForDay(date);
        }
        catch(err)
        {
          Logger.log(calendar_name+" has no calendars to read")
          Logger.log(calendar_name+" has to change his/her rights to access calendar")
          list_events[k][i] = "?Check permissions"
        
          aCal.unsubscribeFromCalendar();
          continue;
        }
        

        if(eventToday.length>0)
        { 
          var string_events = ""; 
          for (var j = 0 ; j<eventToday.length ; j++)
          {
            // Logger.log(eventToday.length);
            var theEvent = eventToday[j];
            var st = Utilities.formatDate(theEvent.getStartTime(), "GMT+9", "HH:mm");
            var et = Utilities.formatDate(theEvent.getEndTime(), "GMT+9", "HH:mm");
            
            string_events = string_events + theEvent.getTitle() + " " + st + " - " + et + "\n";
          }
          list_events[k][i] = string_events;
          sheet.getRange(cell_id).setValue(string_events); 
        }
        if (calendar_name !== email_main_access)
        {
          aCal.unsubscribeFromCalendar();
        }
          
      }    
    }
  } 
  
  Logger.log(list_events)
  return list_events;
}  

function read_new_date(number_days)
{ 
  
  var date = new Date();
  date = date_jst(new Date(date));
  var ss = SpreadsheetApp.openByUrl(url_ss);
  var sheet = ss.getSheetByName("Data");
  var calendar_names=sheet.getRange(2,6,sheet.getLastRow()-1).getValues();
  var email = Session.getActiveUser().getEmail();
  date = new Date(date.getTime()+1000*60*60*24*number_days);
  var new_list = [];
  // Logger.log(date)
  // return
    
  for(var i = 0 ; i<calendar_names.length ; i++)
  {
    new_list[i] = "";
    
    if (calendar_names[i].toString()!=="")
    {
      var calendar_name = calendar_names[i].toString();
      
      try
      {
        var calendar  =  Calendar.Calendars.get(calendar_name);
      }
      catch(err)
      {
        Logger.log(calendar_name+" has to change his/her rights to access calendar")
        new_list[i] = "?Check permissions"
        continue;
      }

      if(typeof calendar !== 'undefined')
      {
        
        if (calendar_name !== email)
        {
          var aCal=CalendarApp.subscribeToCalendar(calendar.id);
          try
          {
            var eventToday=CalendarApp.getCalendarsByName(aCal.getName())[0].getEventsForDay(date);
          }
          catch(err)
          {
            Logger.log(calendar_name+" has no calendars to read")
            new_list[i] = "?Check permissions";
            aCal.unsubscribeFromCalendar();
            continue;
          }  
        }
        else
        {
          var eventToday = CalendarApp.getCalendarById(calendar.id).getEventsForDay(date);
        }

        if(eventToday.length>0)
        { 
          var string_events = ""; 
          for (var j = 0 ; j<eventToday.length ; j++)
          {
            // Logger.log(eventToday.length);
            var theEvent = eventToday[j];
            var st = Utilities.formatDate(theEvent.getStartTime(), "GMT+9", "HH:mm");
            var et = Utilities.formatDate(theEvent.getEndTime(), "GMT+9", "HH:mm");
            
            string_events = string_events + theEvent.getTitle() + " " + st + " - " + et + "\n";
          }
          new_list[i] = string_events;
        }
        if (calendar_name !== email)
        {
          aCal.unsubscribeFromCalendar();
        }
          
      }    
    }
  } 
  // Logger.log(new_list)
  return new_list;
  // Logger.log(new_list)
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

function test()
{

}

function date_jst(date)
{ 
  // By default, new Date() function in Google Apps Script gets the standard time from America/Los_Angeles (Pacific time)   
  var date_jst = new Date(date.getTime()+1000*60*60*14);
  // Logger.log(date_jst);
  return date_jst;
}

function active_user()
{
  var email_user = Session.getActiveUser().getEmail();
  var ss = SpreadsheetApp.openByUrl(url_ss);
  var sheet = ss.getSheetByName("Data");
  var list_email = sheet.getRange(2,6,sheet.getLastRow()-1).getValues();
  var user = {};
  
  user.index=find_name(list_email,email_user)-1; 
  user.name = sheet.getRange(user.index+1,2,1).getValues()+"さん";
  user.email =  email_user;

  Logger.log(user.email);
  return user;
}

function read_booking_active_user(row)
{
  var ss = SpreadsheetApp.openByUrl(url_ss);
  var sheet = ss.getSheetByName("Booking");
  // r = 5;
  r = parseInt(row);
  c = sheet.getRange(r+1,2).getValue();
  data = [];
  data[0] = r;

  if (c!=0)
  { 
    var d = sheet.getSheetValues(r+1, 3, 1, c*2);
    
    var i = 1;
    for (var j = 1 ; j<c*2 ; j=j+2)
    {
      data[i] = d[0][j-1];
      i++;
        // Logger.log(d[i][j])
    }
  }else
  {
    data[1] = "";
  }  
  
  // Logger.log(data);
  return data;
}

function read_booking()
{
  var ss = SpreadsheetApp.openByUrl(url_ss);
  var sheet = ss.getSheetByName("Booking");

  var data_ss = sheet.getRange(2,2,sheet.getLastRow()-1,sheet.getLastColumn()-1).getValues();
  // Logger.log(data_ss)
  data = [];

  for (var i = 0 ; i < sheet.getLastRow()-1 ; i++)
  { 
    data[i] = "";
    if (data_ss[i][0]!=0)
    {
      for (var j = 1 ; j < data_ss[i][0]*2 ; j=j+2)
      {
        data[i] = data[i]+data_ss[i][j]+"<br>";
      }
    }
  }
  Logger.log(data)
  return data;
}

// Trigger to delete one day before bookings. It fires around 12:00am.
// It also updates the spreadsheet
function deleteBooking_trigger()
{
  var yesterday = new Date();
  yesterday = new Date(yesterday.getTime()-1000*60*60*10);
  // Logger.log(yesterday);

  var calendars = CalendarApp.getCalendarsByName(CalendarApp.subscribeToCalendar(meeting_calendar_id).getName())[0].getEventsForDay(yesterday);
  // Logger.log(calendars.length);
  var list = [];
  for (var i = 0 ; i < calendars.length ; i++)
  {
    list[i] = calendars[i].getId();
  }
  // Logger.log(list);

  var ss = SpreadsheetApp.openByUrl(url_ss);
  var sheet = ss.getSheetByName("Booking");
  var data_ss = sheet.getRange(2,3,sheet.getLastRow()-1,sheet.getLastColumn()-1).getValues();
  var num_booking;
  var index;
  for (var j = 0 ; j<list.length ; j++)
  { 
    calendars[j].deleteEvent();
    for (var i = 0 ; i<data_ss.length ; i++)
    {
      num_booking = sheet.getRange(i+2,2).getValue();
      index = data_ss[i].indexOf(list[j]);
      if (index != -1)
      {
        // Logger.log(sheet.getRange(i+2,index+2,1,2).getValues())
        sheet.getRange(i+2,index+2-2*j,1,2).deleteCells(SpreadsheetApp.Dimension.COLUMNS);
        num_booking--;
      }
      sheet.getRange(i+2,2).setValue(num_booking);
    } 
  }

  // the following lines clear the data in the columns "status" and "memo"
  var sheet2 = ss.getSheetByName("Data");
  sheet2.getRange(2,3,sheet2.getLastRow()-1,2).clearContent();

  // updates the spreadsheet
  // read_calendar_events_today();
  
}

// function getScriptURL() // This function is used to reload the page 
// {
//   return ScriptApp.getService().getUrl();
// }

function getUserEmails()
{
  var ss = SpreadsheetApp.openByUrl(url_ss);
  var sheet = ss.getSheetByName("Data");
  var employees = {}
  employees.list_email = sheet.getRange(2,6,sheet.getLastRow()-1).getValues();
  employees.list_names = sheet.getRange(2,2,sheet.getLastRow()-1).getValues();
  // Logger.log(employees.list_email[1])
  // Logger.log(employees.list_names.length)
  return employees;
}

function sendMessage(msg)
{
  Logger.log(msg.email)
  Logger.log(msg.memo)
  var subject = Session.getActiveUser().getEmail()+"が送ったダッシュボードからの伝言";
  // GmailApp.sendEmail(msg.email, subject, msg.memo);
  var email_list = getUserEmails().list_email;
  var index = find_name(email_list,msg.email.trim())-1;
  SpreadsheetApp.openByUrl(url_ss).getSheetByName("Data").getRange(index+1,7).setValue("1");

  return index;
}

function readInfoRefresh()
{
  var ss = SpreadsheetApp.openByUrl(url_ss);
  var sheet = ss.getSheetByName("Data");
  
  var r = sheet.getLastRow();
  
  datass = sheet.getSheetValues(2, 3, r-1, 3);
  for (var i = 0 ; i<datass.length ; i++)
  {
    if (datass[i][2]!=="")
    {
      datass[i][2] = datass[i][2].replace(/\n/g,"<br>"); // /g all matches
    } 
  }

  sheet = ss.getSheetByName("Booking");
  var data_ss = sheet.getRange(2,2,r-1,sheet.getLastColumn()-1).getValues();
  // Logger.log(data_ss)
  var data;
  var list_flags_mgs = readMessages();

  for (var i = 0 ; i < r-1 ; i++)
  { 
    data= "";
    if (data_ss[i][0]!=0)
    {
      for (var j = 1 ; j < data_ss[i][0]*2 ; j=j+2)
      {
        data = data+data_ss[i][j]+"<br>";
      }
      
    }
    Logger.log(r)
    Logger.log(data)
    datass[i][3] = data;
    datass[i][4] = list_flags_mgs[i];
  }

  Logger.log(datass)
  return datass; 
}

function readMessages()
{
  var ss = SpreadsheetApp.openByUrl(url_ss);
  var sheet = ss.getSheetByName("Data");
  
  var r = sheet.getLastRow();
  var msg_flags = sheet.getSheetValues(2, 7, r-1, 1);
  
  Logger.log(msg_flags);
  return msg_flags;
}

function alreadyReadMsg(index)
{
  
  Logger.log(index)
  SpreadsheetApp.openByUrl(url_ss).getSheetByName("Data").getRange(parseInt(index)+1,7).setValue("0");
}

function isNewUser()
{
  var list_users = getUserEmails().list_email;
  var current_user = Session.getActiveUser().getEmail();
  var index_user = find_name(list_users, current_user);
  if (index_user == null)
  {
    // add_new_user(current_user);
    // Logger.log("New user: "+current_user);
    return current_user;
  }
  return -1;
}

function add_new_user(name_user) 
{
  var current_user = Session.getActiveUser().getEmail();
  var ss = SpreadsheetApp.openByUrl(url_ss);
  var sheet = ss.getSheetByName("Data");
  var list_names = sheet.getRange(2,2,sheet.getLastRow()-1).getValues();
  var index_user = find_name(list_names, name_user.trim());
  // Logger.log(sheet.getLastRow())
  if (index_user == null)
  {
    var row_index = sheet.getLastRow()+1;
    sheet.getRange(row_index,1).setValue(row_index-1);
    sheet.getRange(row_index,2).setValue(name_user.trim());
    sheet.getRange(row_index,6).setValue(current_user.toString());
    sheet.getRange(row_index,7).setValue(0);
    
  }
  else
  {
    sheet.getRange(index_user,6).setValue(current_user.toString());
  }
}

function getScriptURL() // This function is used to reload the page 
{
  Logger.log(ScriptApp.getService().getUrl());
  return ScriptApp.getService().getUrl();
}

function read_all_calendars10()
{
  var ss = SpreadsheetApp.openByUrl(url_ss);
  var sheet = ss.getSheetByName("Data");
  var calendar_names=sheet.getRange(2,6,sheet.getLastRow()-1).getValues();

  var date = new Date();

  var array_dates = [];
  array_dates[0]= date_jst(new Date(date));

  for (var i = 1 ; i<=10 ; i++)
  {
    array_dates[i] = new Date(array_dates[0].getTime()+1000*60*60*24*i);
  }
  
  var list_events = Array(calendar_names.length);

  for (var n = 0 ; n<list_events.length ; n++)
  {
    list_events[n] = Array(array_dates.length);
  }
  
  for (var i = 0 ; i<calendar_names.length ; i++)
  {
    if (calendar_names[i].toString()!=="")
    {
      var calendar_name = calendar_names[i].toString();
      
      try
      {
        var calendar  =  Calendar.Calendars.get(calendar_name);
      }
      catch(err)
      {
        Logger.log(calendar_name+" has to change his/her rights to access calendar")
        // list_events[i][j] = "?Check permissions"
        for (var n = 0 ; n<array_dates.length ; n++)
        {
          // sheetC.getRange(i+1,n).setValue("?Check permissions");
          list_events[i][n]="?Check permissions";
        }
        
        continue;
      }

      Logger.log(calendar_name)
      if(typeof calendar !== 'undefined')
      {
        
        var aCal=CalendarApp.subscribeToCalendar(calendar.id);
        
        for (var j = 0 ; j<array_dates.length ; j++ )
        {
          
          try
          {
            var eventToday = CalendarApp.getCalendarsByName(aCal.getName())[0].getEventsForDay(array_dates[j]);
          }
          catch(err)
          {
            Logger.log(calendar_name+" has no calendars to read")
            Logger.log(calendar_name+" has to change his/her rights to access calendar")
            sheetC.getRange(i+1,j+1).setValue("?Check permissions");
            list_events[i][j]="?Check permissions";
            // aCal.unsubscribeFromCalendar();
            continue;
          }
          
          var string_events = " ";
          if(eventToday.length>0)
          {             
            for (var k = 0 ; k<eventToday.length ;k++)
            {
              // Logger.log(eventToday.length);
              var theEvent = eventToday[k];
              var st = Utilities.formatDate(theEvent.getStartTime(), "GMT+9", "HH:mm");
              var et = Utilities.formatDate(theEvent.getEndTime(), "GMT+9", "HH:mm");
              
              string_events = string_events + theEvent.getTitle().toString() + " " + st.toString() + " - " + et.toString() + "\n";
            }
        
            // sheetC.getRange(i+1,j+1).setValue(string_events); 
          }
          list_events[i][j] = string_events;
          
        }
        if (calendar_name !== email_main_access)
        {
          aCal.unsubscribeFromCalendar();
        }  
      }    
    }
  }
  
  var sheetC = ss.getSheetByName("Calendar");
  for(var i=0 ; i<calendar_names.length ; i++)
  {
    for (var j=0 ; j<array_dates.length ; j++)
    {
      sheetC.getRange(i+1,j+1).setValue(list_events[i][j])
    }
  }
  // Logger.log(list_events)
  return list_events;
}

function read_ss_calendars()
{
  var ss = SpreadsheetApp.openByUrl(url_ss);
  var sheetC = ss.getSheetByName("Calendar");
  Logger.log(sheetC.getRange(1,1,sheetC.getLastRow(),sheetC.getLastColumn()-1).getValues());
  return sheetC.getRange(1,1,sheetC.getLastRow(),sheetC.getLastColumn()-1).getValues();
}
