'use strict';
/*
.SYNOPSIS
    Script is checking a server for downtimes above 5 min within given period of time (30 DAYS).
.DESCRIPTION
    Script is checking a server for downtimes above 5 min within given period of time (30 DAYS).
    It is possible to pass arguments to the script in order to change downtime interval and 
    desired searching date for another number of days. 
    First argument after the script is number of days, and the second is downtime interval in minutes.
    Downtime is rounded down to MINUTES value. E.g. 6:59 will be rounded to 6 min
    If downtime took 5:59, then it is first rounded to 5 MINUTES and then checked against 
    specified downtime interval.
    In that case it would not be recorded to log, because rounded value (5) is equal 
    (but not higher) to downtime interval (5).
    Script uses WMI object which is not available on w2k, so it will be usable on >= w2k3 systems!
    For better readability of the report use Excel with "tab" separated entries.
.NOTES
    File Name      : availability_3.3.js
    Author         : Damian Danak
    Prerequisite   : Windows Server 2003 and higher.
    Tested         : Target machines with Windows Server 2003 to 2012 R2
    Version        : 3.3.0
    Date           : 12-09-2016
.EXAMPLE
    On Windows console type:
    cscript //nologo availability_3.3.js > availablity_report.log
.EXAMPLE
    On Windows console type:
    cscript //nologo availability_3.3.js 30 5 > availablity_report.log
.HISTORY
    v3.3 09-09-2016 16:00:00 Rewriting script - logic stays mostly the same
    v1.0 11-12-2014 10:50:43 Initial code created
.LICENCE
    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.
*/

//
// Variables declaration
//

var CURRENT_TIME = new Date(),
  downtimeCount = 1,
  eventsCount = 0,
  startDate,
  startDateOut,
  DOWNTIME_INTERVAL = 5,
  DOWNTIME_WARNING = 720,
  TIME_BACK = -30,
  EVT_ID_6006 = 6006,
  EVT_ID_6005 = 6005,
  EVT_ID_1074 = 1074,
  EVT_ID_1076 = 1076,
  EVT_ID_6008 = 6008,
  EVT_ID_6009 = 6009,
  EVT_ID_41 = 41,
  MINUTES = 1000 * 60,
  loggedEvents,
  logged41,
  logged6008,
  logged1074,
  loggedExpected,
  msgNoLogs = "",
  msgNoReboots = "",
  msgWarning = "",
  msgShortDowntime = "There was no downtime above " + DOWNTIME_INTERVAL +
    " min found on the server.",
  enumEvt41,
  comment41,
  msg41,
  msgUsr41,
  msgTime41,
  objEvent41,
  upEventTime41,
  enumEvtIrreg,
  downStampUnexp,
  upStampUnexp,
  unexpectedDowntime,
  msgUnexpectedDowntime,
  msg6008,
  comment6008Fail,
  comment6008Success,
  msgTimeDown,
  msgTimeUp,
  objEvent6008,
  upEventTime6008,
  enumEvtReg,
  downStampExp,
  upStampExp,
  expectedDowntime,
  msgExpectedDowntime,
  msg1074 = "",
  msg2nd1074 = "",
  commentUsr1074 = "N/A",
  comment1074 = "",
  msgUsr1074,
  msgTime6005,
  e6005,
  msgTime6006,
  e6006,
  count6006 = 0,
  count1074 = 0,
  objEventReg,
  commentAtNoEvent,
  HOSTNAME = ".",
  WMI_NAMESPACE = "\\root\\cimv2",
  wmiObject = GetObject( "winmgmts:{impersonationLevel=impersonate}!\\\\" + 
    HOSTNAME + WMI_NAMESPACE ),
  wmiDateObject = WScript.CreateObject( "WbemScripting.SWbemDateTime" ),
  wshShell = WScript.CreateObject( "WScript.Shell" ),
  COMPUTER_NAME = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" );

//
// Setting up variables
//

if ( WScript.Arguments.Count() ) {
  TIME_BACK = WScript.Arguments.Item( 0 );
  if ( TIME_BACK > 0 ) {
    TIME_BACK = - ( TIME_BACK );
  }
  if ( WScript.Arguments.Count() == 2 ) {
    DOWNTIME_INTERVAL = WScript.Arguments.Item( 1 );
  }
}
startDate = dateToUTCString( calculateTime( CURRENT_TIME, TIME_BACK ) );
startDateOut = convertUTCtoDateString( startDate );
loggedEvents = wmiObject.ExecQuery( "SELECT * FROM Win32_NTLogEvent WHERE " +
  "(Logfile = 'System' AND " +
  "TimeGenerated > '" + startDate + "') AND " +
  "(EventCode = '" + EVT_ID_6005 + "' OR " +
  "EventCode = '" + EVT_ID_1074 + "' OR " +
  "EventCode = '" + EVT_ID_1076 + "' OR " +
  "EventCode = '" + EVT_ID_41 + "' OR " +
  "EventCode = '" + EVT_ID_6008 + "' OR " +
  "EventCode = '" + EVT_ID_6009 + "' OR " +
  "EventCode = '" + EVT_ID_6006 + "')" );
logged41 = wmiObject.ExecQuery( "SELECT * FROM Win32_NTLogEvent WHERE " +
  "(Logfile = 'System' AND " +
  "TimeGenerated > '" + startDate + "') AND " +
  "(EventCode = '" + EVT_ID_41 + "')" );
logged6008 = wmiObject.ExecQuery( "SELECT * FROM Win32_NTLogEvent WHERE " +
  "(Logfile = 'System' AND " +
  "TimeGenerated > '" + startDate + "') AND " +
  "(EventCode = '" + EVT_ID_6008 + "')" );
logged1074 = wmiObject.ExecQuery( "SELECT * FROM Win32_NTLogEvent WHERE " +
  "(Logfile = 'System' AND " +
  "TimeGenerated > '" + startDate + "') AND " +
  "(EventCode = '" + EVT_ID_1074 + "')" );
loggedExpected = wmiObject.ExecQuery( "SELECT * FROM Win32_NTLogEvent WHERE " +
  "(Logfile = 'System' AND " +
  "TimeGenerated > '" + startDate + "') AND " +
  "(EventCode = '" + EVT_ID_1074 + "' OR " +
  "EventCode = '" + EVT_ID_6006 + "' OR " +
  "EventCode = '" + EVT_ID_6005 + "')" );

//
// Start of actual script
//

if ( ( !logged6008.Count ) || ( !logged41.Count ) ) {
  downtimeCount = 0;
  msgNoReboots = "No unexpected reboots have been detected on the server. ";
}
if ( !loggedEvents.Count ) {
  downtimeCount = 0;
  msgNoLogs = "No events related to the downtime have been found on the server. ";
}

if ( logged41.Count ) {
  enumEvt41 = new Enumerator( logged41 );
  for (; !enumEvt41.atEnd(); enumEvt41.moveNext() ) {
    objEvent41 = enumEvt41.item();
    if ( objEvent41.EventCode == EVT_ID_41 && 
    ( objEvent41.SourceName == "Microsoft-Windows-Kernel-Power" ||
    objEvent41.SourceName == "Kernel-Power" ) ) {
      downtimeCount = 1;
      msg41 = objEvent41.Message;
      msgUsr41 = objEvent41.User;
      wmiDateObject.Value = objEvent41.TimeGenerated;
      upEventTime41 = wmiDateObject.GetVarDate( true );
      upEventTime41 = new Date(upEventTime41);
      msgTime41 = convertUTCtoDateString( timeToUTCString( upEventTime41 ) );
      comment41 = "Unexpected reboot has been detected [Event ID 41], " +
        "but it was not taken to calculate downtime. " +
        "Please compare it to the associated logs (e.g. [Event ID 6008]). " +
        "If no associated events have been logged, " +
        "then downtime probably took less than " + 
        DOWNTIME_INTERVAL + " minutes. " +
        "For detailed information please refer to the " +
        "Windows System Log [Event ID 41] from " + msgTime41;
      writeLog( "N/A", msgTime41, "N/A", msgUsr41, comment41 );
    }
  }
}

if ( logged6008.Count ) {
  enumEvtIrreg = new Enumerator( logged6008 );
  for ( ; !enumEvtIrreg.atEnd(); enumEvtIrreg.moveNext() ) {
    objEvent6008 = enumEvtIrreg.item();
    if ( ( objEvent6008.EventCode == EVT_ID_6008 ) && ( objEvent6008.SourceName == "EventLog" ) ) {
      msg6008 = objEvent6008.Message;
      wmiDateObject.Value = objEvent6008.TimeGenerated;
      upEventTime6008 = wmiDateObject.GetVarDate( true );
      upEventTime6008 = new Date( upEventTime6008 );
      msgTimeUp = convertUTCtoDateString( timeToUTCString( upEventTime6008 ) );
      upStampUnexp = upEventTime6008.getTime();
      msgTimeDown = findDate( objEvent6008.Message );
      comment6008Fail = "Cannot calculate downtime. For details please refer to the " +
        "Windows System Log [Event ID 6008] from " + msgTimeUp;
      if ( msgTimeDown ) {
        msgTimeDown = convertUTCtoDateString( timeToUTCString( msgTimeDown ) );
        downStampUnexp = findDate( objEvent6008.Message ).getTime();
        msgUnexpectedDowntime = Math.floor( ( upStampUnexp - downStampUnexp ) / MINUTES );
      } else if ( !msgTimeDown ) {
          downtimeCount = 1;
          writeLog( "N/A", "N/A", "N/A", "N/A", comment6008Fail );
      } else {}
      if ( ( upStampUnexp - downStampUnexp > 0 ) && 
      ( downStampUnexp != 0 ) && 
      ( msgUnexpectedDowntime > DOWNTIME_INTERVAL ) ) {
        downtimeCount = 1;
        unexpectedDowntime = upStampUnexp - downStampUnexp;
        msgWarning = createMsgWarning ( msgUnexpectedDowntime );
        comment6008Success = "Message from the Windows System Log [Event ID 6008]: " +
          msg6008.replace(/(\r\n|\n|\r)/gm," ") + msgWarning;
        writeLog( msgTimeDown, msgTimeUp, msgUnexpectedDowntime, "N/A", comment6008Success );
      }
    }
  }
}

if ( loggedExpected.Count ) {
  enumEvtReg = new Enumerator( loggedExpected );
  for ( enumEvtReg.moveFirst(); !enumEvtReg.atEnd(); enumEvtReg.moveNext() ) {
    objEventReg = enumEvtReg.item();
    if ( (upStampExp - downStampExp > 0 ) &&
    ( upStampExp - downStampExp != expectedDowntime ) &&
    ( logged1074.Count > 0 ) && 
    ( objEventReg.EventCode != EVT_ID_1074 ) &&
    ( msgExpectedDowntime > DOWNTIME_INTERVAL ) ) {
      downtimeCount = 1;
      msgWarning = createMsgWarning( msgExpectedDowntime );
      comment1074 = msgWarning + comment1074;
      expectedDowntime = upStampExp - downStampExp;
      writeLog( msgTime6006, msgTime6005, msgExpectedDowntime, commentUsr1074, comment1074 );
    } else if ( ( upStampExp - downStampExp > 0 ) &&
      ( upStampExp - downStampExp != expectedDowntime ) &&
      ( logged1074.Count == 0 ) && ( msgExpectedDowntime > DOWNTIME_INTERVAL ) ) {
        downtimeCount = 1;
        msgWarning = createMsgWarning( msgExpectedDowntime );
        comment1074 = "Ordinary shutdown. No related comments have been found. " + msgWarning;
        expectedDowntime = upStampExp - downStampExp;
        writeLog( msgTime6006, msgTime6005, msgExpectedDowntime, commentUsr1074, comment1074 ); 
    } else {}
    if ( objEventReg.EventCode == EVT_ID_6005 ) {
      wmiDateObject.Value = objEventReg.TimeGenerated;
      e6005 = wmiDateObject.GetVarDate( true );
      msgTime6005 = new Date( e6005 );
      upStampExp = msgTime6005.getTime();
      msgTime6005 = convertUTCtoDateString( timeToUTCString( msgTime6005 ) );
    } else if ( objEventReg.EventCode == EVT_ID_6006 ) {
        count6006 ++;
        wmiDateObject.Value = objEventReg.TimeGenerated;
        e6006 = wmiDateObject.GetVarDate( true );
        msgTime6006 = new Date( e6006 );
        downStampExp = msgTime6006.getTime();
        msgTime6006 = convertUTCtoDateString( timeToUTCString( msgTime6006 ) );
        msgExpectedDowntime = Math.floor( ( upStampExp - downStampExp ) / MINUTES );
    } else if ( ( objEventReg.EventCode == EVT_ID_1074 ) &&
      ( ( objEventReg.SourceName == "USER32" ) ||
      ( objEventReg.SourceName == "User32" ) ) ) {
        count1074 ++;
        if ( count1074 == count6006 ) {
          msg1074 = objEventReg.Message.replace( /(\r\n|\n|\r)/gm, " " );
          msgUsr1074 = objEventReg.User;
          commentUsr1074 = msgUsr1074;
          msg2nd1074 = "";
        } else {
            msg2nd1074 = ( "\r \t \t \t \t \t" +
              objEventReg.User + "\t" +
              "Message from the Windows System Log [Event ID 1074]: " + 
              objEventReg.Message.replace( /(\r\n|\n|\r)/gm, " " ) );
            count1074 --;
        }
        comment1074 = "Message from the Windows System Log [Event ID 1074]: " + msg1074 + msg2nd1074;
    } else {}
  }
  if ( ( upStampExp - downStampExp > 0 ) &&
  ( upStampExp - downStampExp != expectedDowntime ) &&
  ( enumEvtReg.atEnd() ) &&
  ( logged1074.Count > 0 ) &&
  ( msgExpectedDowntime > DOWNTIME_INTERVAL ) ) {
    downtimeCount = 1;
    msgWarning = createMsgWarning( msgExpectedDowntime );
    comment1074 = msgWarning + comment1074;
    expectedDowntime = upStampExp - downStampExp;
    writeLog( msgTime6006, msgTime6005, msgExpectedDowntime, commentUsr1074, comment1074 );
  }
}

if ( downtimeCount === 0 ) {
  commentAtNoEvent = ( msgNoLogs == "" ) ? ( msgNoReboots + msgShortDowntime ) : msgNoLogs;
  writeLog( "N/A", "N/A", "N/A", "N/A", commentAtNoEvent );
}

//
// Functions declarations
//

// returns JS date by decreasing given date for given number of DAYS
function calculateTime( givenDate, backDays ) {
  return new Date( givenDate.getTime() + backDays * 24 * 60 * 60 * 1000 );
}

// Converting JS date to UTC/WMI compatible format [20150115000000.000000-000]
function dateToUTCString( dt ) {
  var year = dt.getFullYear(),
    month = dt.getMonth() + 1,
    day = dt.getDate(),
    targetDate = year;
  day = day.toString();
  month = month.toString();
  if ( month.length === 1 ) {
    month = "0" + month;
  }
  targetDate = targetDate + month;
  if ( day.length === 1 ) {
    day = "0" + day;
  }
  targetDate = targetDate + day + "000000.000000" + "-000"
  return targetDate;
}

// Converting JS time to UTC/WMI compatible format [20150115140914.000000-000]
function timeToUTCString( dt ) {
  var year = dt.getFullYear(),
    month = dt.getMonth() + 1,
    day = dt.getDate(),
    h = dt.getHours(),
    m = dt.getMinutes(),
    s = dt.getSeconds(),
    targetDate = year;
  month = month.toString();
  day = day.toString();
  h = h.toString();
  m = m.toString();
  s = s.toString();

  if ( month.length === 1 ) {
    month = "0" + month;
  }

  targetDate = targetDate + month;
  if ( day.length === 1 ) {
    day = "0" + day;
  }
  if ( h.length === 1 ) {
    h = "0" + h;
  }
  if ( m.length === 1 ) {
    m = "0" + m;
  }
  if ( s.length === 1 ) {
    s = "0" + s;
  }
  targetDate = targetDate + day + h + m + s + ".000000" + "-000";
  return targetDate;
}

// converts UTC/WMI date to human friendly format '13/04/2014'
function convertUTCtoDateString( wmiDate ) {
  var outputDate;
  if ( !wmiDate) {
    return "null date";
  }
  if ( wmiDate.substr( 6, 1 ) == 0 ) {
    outputDate = wmiDate.substr( 7, 1 ) + "/";
  } else {
      outputDate = wmiDate.substr( 6, 2 ) + "/";
  }
  if ( wmiDate.substr( 4, 1 ) == 0 ) {
    outputDate = outputDate + wmiDate.substr( 5, 1 ) + "/";
  } else {
    outputDate = outputDate + wmiDate.substr( 4, 2 ) + "/";
  }
  outputDate = outputDate +
    wmiDate.substr( 0, 4 ) + " " +
    wmiDate.substr( 8, 2 ) + ":" +
    wmiDate.substr( 10, 2 ) + ":" +
    wmiDate.substr( 12, 2 );
  return ( outputDate );
}

// converts UTC/WMI date to valid JS Date object input:
// new Date('8/24/2014 14:52:10') or (not in this case)
// new Date(2014, 7, 24, 14, 52, 10);
function convertUTCtoJSDateInput( wmiDate ) {
  var outputDate;
  if ( !wmiDate ) {
    return "null date";
  }
  if ( wmiDate.substr( 4, 1 ) == 0 ) {
    outputDate = wmiDate.substr( 5, 1 ) + "/";
  } else {
    outputDate = wmiDate.substr( 4, 2 ) + "/";
  }
  if ( wmiDate.substr( 6, 1 ) == 0 ) {
    outputDate = outputDate + wmiDate.substr( 7, 1 ) + "/";
  } else {
    outputDate = outputDate + wmiDate.substr( 6, 2 ) + "/";
  }
  outputDate = outputDate + 
    wmiDate.substr( 0, 4 ) + " " +
    wmiDate.substr( 8, 2 ) + ":" +
    wmiDate.substr( 10, 2 ) + ":" +
    wmiDate.substr( 12, 2 );
  return ( outputDate );
}

// Writes "tab" separated entry in line depending on given parameters
function writeLog( systemDown, systemUp, downtime, user, comments ) {
  if ( !eventsCount ) {
    // When this is the first entry
    eventsCount = eventsCount + 1;
    return WScript.Echo( startDateOut + "\t" +
      COMPUTER_NAME + "\t" +
      systemDown + "\t" +
      systemUp + "\t" +
      downtime + "\t" +
      user + "\t" + 
      comments );
  } else {
      // When this is a second or next lines
      return WScript.Echo( "\t" + "\t" +
        systemDown + "\t" +
        systemUp + "\t" +
        downtime + "\t" +
        user + "\t" +
        comments );
  }
}

// Creates a warning message when given downtime is above established time 
// (e.g. when downtime is above 720 min (12h))
function createMsgWarning ( downtimeStamp ) {
  if ( downtimeStamp > DOWNTIME_WARNING ) {
    return ( "[NOTE]: Downtime took more than " + Math.floor((downtimeStamp / 60)) +
      " hours, if you suspect data inconsistency please check system manually. " )
  } else {
      return ""
  }
}

// Seeking for date and time in text string and outputs JS Date when appropriate match is found
function findDate( eventMsg ) {
  var date1 = /(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})/,
      date2 = /\W(\d{1,2})[/.-]\W(\d{1,2})[/.-]\W(\d{4})/,
      time = /(\d{1,2}):(\d{2}):(\d{2})/,
      timePM = /(\d{1,2}):(\d{2}):(\d{2})\s(PM|pm)/,
      timeAM = /(\d{1,2}):(\d{2}):(\d{2})\s(AM|am)/,
      matchTime = time.exec( eventMsg ),
      matchTimePM = timePM.exec( eventMsg) ,
      matchTimeAM = timeAM.exec( eventMsg ),
      matchDate1 = date1.exec( eventMsg ),
      matchDate2 = date2.exec( eventMsg );
  if ( matchTimePM !== null ) {
    if ( ( matchDate1 === null ) && ( matchDate2 !== null ) ) {
      return new Date( Number( matchDate2[3] ),
        // JS Month starts from 0
        Number( matchDate2[1] ) - 1,
        Number( matchDate2[2] ),
        // 12h time [02:00:00] PM converts to 24h [14:00:00]
        Number( matchTime[1] ) + 12,
        Number( matchTime[2] ),
        Number( matchTime[3] ) );
    } else if ( ( matchDate1 !== null ) && ( matchDate2 === null ) ) {
        return new Date( Number( matchDate1[3]),
          Number( matchDate1[1] ) - 1,
          Number( matchDate1[2] ),
          Number( matchTime[1] ) + 12,
          Number( matchTime[2] ),
          Number( matchTime[3] ) );
    } else if ( ( matchDate2 === null ) && ( match_dateEN2 === null ) ) {
        return false;
    } else {}
  } else if ( matchTimeAM !== null ) {
      if ( ( matchDate1 === null ) && ( matchDate2 !== null ) ) {
        return new Date( Number( matchDate2[3] ),
          Number( matchDate2[1] ) - 1,
          Number( matchDate2[2] ),
          Number( matchTime[1] ),
          Number( matchTime[2] ),
          Number( matchTime[3] ) );
      } else if ( ( matchDate1 !== null ) && ( matchDate2 === null ) ) {
          return new Date( Number( matchDate1[3] ),
            Number( matchDate1[1] ) - 1,
            Number( matchDate1[2] ),
            Number( matchTime[1] ),
            Number( matchTime[2] ),
            Number( matchTime[3] ) );
      } else if ( ( matchDate2 === null ) && ( match_dateEN2 === null ) ) {
          return false;
      } else {}
  } else if ( matchTime !== null ) {
      if ( ( matchDate1 === null ) && ( matchDate2 !== null ) ) {
        return new Date( Number( matchDate2[3] ),
          Number( matchDate2[2] ) - 1,
          Number( matchDate2[1] ),
          Number( matchTime[1] ),
          Number( matchTime[2] ),
          Number( matchTime[3] ) );
      } else if ( ( matchDate1 !== null ) && ( matchDate2 === null ) ) {
          return new Date( Number( matchDate1[3] ),
            Number( matchDate1[2] ) - 1,
            Number( matchDate1[1] ),
            Number( matchTime[1] ),
            Number( matchTime[2] ),
            Number( matchTime[3] ) );
      } else if ( ( matchDate2 === null ) && ( match_dateEN2 === null ) ) {
          return false;
      } else {
          return new Date( Number( matchDate2[3] ),
            Number( matchDate2[2] ) - 1,
            Number( matchDate2[1] ),
            Number( matchTime[1] ),
            Number( matchTime[2] ),
            Number( matchTime[3] ) );
      } 
  } else
      return false;
}
