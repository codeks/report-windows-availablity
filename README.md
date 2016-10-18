# report-windows-availablity
## SYNOPSIS
    Script is checking a server for downtimes above 5 min within given period of time (30 DAYS).
## DESCRIPTION
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
## NOTES
    File Name      : availability_3.3.js
    Author         : Damian Danak
    Prerequisite   : Windows Server 2003 and higher.
    Tested         : Target machines with Windows Server 2003 to 2012 R2
    Version        : 3.3.0
    Date           : 12-09-2016
## EXAMPLE
    On Windows console type:
    cscript //nologo availability_3.3.js > availablity_report.log
## EXAMPLE
    On Windows console type:
    cscript //nologo availability_3.3.js 30 5 > availablity_report.log
## HISTORY
    v3.3 09-09-2016 16:00:00 Rewriting script - logic stays mostly the same
    v1.0 11-12-2014 10:50:43 Initial code created
## LICENCE
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
