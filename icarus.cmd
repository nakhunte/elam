
@echo off
setlocal EnableDelayedExpansion

:: Name: ICARUS.cmd
:: Author: Mario Espinoza
:: Created: Feb 12, 2017
:: Version: 0.1.3 (2017.02.21)

:: OS Compatibility: Windows {Server 2003-2016, XP, Vista, 7, 8, 10}
:: Prerequisites: None (the OS platforms listed above have all required 
::                utilities pre-installed by default)

:: Summary:
:: ICARUS (Internet Connection Analysis Report Uploading System)
:: Using the native command line utilities, this batch script will run many 
:: internet diagnostic tools simultaneously, show the running status and
:: results of the tests, then prompts the user to submit the report to a server

:: Initialize the main variables
call :InitVariables

:: Check if a parameter was passed and run the test-routine specified
call :TaskRunner %1

:: Begin event logging
call :LogEntry "." "Start Of Event Log"

:: Make it so, number 1
goto MainProc




:::::::::: CONFIGURATION VARIABLES ::::::::::::::::::::::::::::::::::::::::::::
:InitVariables
call :LogEntry "." "Initializing Variables"

:: A comma separated list of commands required by this script (no spaces)
set CmdApps=findstr,ping,tracertt,pathping,ipconfig,netstat,cscript

:: Output File Names
set "ReportName=icarus-report.html" %= Live status and results of tests =%
set "LogName=icarus.log" %= verbose reporting of actions and restults =%

:: Date and Time Variable
call :LogEntry "." "Setting the Sart Timestamp for the Report"
for /F "delims=" %%a IN ('date /t') DO set myDate=%%a
for /f "tokens=5 delims=. " %%a in ('echo. ^| time') do (set myTime=%%a)
set myDate=%mydate:~10,4%-%mydate:~4,2%-%mydate:~7,2%
set "ReportTimeStamp=%myDate% @ %myTime%"

:: Others
set "bang=^!^!^!" %= this makes life easier because ! is a special charecter =%
				  %= must use setlocal EnableDelayedExpansion and !bang! variable format =%
set "PingReps=-4 -n 16"
goto:eof



:::::::::: MAIN PROCEDURE :::::::::::::::::::::::::::::::::::::::::::::::::::::
:MinProc
:: Check availability of commands required by this script
call :TestCmdApps

:: Read the script, find the test routines, and add them to a variable array
:: *Note: Test-Routines are denoted by Labels with a leading underscore ":_SubRoutine"
set "count=0"
for /f "delims=:_ tokens=1" %%a in ('findstr /i /r /c:"^\:_[a-z]" %0') do (
    set /a count+=1 > nul
	set "Label[!count!]=%%a"
)

:: Make a log entry of what was found
call :LogEntry "." "Found %count% test procedures in this script"
set "mylist= "
for /l %%n in (1 1 %count%) do (set myList=!myList! !Label[%%n]!)
call :LogEntry "." "The procedures I found are: !myList!"
call :LogEntry "." "Running the connection tests now"

:: Run the connection test routines in the variable array created above
for /L %%n in (1 1 !count!) DO (
	call :LogEntry "." "Performing the !Label[%%n]! procedure"
	start "!Label[%%n]!" /min %0 !Label[%%n]!
	for /f "tokens=2" %%a in ('tasklist /v ^| findstr /i /c:"!Label[%%n]!" ') do (
			set PIDs[%%n]=%%a
			call :LogEntry "." "The process ID for this command is: %%a"
	)
)

:: Wait a couple of seconds before starting the test restults page
ping -n 3 localhost > nul 2>&1

::Create the test report and monitor status of connection tests
call :ICARUS_Report %count%
start %ReportName%

:: Check the status of the connection tests every 3 seconds and update the report 
:WaitLoop
for /l %%n in (1 1 !count!) do (
	ping -n 3 localhost > nul 2>&1
	set FoundPid=
	call :LogEntry "." "Getting the status of Proc ID: !PIDs[%%n]!"
	for /f "tokens=2" %%a in ('tasklist /nh /fi "pid eq !PIDs[%%n]!" 2^>nul') do (
		set "FoundPid=true"
	)
	if defined FoundPid (
		call :LogEntry "." "The !Label[%%n]! procedure is still running"
		call :ICARUS_Report %count%
	)
)
if defined FoundPid goto WaitLoop

:: Update the report one more time and exit the script
call :LogEntry "." "Connection tests finished"
call :ICARUS_Report %count%

:EOP
call :LogEntry "." "Setting the Finish Timestamp for the Report"
call :LogEntry "." "End of Log"
pause
goto:eof

:: Delete the temp files created during connection tests
:: %= del /q *.txt *.done > nul 2>&1 =%



:::::::::: SUB-PROCEDURES :::::::::::::::::::::::::::::::::::::::::::::::::::::

:_GetPublicIP
:: Writ a vbs scrit to get the contents of a webpage showing the public ip address
:: Save the public ip address to the registry, then echo it out
call :MakeVBS > wanip.vbs
set "LineCount=0"
for /f "skip=3 delims=" %%a in ('cscript wanip.vbs') do (
	reg add HKCU\Environment /t REG_EXPAND_SZ /v WANIP /d "%%a"
	echo %%a
)

:: To read the WANIP address from the registry, use this command
:: for /f "skip=3 tokens=3" %%a in ('reg query HKCU\Environment /v WANIP') do (set wanip=%%a)
del /q wanip.vbs
goto:eof


:_Ping2LocalHost	
:: This will verify that that the TCP/IP stack is functioning properly
ping %~1 localhost
goto:eof


:_Ping2LanRouter
:: This will ping the local router (default gateway) and determine if packet loss is occurring on both the LAN
ipconfig /all > ipconfig.txt 2>&1
for /f "tokens=13" %%a in ('ipconfig ^| findstr /i "gateway"') do (set "DG=%%a")
ping %~1 %DG%
goto:eof


_:Ping2GWC1
:: GWC1 is physically located in Phoenix, AZ, as of 2017-02-21
set Gwc1=gwc1.onthenetOffice.com
set "count=0"
for /F "delims=" %%f in ('ping %PingReps% %Gwc1%') do (
    set /a count+=1 > nul
    set "Ping2GWC1[!count!]=%%f"
)
for /L %%n in (1 1 !count!) DO (
	echo !Ping2GWC1[%%n]!
)
goto:eof


_:Ping2GWC2
:: GWC2 is physically located in Ashburn, VA, as of 2017-02-21
set Gwc2=gwc2.onthenetOffice.com
set "count=0"
for /F "delims=" %%f in ('ping %PingReps% %Gwc2%') do (
    set /a count+=1 > nul
    set "Ping2GWC2[!count!]=%%f"
)
for /L %%n in (1 1 !count!) DO (
	echo !Ping2GWC2[%%n]!
)
goto:eof


_:Ping2GWC5
:: GWC5 is physically located in Irvine, CA, as of 2017-02-21
set Gwc5=gwc5.onthenetOffice.com
set "count=0"
for /F "delims=" %%f in ('ping %PingReps% %Gwc5%') do (
    set /a count+=1 > nul
    set "Ping2GWC5[!count!]=%%f"
)
for /L %%n in (1 1 !count!) DO (
	echo !Ping2GWC5[%%n]!
)
goto:eof


_:Ping2GoogleDNS
:: The world famouse public DNS node, courtesy of Google
set GoogleDNS=8.8.4.4
set "count=0"
for /F "delims=" %%f in ('start /min ping %PingReps% %GoogleDNS%') do (
    set /a count+=1 > nul
    set "Ping2GoogleDNS[!count!]=%%f"
)
for /L %%n in (1 1 !count!) DO (
	echo !Ping2GoogleDNS[%%n]!
)
goto:eof


_:Ping2Google
:: Testing the connection to this giant internet service will prove if the connection issue is only with us or to the world
set Google=play.google.com
set "count=0"
for /F "delims=" %%f in ('ping %PingReps% %Google%') do (
    set /a count+=1 > nul
    set "Ping2Google[!count!]=%%f"
)
for /L %%n in (1 1 !count!) DO (
	echo !Ping2Google[%%n]!
)
goto:eof


_:Ping2Amazon
:: Testing the connection to this giant internet service will prove if the connection issue is only with us or to the world
set Amazon=aws.amazon.com
set "count=0"
for /F "delims=" %%f in ('ping %PingReps% %Amazon%') do (
    set /a count+=1 > nul
    set "Ping2Amazon[!count!]=%%f"
)
for /L %%n in (1 1 !count!) DO (
	echo !Ping2Amazon[%%n]!
)
goto:eof


:MakeVBS
:: Creates a VBS script that uses the Windows built-in API to read HTML from any webpage
set lf=^


echo Option Explicit!lf!^
Dim http : Set http = CreateObject( "MSXML2.ServerXmlHttp" )!lf!^
'Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")!lf!^
'Dim txtfile : Set txtFile = objFSO.CreateTextFile("public-ip.txt",2,true)!lf!^
http.Open "GET", "http://icanhazip.com", False!lf!^
http.Send!lf!^
Wscript.Echo http.responseText!lf!^
'txtFile.Write http.responseText!lf!^
'txtFile.Close!lf!^
'Set objFSO = Nothing!lf!^
'Set txtFile = Nothing!lf!^
Set http = Nothing
goto:eof


:TestCmdApps
:: Run each command in the "CmdApps" list, read the "ErrorLevel", and determine if it's available
call :LogEntry "." "Testing each required command for availability"
for %%a in (%cmdApps%) do (
	%%a /? > nul 2>&1
	if !ErrorLevel! GTR 1 (
		set "%%a=false"
		if %%a==findstring (
			call :LogEntry "err" "The comand '%%a' was not accessible and is required by this script"
			call :LogEntry "err" "Sorry but it's not gonna work out. The script is terminated. :-("
			goto:eof
		) else (
			call :LogEntry "err" "The command '%%a' is not functioning, so the %%a tests will be skipped"
		)
	) else (
		set "%%a=true"
		call :LogEntry "." "The command '%%a' is available and functioning"
	)
)
goto:eof



:ICARUS_Report
:: Create an HTML page with the status of the connection tests. This will also be used to transmit the results to an email
set NumOfTests=%1
set lf=^


echo ^<^^^!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1//EN" "http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd"^>!lf!^
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en"^>!lf!^
<head^>!lf!^
<meta http-equiv="Content-Type" content="text/html;charset=utf-8" /^>!lf!^
<meta http-equiv="refresh" content="4" /^>!lf!^
<title^>ONDU REPORT^</title^>!lf!^
<style type="text/css"^>!lf!^
Body {margin:10px auto 10px auto; width: 800px; color:white; font:normal normal normal 14pt/1em "Arial Black",sans-serif; background-color:#0E4254;}!lf!^
H2 {color:orange; font:italic small-caps normal 18pt/1em "Arial Black",sans-serif;}!lf!^
HR {margin:0; padding:0; border:thin dashed black;}!lf!^
.TestLabel {margin-left:10px; padding:0px; font:italic normal normal 14pt/1em "Arial",sans-serif; text-align:left;}!lf!^
.console { overflow: scroll; height: 140px; margin:10px; padding:10px; font:normal normal normal 12pt/1em "Courier New",monospace; border:thin solid black; background-color:white; color:black; text-align:left;}!lf!^
pre {!lf!^
    white-space: pre-wrap;       /* Since CSS 2.1 */!lf!^
    white-space: -moz-pre-wrap;  /* Mozilla, since 1999 */!lf!^
    white-space: -pre-wrap;      /* Opera 4-6 */!lf!^
    white-space: -o-pre-wrap;    /* Opera 7 */!lf!^
    word-wrap: break-word;       /* Internet Explorer 5.5+ */!lf!^
}!lf!^
</style^>!lf!^
</head^>!lf!^
<body^>!lf!^
<div^>^<img src="https://my.onthenetoffice.com/images/logo.png" alt="onthenetOffice: the fresh maker^!" title="onthenetOffice: the fresh maker^!" /^>^</div^>!lf!^
<h2^> OTNO Network Diagnostics Utility^</h2^>!lf!^
<h4^>Start Time: %timeStamp% ^</h4^>

:: Get the running results from all connection tests
for /L %%n in (1 1 !NumOfTests!) DO (
  call :GrabTestResults !Label[%%n]!
)

echo ^</body^>!lf!^
</html^>

goto:eof


:GrabTestResults
set ResultFile=%1.txt
echo.
echo Getting running results from: %ResultFile%
echo.
set/a LineCount=0 > nul
echo ^<div class="TestLabel"^>!lf!^
%TestName% !lf!^
</div^>!lf!^
<pre class="console"^>

if exist %ResultFile% (
	for /F %%a in (%ResultFile%) do (set /a LineCount+=1 > nul)
	set /a start=!LineCount!-5 > nul
	echo The number of lines in %ResultFile% is: !LineCount!
	echo.
	findstr /b Pinging %ResultFile%
	if !LineCount! GTR 4 (
		echo ...
		more +%start% %ResultFile%
	) else (
		type %ResultFile%
	)
)

echo ^</pre^>

goto:eof


:LogEntry
for /F "delims=" %%a IN ('date /t') DO set myDate=%%a
for /f "tokens=5 delims=. " %%a in ('echo. ^| time') do (set myTime=%%a)
set myDate=%mydate:~10,4%-%mydate:~4,2%-%mydate:~7,2%
set "LogTimeStamp=%myDate% %myTime%"
set type=...
set string=%LogTimeStamp% !type! %~2
if %1=="err" (
	set type=^^!^^!^^!
	set string=%LogTimeStamp% !type! %~2
)
echo !string!
goto:eof


:TaskRunner
if [%1] neq [] (
	call :LogEntry "." "The parameter %1 was passed to this script"
	call :LogEntry "." "Now performing the %1 procedure"
	call :_%1 > %1.txt
	call :LogEntry "." "The %1 procedure finished"
	echo Done > %1.done
	exit /b
)
goto:eof


