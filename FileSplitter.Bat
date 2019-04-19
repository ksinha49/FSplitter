@echo off
REM *****************************************************************************
:INTIAL_SETUP
REM *****************************************************************************
setlocal ENABLEDELAYEDEXPANSION
set currDIR=%cd%
set TargetFolder=%CD%\DATA
set ProcessFolder=%CD%\TEMP
set ResultFolder=%CD%\OUTPUT

IF NOT EXIST %ProcessFolder% MKDIR %ProcessFolder%
IF NOT EXIST %ResultFolder% MKDIR %ResultFolder%
DEL /Q %ProcessFolder%\*
DEL /Q %ResultFolder%\*

SET /A P_NUM=0

REM *Edit this value to change the number of lines per file.Also change in FileSplitter.vbs*
SET LPF=160000

REM ****************************************************************************

:LOAD_DISPLAY_WINDOW

title Sircon Conversion Files Formatter
color 0a
set load=
set /a loadnum=0
cls
set "header= STARTING FILE CONVERSIONS..."
set "trailer= PROCESS 1 STARTED"
CALL :LOADING


REM ****************************************************************************
:MAIN_PROCESS_STEP

REM * ADD NEW FILETYPES FOR PROCESSING IN HERE *
REM * Edit this value to change the name of each short file. It will be followed by a number indicating where it is in the list. *

:LICENSE

SET FILETYPE=License_
CALL :PROCESS_F2F_JOB
DEL /Q %ProcessFolder%\*



:LICLOA

SET FILETYPE=LicenseLOA_
CALL :PROCESS_F2F_JOB
DEL /Q %ProcessFolder%\*


:APPTLOA

SET FILETYPE=ApptLOA_
CALL :PROCESS_F2F_JOB
DEL /Q %ProcessFolder%\*


GOTO :EOJ


REM ****************************************************************************


:PROCESS_F2F_JOB

cd %currDIR%
set /a loadnum=%loadnum% +1
SET /A P_NUM=%P_NUM% +1
cls
set "header=        PROCESS %P_NUM% [Changing xlsx to csv %FILETYPE% Files]....... "  
CALL :LOADING
CALL :FILE_SP1


set /a loadnum=%loadnum% +1 
SET /A P_NUM=%P_NUM% +1
cls
set "header=        PROCESS %P_NUM% [Splitting  %FILETYPE% Files]....."
CALL :LOADING
CALL :FILE_SP2

set /a loadnum=%loadnum% +1 
SET /A P_NUM=%P_NUM% +1
cls
set "header=        PROCESS %P_NUM% [Formatting %FILETYPE% Files]....."
CALL :LOADING
CALL :FILE_SP3
GOTO :END

REM ****************************************************************************

:LOADING
set /a loadloop=0
set load=
:loop_bar
if !loadloop! leq !loadnum! (
set load=%load%лл
set /a loadloop=%loadloop% +1
goto :loop_bar
)
echo.
echo !header!
echo                   SPLIT count = %LPF% LINES/FILE
echo -------------------------------------------------------------------------
echo !load!
echo -------------------------------------------------------------------------
IF %P_NUM% EQU 1 (
echo !trailer!
)else (
set /a disp_num= %P_NUM% -1
echo PROCESS !disp_num! !trailer!
)
if !loadnum! gtr 36 set /a loadnum=1
GOTO :END

REM ****************************************************************************


:FILE_SP1

SET NEWFILIND=Y
CALL :GetstartTime
COPY /-Y %TargetFolder%\*%FILETYPE%* %ProcessFolder%

@ECHO CONVERTING....
CALL :xlsTocsvConverter 
CALL :GetendTime
CALL :GetTimeElapsed
SET "trailer= STEP PROCESSING TIME %Process_time% "
GOTO :END


:FILE_SP2
@ECHO PROCESSING:
for %%f in (%ProcessFolder%\*) do (
SET FILESEL=%%~nf
SET BFN=%FILESEL%.csv
@ECHO !BFN!
)
sleep 15
CALL :FILESPLITTER
GOTO :END

:FILE_SP3
CALL :GetstartTime
DEL /Q %ProcessFolder%\*

for %%g in (%ResultFolder%\%FILETYPE%*) do (
SET FILESEL1=%%~ng
SET "BFN1=!FILESEL1!.csv"
CALL :FILEFORMATTER
)
CALL :GetendTime
CALL :GetTimeElapsed
SET "trailer= STEP PROCESSING TIME %Process_time% "
CLS
SET /A P_NUM=%P_NUM% +1
call :LOADING

GOTO :END

REM ******************************************************************************

:FILESPLITTER

SET SFN=%FILETYPE%
SET SFX=%BFN:~-3%
CALL :GetstartTime
CALL :PrintJob
@ECHO MERGING ALL INPUT FILES
COPY %ProcessFolder%\*.CSV  %ProcessFolder%\%SFN%.%SFX%
@ECHO OUTPUT IN FOLDER %ResultFolder%
CALL :PrintJob
@ECHO SPLITING IN SELECTED COUNT %LPF%
set "inputfile=%ProcessFolder%\%SFN%.%SFX%"
set "outputfile=%ResultFolder%\%SFN%" 
call Filesplitter.vbs %inputfile% %outputfile% 
CALL :GetendTime
CALL :GetTimeElapsed
SET "trailer= STEP PROCESSING TIME %Process_time% "

GOTO :END


REM ****************************************************************************
:FILEFORMATTER

IF %FILETYPE% == License_ (
for /f "usebackq tokens=1-10 delims=," %%A in (%ResultFolder%\!BFN1!) do (
IF NOT "%%B" == "" (
echo %%A,%%B,%%C,%%D,%%E,%%F,%%G,%%H
)
)>> %ProcessFolder%\!BFN1!
)

IF %FILETYPE% == LicenseLOA_ (
for /f "usebackq tokens=1-10 delims=," %%A in (%ResultFolder%\!BFN1!) do (    
if "%%E" == "8" (
    echo %%A,%%B,%%C,%%D,%%E,A,%%F
)ELSE IF NOT "%%B" == "" (
echo %%A,%%B,%%C,%%D,%%E,I,%%F
) 
)>> %ProcessFolder%\!BFN1!
)

IF %FILETYPE% == ApptLOA_ (
for /f "usebackq tokens=1-10 delims=," %%A in (%ResultFolder%\!BFN1!) do (  
if "%%F" == "1" (
echo %%A,%%B,%%C,%%D,%%E,%%F,A,%%G
)ELSE IF NOT "%%B" == "" (
echo %%A,%%B,%%C,%%D,%%E,%%F,I,%%G
)
)>> %ProcessFolder%\!BFN1!
)
copy  %ProcessFolder%\!BFN1! %ResultFolder%\!BFN1! /b/v/y
GOTO :END

REM ****************************************************************************
:xlsTocsvConverter 
for %%f in (%ProcessFolder%\*.xls*) do (
CALL xlstocsv.vbs %%f %ProcessFolder%\%%~nf.csv
)
DEL %ProcessFolder%\*.xls*
GOTO :END

REM ****************************************************************************

:GetstartTime
for /F "tokens=1-4 delims=:.," %%a in ("%time%") do (
   set /A "start=(((%%a*60)+1%%b %% 100)*60+1%%c %% 100)*100+1%%d %% 100"
)
GOTO :END


:GetendTime
for /F "tokens=1-4 delims=:.," %%a in ("%time%") do (
   set /A "end=(((%%a*60)+1%%b %% 100)*60+1%%c %% 100)*100+1%%d %% 100"
)
GOTO :END

:GetTimeElapsed
set /A elapsed=end-start
set /A hh=elapsed/(60*60*100), rest=elapsed%%(60*60*100), mm=rest/(60*100), rest%%=60*100, ss=rest/100, cc=rest%%100
if %mm% lss 10 set mm=0%mm%
if %ss% lss 10 set ss=0%ss%
if %cc% lss 10 set cc=0%cc%
SET Process_time=%hh%:%mm%:%ss%
GOTO :END


REM ****************************************************************************


:PrintJob
set /a loadnum=%loadnum% +1
cls
CALL :LOADING

GOTO :END




:EOJ

set /a loadnum=36
cls
CALL :LOADING
@echo  JOB COMPLETE PRESS ANY KEY TO EXIT
pause>NUL
exit



:END