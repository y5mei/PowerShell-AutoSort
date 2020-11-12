#######################################################################
#	In-Process Failed CMM Reports Sorting Program
#
#	Author:			Chris Mei (Quality Inspector)
#	E-mail:			Chris.Yaowen@linamar.com
#				ESTON MANUFACTURING
#				277 Slivercreek Parkway N,
#				Guelph, ON, anada, N1H 1E6
#	Date Created:		March, 09, 2017
#	Purpose:		Sort all the in-process failed CMM reports
#				into the correct "share" folder
#	Tested OS:		WINDOWS 7; 
#				WINDOWS 10;
#######################################################################
# WHAT CHANGED IN REGISTRY
#	1. ADD A KEY "CLEAN BY TIME" IN REGISTRY LOCAITON:
#		HKEY_CLASSES_ROOT\Directory\shell
#	2. VALUE "Command" WAS CREATED, AND SET TO BE:
#		C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe  -file "D:\Users\Cmmeston\AppData\Roaming\Microsoft\Windows\SendTo\clean.ps1"
#	3. To change the executionpolicy, TYPE THE FOLLOWING IN POWERSHELL:
#		set-executionpolicy -executionpolicy Unrestricted
#	4. CHANGE PERMISSION OF POWERSHELL AT:
#		HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\PowerShell\1\ShellIds\Microsoft.PowerShell
<#
Push-Location
Set-Location Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\PowerShell\1\ShellIds\Microsoft.PowerShell\
Set-ItemProperty .\ -Name "ExecutionPolicy" -Value "Bypass"
Pop-Location
#>
 
 
Write-Host "WELCOME TO THE SORT IN-PROCESS FOLDER PROGRAM"
#SET CURRENT TIME AND GET CURRENT DIR
$now = Get-Date
echo "Current Time is $now"
$currentdir = (Get-Item -Path ".\" -Verbose).FullName
echo "Current DIR is $currentdir"
echo "******************************************************************"
echo "Please input the number of hours that you want to Sort:"
$cleantime = Read-Host
$intcleantime = [int]$cleantime
echo "*******************************************************************"
 

 
Switch ($cleantime)
{
{$intcleantime -gt 60}{Write-Host "The number you input is too big [We can only clean the folder within 30 hours]"; break}
default {
 
 
# Creat a folder to contain all the failed reports
$newfolder ="In-Process-Fail"
New-Item -ItemType Directory -Path ".\$newfolder" -Force | out-null


#Display current working directory 
$currentdir = (Get-Item -Path ".\" -Verbose).FullName
echo "Current DIR is $currentdir"
 
# find all the failed reports 
Get-ChildItem -Path .\ -Filter *f*.pdf -Exclude *P.pdf, *HFV6*P*.pdf, *SHA*P*pdf, *flu*p*pdf, *fo*, *set*, *first*, *ENG* -include *.* -recurse -force|
where {$_.LastWriteTime -gt (date).addhours(-$cleantime)}|
%{Move-Item -Path $_.FullName -Destination ".\$newfolder\" -Force}
 
#COUNT NUMBER OF FAILED FILES exclude of first-off and SET-UP request
$NumInprocessFail = 0
gci -Path ".\$newfolder\" -Filter *.pdf|foreach {$NumInprocessFail++}
 
<#
#COUNT NUMBER OF FAILED FILES
$NumInprocessFail = -1
$myfiles = Get-ChildItem -Path ".\$newfolder\" -Exclude *fo*, *set*, *first*
foreach ($fl in $myfiles){$NumInprocessFail += 1}
#>
#$NumInprocessFail = Get-ChildItem -Path ".\$newfolder\" -file *.pdf | Measure-Object | %{$_.Count}

Write-Host "
*******************************************************************************
                        WITHIN $cleantime HOURS ,
THE TOTAL NUMBER OF FAILED INPROCESS CHECK ARE $NumInprocessFail
*******************************************************************************
"-foreground yellow -nonewline







#CREAT THE STEP2.BAT FILE
$STEP2 = @"
%Get current year%
REG ADD "HKCU\Control Panel\International" /f /v sShortDate /t REG_SZ /d "dd/MM/yyyy" >nul
set thisyear=%date:~6,4%

set rootdir=E:\Google Drive\Auto Sort\Demo\Archive_Folder

%Set the directory constant that where the failed in process reports should go%
set dir_hfv6="%rootdir%\%thisyear%\4. HFV6"
set dir_ka="%rootdir%\%thisyear%\3. KA Adapter"
set dir_ltg="%rootdir%\%thisyear%\2. LTG & 6.2L"
set dir_8speed="%rootdir%\%thisyear%\1. 8 Speed"
set dir_gep="%rootdir%\%thisyear%\5. GEP"
set dir_dmax="%rootdir%\%thisyear%\6. DMAX"
set dir_shaft="%rootdir%\%thisyear%\7. Shaft"


%Cut and Paste HFV6 reports%
xcopy /i /y .\*HFV6*.pdf %dir_hfv6%
del .\*HFV6*.pdf
xcopy /i /y .\*LGX*.pdf %dir_hfv6%
del .\*LGX*.pdf


%Cut and Paste 8SPEED reports%
xcopy /i /y .\*speed*.pdf %dir_8speed%
del .\*speed*.pdf


%Cut and Paste KA reports%
xcopy /i /y .\*ka*.pdf %dir_ka%
del .\*ka*.pdf


%Cut and Paste LTG reports%
xcopy /i /y .\*ltg*.pdf %dir_ltg%
del .\*ltg*.pdf


%Cut and Paste GEP reports%
xcopy /i /y .\*hummer*.pdf %dir_gep%
del .\*hummer*.pdf
xcopy /i /y .\*cyl*.pdf %dir_gep%
del .\*cyl*.pdf
xcopy /i /y .\*gep*.pdf %dir_gep%
del .\*gep*.pdf
xcopy /i /y .\*gen*.pdf %dir_gep%
del .\*gen*.pdf


%Cut and Paste DMAX reports%
xcopy /i /y .\*dmax*.pdf %dir_dmax%
del .\*dmax*.pdf


%Cut and Paste SHAFT reports%
xcopy /i /y .\*shaft*.pdf %dir_shaft%
del .\*shaft*.pdf

%If there exist not proper named pdf reports, stop this program%
@echo off
if exist *.pdf (
echo ***************************************************************
echo ***************************************************************
echo ***************************************************************
echo                 Unidentified reports are found
echo No Production Line Information. Please change the file name
echo E.G. 17-12075 8SPEED 555 F.PDF
echo ***************************************************************
echo ***************************************************************
echo ***************************************************************
Pause
EXIT)
%If everything good, then delete this in process fail folder and this bat file%
set batfile="%CD%"
cd..& rd /s /q %batfile%
Pause

"@
$STEP2 | Out-File ".\$newfolder\Step2.bat" -Encoding ASCII
}
 
}
Write-Host -NoNewLine 'Press any key to continue...';
 
#$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');