@echo off
c:\Windows\System32\cmd.exe /q /c CustomerOpenAccountsReport.pl CompleteOpenAccounts.csv InsuranceExpirationReport.csv
set yy=%date:~-4%
set dd=%date:~-7,2%
set mm=%date:~-10,2%
set MYDATE=%yy%_%mm%_%dd%
echo %MYDATE%
start chrome %CD%\%MYDATE%\OpenAcctOverview_%MYDATE%.html




