@echo off
:input
title random_name-bat
::bat标题
:1
:a
echo 输入s开始，c关闭
set /p var=请输入:
if %var%==s goto a1
if %var%==c goto a2
echo 输入超出预定值,请重新输入&pause&cls&goto a
:a1
for %%i in (%0) do set aa=%%~dpi 
cd /d %aa%
echo START，若需重新执行请按Enter


@echo off
for /f "tokens=1* delims=:" %%a in ('findstr /n . list.txt') do (
   set .%%a=%%b&set /a n=%%a
)
:loop
   set /a w=%random%%%n+1
   if not defined .%w% goto loop
   call echo %%.%w%%%
   pause>nul



)
if "%num%"=="1" cls & goto 1 
goto loop

pause>nul
cls & goto input
:a2
pause
