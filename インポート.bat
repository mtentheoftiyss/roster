@set /p answer="���s���܂���(y/n)�H" : %answer%
@if "%answer%" neq "y" exit

cscript vbac.wsf combine
