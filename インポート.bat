@set /p answer="実行しますか(y/n)？" : %answer%
@if "%answer%" neq "y" exit

cscript vbac.wsf combine
