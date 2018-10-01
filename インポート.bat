@set /p answer="é¿çsÇµÇ‹Ç∑Ç©(y/n)ÅH" : %answer%
@if "%answer%" neq "y" exit

cscript vbac.wsf combine
