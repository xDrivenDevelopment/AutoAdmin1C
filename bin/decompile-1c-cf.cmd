rem путь к diff-1c-cf луше сделать полным.
rem если svn текущий путь ставит там где bat, то bzr временный. 
echo %~dp0
echo %CD%\%1
echo %CD%\%2
wscript.exe G:\repos\git\v83unpack\bin\decompile-1c-cf.js %CD%\%1 %CD%\%2 %3