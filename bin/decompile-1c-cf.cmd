rem ���� � diff-1c-cf ��� ᤥ���� �����.
rem �᫨ svn ⥪�騩 ���� �⠢�� ⠬ ��� bat, � bzr �६����. 
echo %~dp0
echo %CD%\%1
echo %CD%\%2
wscript.exe G:\repos\git\v83unpack\bin\decompile-1c-cf.js %CD%\%1 %CD%\%2 %3