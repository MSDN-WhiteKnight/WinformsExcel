@echo off
cd %~dp0
copy /Y /V ..\bin\Release\WinFormsExcel.dll /B lib\net40\
copy /Y /V ..\bin\Release\WinFormsExcel.pdb /B lib\net40\
copy /Y /V ..\bin\Release\WinFormsExcel.xml lib\net40\
copy /Y /V ..\readMe.txt .\
"C:\Distr\Microsoft\nuget 2.8.6\nuget.exe" pack WinFormsExcel.dll.nuspec
