::automatically copy msedgedriver (needs to be in same directory) to python directory, 
@echo off
for /D %%J in ("%LOCALAPPDATA%\Programs\Python\Python3*") do set pydir=%%~J
echo python directory found: %pydir%
echo copying webdriver
@echo on
copy %CD%\msedgedriver.exe %pydir%\msedgedriver.exe

@echo off
