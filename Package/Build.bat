@echo off

Del /s Configure-Package.zip >nul 2>&1

copy ..\Configure.exe Support\. > nul
copy ..\..\Databases\Reps.mdb Support\. > nul
copy ..\..\Databases\RepsTest.mdb Support\. > nul
copy ..\..\Databases\Local.mdb Support\. > nul
copy ..\..\Databases\LocalTest.mdb Support\. > nul

echo .

echo Now run Build.bat in the ZipBuild folder
echo .

pause

copy ..\ZipBuild\Server.zip Support\. > nul
echo .

echo Now run Configure.bat in the support folder
echo .

pause

"c:\Program Files\7-Zip\7z.exe" a Configure-Package.zip -xr!*.bat -xr!.gitignore -xr!*.txt -x!Support\


pause