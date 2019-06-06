rmdir Support /s /Q
del *.exe
del *.mdb
del Server.zip
mkdir Support

copy ..\..\Loader\Loader.exe .
copy ..\..\Minder\Minder.exe .
copy ..\..\Ware\MWare.exe .
copy ..\..\Reps\MReps.exe .
copy ..\..\Client\MMOS.exe .
copy ..\..\Admin\MAdmin.exe .
rem copy ..\..\Databases\Local.mdb . don't copy  - as these need to added to the configure app path so tables can be re-attached
copy ..\..\Databases\Central.mdb .
copy ..\..\Databases\CentralTest.mdb .
copy ..\..\Client\Package\Support\*.ocx Support\.
copy ..\..\Client\Package\Support\*.dll Support\.
copy ..\..\Client\Package\Support\*.tlb Support\.

"c:\Program Files\7-Zip\7z.exe" a Server.zip -xr!*.bat -xr!.gitignore -xr!*.txt

pause