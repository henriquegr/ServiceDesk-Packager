%~dp0\grep.exe -r "^TABLE" %~dp0\..\Userload\*.* | %~dp0\sed "s/.*:TABLE//g" | %~dp0\\usort -u > %~dp0\..\bkp\lst_bkp_tab.txt  