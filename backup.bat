@echo off
setlocal EnableDelayedExpansion
rem 调用参数: %1->(ocl模式)数据库用户/密码@SID %2->目标路径(不带'\'结尾) %3->目标文件名 %4->执行日期 %5->ocl
rem 调用参数: %1->(file模式)备份全路径文件名 %2->目标路径(不带'\'结尾) %3->目标文件名 %4->执行日期 %5->file

rem 压缩密码不能使用: !\
set zip_pwd="这里写压缩密码"

set date1=%date:~0,10%

if "%5" == "ocl" (
	echo +++begin %date% %time% : backup and dmp %3 to %2 ^<br^> >> %2\backup_%4_log.htm
	echo ---oracle %0 %1 %2 %3 %4 %5 >> %2\backup_%4_log.htm
	echo ---exp %1 file=%2\%3_%4.dmp log=%2\%3_%4.log ^<br^> >>  %2\backup_%4_log.htm
	exp %1 file=%2\%3_%4.dmp log=%2\%3_%4.log
	echo +++begin %date% %time% : Rar %3 to %2 ^<br^> >> %2\backup_%4_log.htm
	rem echo c:\progra~1\winrar\rar a -ep -df -mt3 -m5 -pzkwt-123 %2\%3_%4.rar %2\%3*.* 
	c:\progra~1\winrar\rar a -ep -df -mt3 -m5 -p%zip_pwd% %2\%3_%4.rar %2\%3*.*
)

if "%5" == "file" (
rem	echo begin  %date% %time% : backup and copy %3 to %2 ^<br^> >> %2\backup_%4_log.htm
	echo +++begin %date% %time% : Rar %3 to %2 ^<br^> >> %2\backup_%4_log.htm
	echo ---file %0 %1 %2 %3 %4 %5 >> %2\backup_%4_log.htm
	rem echo c:\progra~1\winrar\rar a -ep -mt3 -m5 -pzkwt-123 %2\%3_%4.rar %1
	c:\progra~1\winrar\rar a -ep -mt3 -m5 -p%zip_pwd% %2\%3_%4.rar %1
)

echo "" > %2\%3_%4.rar.ok
echo ###End %date% %time% : backup %2\%3 ^<br^> >> %2\backup_%4_log.htm
exit
