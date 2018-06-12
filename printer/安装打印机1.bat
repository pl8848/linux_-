@echo off 
@color 02



@cscript //h:cscript //s

@echo +++++本程序安装打印机到你的电脑，请在后面选择1/2/3/4来安装你要的打印机。+++++

::@cd C:\Windows\System32\Printing_Admin_Scripts\zh-CN
@echo off
@echo ------------------------------------------------
@echo ｜1.Kyocera Mita CS-2550（生产二楼七海区域）   ｜
@echo ｜2.Kyocera Mita CS-2550（生产二楼中间）  　　 ｜
@echo ｜3.Kyocera Mita CS-2550（生产二楼思瑞区域） 　｜
@echo ｜4.Kyocera Mita CS-2550 （办公楼三楼）        ｜
@echo ｜5.Kyocera Mita CS-2550 （办公楼2楼  　   　　｜
@echo ｜6.退　出　　　　　　　　　　 　　　　　　　　｜　
echo ------------------------------------------------
@set /p slection=请选择您要安装的打印机(输入编号即可):   


echo ***************************************
@echo off
@if %slection%==1 goto 1
@if %slection%==2 goto 2  
@if %slection%==3 goto 3
@if %slection%==4 goto 4 
@if %slection%==5 goto 5
@if %slection%==6 goto 6



@goto EXITING

:1
@echo 添加打印端口**********************************************************
@echo off
prnport -a -r 10.135.51.162_9100 -h 10.135.51.162 -o raw
@echo 安装Kyocera Mita CS-2550驱动***************************************
@echo off
rundll32 printui.dll,PrintUIEntry /ia /h x86 /m "Kyocera Mita CS-2550" /f  "prnky004.inf"  /w /q && @echo 驱动成功安装
@echo 安装Kyocera Mita CS-2550打印机*************************************
@echo off
(prnmngr -a -p "Kyocera Mita CS-2550" -m "Kyocera Mita CS-2550" -r 10.135.51.162_9100) && (prncnfg -t -p "Kyocera Mita CS-2550" -l "生产二楼七海区域" -m "故障排除请联系IT")
@echo off
pause
@echo Kyocera Mita CS-2550正在打印测试页************************
prnqctl -p "Kyocera Mita CS-2550" -e && @echo 打印测试页发送成功！
pause
cls

EXIT  



:2
@echo 添加打印端口**********************************************************
@echo off
prnport -a -r 10.135.51.160_9100 -h 10.135.51.160 -o raw
@echo 安装Kyocera Mita CS-2550驱动***************************************
@echo off
rundll32 printui.dll,PrintUIEntry /ia /h x86 /m "Kyocera Mita CS-2550" /f  "prnky004.inf"  /w /q && @echo 驱动成功安装
@echo 安装HP LaserJet 2420 PCL 5e打印机*************************************
@echo off
(prnmngr -a -p "Kyocera Mita CS-2550" -m "Kyocera Mita CS-2550" -r 10.135.51.160_9100) && (prncnfg -t -p "Kyocera Mita CS-2550" -l "生产大楼二楼中间区域" -m "故障排除请联系IT")
@echo off
pause
@echo Kyocera Mita CS-2550正在打印测试页************************
prnqctl -p "Kyocera Mita CS-2550" -e && @echo 打印测试页发送成功！
pause
cls
EXIT

:3
@echo 添加打印端口**********************************************************
@echo off
prnport -a -r 10.135.51.161_9100 -h 10.135.51.161 -o raw
@echo 安装Kyocera Mita CS-2550驱动***************************************
@echo off
rundll32 printui.dll,PrintUIEntry /ia /h x86 /m "Kyocera Mita CS-2550" /f  "prnky004.inf"  /w /q && @echo 驱动成功安装
@echo 安装Kyocera Mita CS-2550打印机*************************************
@echo off
(prnmngr -a -p "Kyocera Mita CS-2550" -m "Kyocera Mita CS-2550" -r 10.135.51.161_9100) && (prncnfg -t -p "Kyocera Mita CS-2550" -l "生产二楼思瑞区域" -m "故障排除请联系IT")
@echo off
pause
@echo Kyocera Mita CS-2550正在打印测试页************************
prnqctl -p "Kyocera Mita CS-2550" -e && @echo 打印测试页发送成功！
pause
cls
EXIT


:4
@echo 添加打印端口**********************************************************
@echo off
prnport -a -r 10.135.48.126_9100 -h 10.135.48.126 -o raw
@echo 安装Kyocera Mita CS-2550驱动***************************************
@echo off
rundll32 printui.dll,PrintUIEntry /ia /h x86 /m "Kyocera Mita CS-2550" /f  "prnky004.inf"  /w /q && @echo 驱动成功安装
@echo 安装Kyocera Mita CS-2550打印机*************************************
@echo off
(prnmngr -a -p "Kyocera Mita CS-2550" -m "Kyocera Mita CS-2550" -r 10.135.48.126_9100) && (prncnfg -t -p "Kyocera Mita CS-2550" -l "办公三楼" -m "故障排除请联系IT")
@echo off
pause
@echo Kyocera Mita CS-2550正在打印测试页************************
prnqctl -p "Kyocera Mita CS-2550" -e && @echo 打印测试页发送成功！
pause
cls
EXIT

:5
@echo 添加打印端口**********************************************************
@echo off
prnport -a -r 10.135.48.66_9100 -h 10.135.48.66 -o raw
@echo 安装Kyocera Mita CS-2550驱动***************************************
@echo off
rundll32 printui.dll,PrintUIEntry /ia /h x86 /m "Kyocera Mita CS-2550" /f  "prnky004.inf"  /w /q && @echo 驱动成功安装
@echo 安装Kyocera Mita CS-2550打印机*************************************
@echo off
(prnmngr -a -p "Kyocera Mita CS-2550" -m "Kyocera Mita CS-2550" -r 10.135.48.66_9100) && (prncnfg -t -p "Kyocera Mita CS-2550" -l "办公三楼" -m "故障排除请联系IT")
@echo off
pause
@echo Kyocera Mita CS-2550正在打印测试页************************
prnqctl -p "Kyocera Mita CS-2550" -e && @echo 打印测试页发送成功！
pause
cls
EXIT

:EXITING
@echo 输入错误请重新运行本程序
pause
cls
:6  
exit  