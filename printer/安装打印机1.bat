@echo off 
@color 02



@cscript //h:cscript //s

@echo +++++������װ��ӡ������ĵ��ԣ����ں���ѡ��1/2/3/4����װ��Ҫ�Ĵ�ӡ����+++++

::@cd C:\Windows\System32\Printing_Admin_Scripts\zh-CN
@echo off
@echo ------------------------------------------------
@echo ��1.Kyocera Mita CS-2550��������¥�ߺ�����   ��
@echo ��2.Kyocera Mita CS-2550��������¥�м䣩  ���� ��
@echo ��3.Kyocera Mita CS-2550��������¥˼������ ����
@echo ��4.Kyocera Mita CS-2550 ���칫¥��¥��        ��
@echo ��5.Kyocera Mita CS-2550 ���칫¥2¥  ��   ������
@echo ��6.�ˡ����������������������� ��������������������
echo ------------------------------------------------
@set /p slection=��ѡ����Ҫ��װ�Ĵ�ӡ��(�����ż���):   


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
@echo ��Ӵ�ӡ�˿�**********************************************************
@echo off
prnport -a -r 10.135.51.162_9100 -h 10.135.51.162 -o raw
@echo ��װKyocera Mita CS-2550����***************************************
@echo off
rundll32 printui.dll,PrintUIEntry /ia /h x86 /m "Kyocera Mita CS-2550" /f  "prnky004.inf"  /w /q && @echo �����ɹ���װ
@echo ��װKyocera Mita CS-2550��ӡ��*************************************
@echo off
(prnmngr -a -p "Kyocera Mita CS-2550" -m "Kyocera Mita CS-2550" -r 10.135.51.162_9100) && (prncnfg -t -p "Kyocera Mita CS-2550" -l "������¥�ߺ�����" -m "�����ų�����ϵIT")
@echo off
pause
@echo Kyocera Mita CS-2550���ڴ�ӡ����ҳ************************
prnqctl -p "Kyocera Mita CS-2550" -e && @echo ��ӡ����ҳ���ͳɹ���
pause
cls

EXIT  



:2
@echo ��Ӵ�ӡ�˿�**********************************************************
@echo off
prnport -a -r 10.135.51.160_9100 -h 10.135.51.160 -o raw
@echo ��װKyocera Mita CS-2550����***************************************
@echo off
rundll32 printui.dll,PrintUIEntry /ia /h x86 /m "Kyocera Mita CS-2550" /f  "prnky004.inf"  /w /q && @echo �����ɹ���װ
@echo ��װHP LaserJet 2420 PCL 5e��ӡ��*************************************
@echo off
(prnmngr -a -p "Kyocera Mita CS-2550" -m "Kyocera Mita CS-2550" -r 10.135.51.160_9100) && (prncnfg -t -p "Kyocera Mita CS-2550" -l "������¥��¥�м�����" -m "�����ų�����ϵIT")
@echo off
pause
@echo Kyocera Mita CS-2550���ڴ�ӡ����ҳ************************
prnqctl -p "Kyocera Mita CS-2550" -e && @echo ��ӡ����ҳ���ͳɹ���
pause
cls
EXIT

:3
@echo ��Ӵ�ӡ�˿�**********************************************************
@echo off
prnport -a -r 10.135.51.161_9100 -h 10.135.51.161 -o raw
@echo ��װKyocera Mita CS-2550����***************************************
@echo off
rundll32 printui.dll,PrintUIEntry /ia /h x86 /m "Kyocera Mita CS-2550" /f  "prnky004.inf"  /w /q && @echo �����ɹ���װ
@echo ��װKyocera Mita CS-2550��ӡ��*************************************
@echo off
(prnmngr -a -p "Kyocera Mita CS-2550" -m "Kyocera Mita CS-2550" -r 10.135.51.161_9100) && (prncnfg -t -p "Kyocera Mita CS-2550" -l "������¥˼������" -m "�����ų�����ϵIT")
@echo off
pause
@echo Kyocera Mita CS-2550���ڴ�ӡ����ҳ************************
prnqctl -p "Kyocera Mita CS-2550" -e && @echo ��ӡ����ҳ���ͳɹ���
pause
cls
EXIT


:4
@echo ��Ӵ�ӡ�˿�**********************************************************
@echo off
prnport -a -r 10.135.48.126_9100 -h 10.135.48.126 -o raw
@echo ��װKyocera Mita CS-2550����***************************************
@echo off
rundll32 printui.dll,PrintUIEntry /ia /h x86 /m "Kyocera Mita CS-2550" /f  "prnky004.inf"  /w /q && @echo �����ɹ���װ
@echo ��װKyocera Mita CS-2550��ӡ��*************************************
@echo off
(prnmngr -a -p "Kyocera Mita CS-2550" -m "Kyocera Mita CS-2550" -r 10.135.48.126_9100) && (prncnfg -t -p "Kyocera Mita CS-2550" -l "�칫��¥" -m "�����ų�����ϵIT")
@echo off
pause
@echo Kyocera Mita CS-2550���ڴ�ӡ����ҳ************************
prnqctl -p "Kyocera Mita CS-2550" -e && @echo ��ӡ����ҳ���ͳɹ���
pause
cls
EXIT

:5
@echo ��Ӵ�ӡ�˿�**********************************************************
@echo off
prnport -a -r 10.135.48.66_9100 -h 10.135.48.66 -o raw
@echo ��װKyocera Mita CS-2550����***************************************
@echo off
rundll32 printui.dll,PrintUIEntry /ia /h x86 /m "Kyocera Mita CS-2550" /f  "prnky004.inf"  /w /q && @echo �����ɹ���װ
@echo ��װKyocera Mita CS-2550��ӡ��*************************************
@echo off
(prnmngr -a -p "Kyocera Mita CS-2550" -m "Kyocera Mita CS-2550" -r 10.135.48.66_9100) && (prncnfg -t -p "Kyocera Mita CS-2550" -l "�칫��¥" -m "�����ų�����ϵIT")
@echo off
pause
@echo Kyocera Mita CS-2550���ڴ�ӡ����ҳ************************
prnqctl -p "Kyocera Mita CS-2550" -e && @echo ��ӡ����ҳ���ͳɹ���
pause
cls
EXIT

:EXITING
@echo ����������������б�����
pause
cls
:6  
exit  