
@echo off

rem ��Ҫreg.exe��֧��
rem �޷���֤���С�Ӣ֮����������ԵĲ���ϵͳ�ϵõ���ȷ���
for /f "skip=4 delims= " %%a in ('reg query "HKEY_CURRENT_USER\Control Panel\International" /v sShortDate') do set DateFormat=%%a
set DateFormat=%DateFormat:yyyy/M/d%
reg add "HKEY_CURRENT_USER\Control Panel\International" /v sShortDate /t REG_SZ /d yyyy-M-d /f>nul
set Today=%date: =%
reg add "HKEY_CURRENT_USER\Control Panel\International" /v sShortDate /t REG_SZ /d %DateFormat% /f>nul
set "Week=Mon Tue Wed Thu Fri Sat Sun ����һ ���ڶ� ������ ������ ������ ������ ������"
for %%a in (%Week%) do call set "Today=%%Today:%%a=%%"


::��ȡʱ���е�Сʱ ����ʽ����Ϊ��24Сʱ��
set timevar=%time:~0,2%
if /i %timevar% LSS 10 (
set timevar=0%time:~1,1%
)


::��ȡʱ���еķ֡��� ����ʽ����Ϊ��3220 ����ʾ 32��20��
::set timevar=%timevar%%time:~3,2%%time:~6,2%
@echo %Today%--%time%%time:~6,2% >>"f:\play\������ʷ.txt"



@echo off
	 set ProcessName1=KuGou.exe
	 set processName=wmplayer.exe
	 set processName2=cmd.exe
	 

 :panduang

::����mp3�ļ�
::	��ʼ��co

	set co = 1
	for /r "Z:\it\music" %%a in (*.mp3) do (
		set fn=%%~na
		if "!fn:~0,1!" neq "0" (
		set /a co += 1
		echo %%~a>>1.txt
               
		)
	)      


::�������%co%<3����KuGou.exe"�ṷ,��������WMP

::	echo  %co% 
	
::if %co% lss 3 goto :kugou 
	if defined co (echo ����str�Ѿ�����ֵ����ֵΪ%co%) else (goto  :kugou ) 
	if %co% LSS 3  goto :kugou 
	if !co! GEQ 3  goto :wmplayer
	
	::echo "%te%"
	echo wmp f:\play\������ʷ.txt
	goto :wmplayer

	echo ͳ�Ʋ��ɹ�
	ping -n 20 127.1 >nul 
	goto :exit
 
:kugou

taskkill /f /im "%ProcessName%"
taskkill /f /im "%ProcessName1%"

echo %ProcessName1%>>f:\play\������ʷ.txt
ping -n 15 127.1 >nul 
	start ""  "D:\Program Files\KuGou\KGMusic\KuGou.exe"

::��������������ò���ʱ��
	ping -n 900 127.1 >nul
	echo no��
	goto :exit

:wmplayer
   
taskkill /f /im "%ProcessName%"
taskkill /f /im "%ProcessName1%"

::����ɨ��ý���ļ�
  
	for /f "tokens=1* delims=:" %%i in ('findstr /n ".*" 1.txt') do (
      if %%i==1   copy /y "%%j" "f:\play"  && del /q /s "%%j"  &&  echo %%j>>f:\play\������ʷ.txt
      if %%i==2   copy /y "%%j" "f:\play"  && del /q /s "%%j"  &&  echo %%j>>f:\play\������ʷ.txt
      if %%i==3   copy  /y "%%j" "f:\play" && del /q /s "%%j"  &&  echo %%j>>f:\play\������ʷ.txt
	    if %%i==4   copy  /y "%%j" "f:\play" && del /q /s "%%j"  &&  echo %%j>>f:\play\������ʷ.txt

	)

	start /min ""  "D:\Program Files\Windows Media Player\wmplayer.exe"
	 ping -n 22 127.1 >nul 
	 
	 
	  echo �����������ϵͳ
	 start /max ""  "f:\musicbak\Playlists\�Զ������б�.wpl"

 ::��������������ò���ʱ��
	ping -n 900 127.1 >nul
	
::�����ȹز�������������ɾ��
	
	taskkill /f /im "%ProcessName%"
	taskkill /f /im "%ProcessName1%"
::����ɾ��ǰ�����Ѳ��Ź����ļ�
         for /f "delims=" %%i in ('dir "f:\play\*.mp3"  /s /b') do copy /y "%%i" "f:\mp3����\"


	    
 
		
		ping -n 3 127.1 >nul

::taskkill /f /im "%ProcessName2%" ���ﲻ���ȹ�CMD


goto :exit

:exit
del /q /s 1.txt
del /q /s "f:\play\*.mp3"
del /q /s "f:\play\*.jpg"
del /q /s "Z:\it\music\*.jpg"
del /q /s "Z:\it\music\*.txt"

taskkill /f /im "%ProcessName%"
taskkill /f /im "%ProcessName1%"
taskkill /f /im "%ProcessName2%"
 exit