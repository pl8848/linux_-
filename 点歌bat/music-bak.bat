@echo on
	 set ProcessName1=KuGou.exe
	 set processName=wmplayer.exe
	 set processName2=cmd.exe
	 start /min ""  "D:\Program Files\Windows Media Player\wmplayer.exe" 
	 ping -n 20 127.1 >nul  

 :panduang

	::����mp3�ļ�
	::	��ʼ��co

	set co = 1
	for /r "f:\music" %%a in (*.mp3) do (
		set fn=%%~na
		if "!fn:~0,1!" neq "0" (
		set /a co += 1
		echo %%~a>>1.txt
		)
	)

	::�������%co%<3����KuGou.exe"�ṷ,��������WMP

	echo  %co% 
	
	::if %co% lss 3 goto :kugou 
	if defined co (echo ����str�Ѿ�����ֵ����ֵΪ%co%) else (goto  :kugou ) 
	if %co% LSS 3  goto :kugou 
	if !co! GEQ 3  goto :wmplayer
	
	::echo "%te%"
	echo wmp
	goto :wmplayer

	echo ͳ�Ʋ��ɹ�
	ping -n 20 127.1 >nul 
	goto :exit
 
:kugou

taskkill /f /im "%ProcessName%"
taskkill /f /im "%ProcessName1%"
	start ""  "D:\Program Files\KuGou\KGMusic\KuGou.exe"

	::��������������ò���ʱ��
	ping -n 800 127.1 >nul
	echo no��
	goto :exit

:wmplayer
   
taskkill /f /im "%ProcessName%"
taskkill /f /im "%ProcessName1%"
	  ping -n 2 127.1 >nul  
	  echo �����������ϵͳ
	 start /max ""  "f:\musicbak\Playlists\�Զ������б�.wpl"

   ::��������������ò���ʱ��
	ping -n 800 127.1 >nul
	
	::�����ȹز�������������ɾ��
	
	taskkill /f /im "%ProcessName%"
	taskkill /f /im "%ProcessName1%"
	::����ɾ��ǰ�����Ѳ��Ź����ļ�
	for /f "delims=" %%i in ('dir "f:\Music\*.mp3"  /s /b') do copy "%%i" "F:\mp3����\"
 
	for /f "tokens=1* delims=:" %%i in ('findstr /n ".*" 1.txt') do (
      if %%i==1   del /s /q "%%j"
      if %%i==2   del /s /q "%%j"
      if %%i==3   del /s /q "%%j"
		)
		
		ping -n 3 127.1 >nul

		::taskkill /f /im "%ProcessName2%" ���ﲻ���ȹ�CMD


goto :exit

:exit
del /q /s 1.txt
taskkill /f /im "%ProcessName%"
taskkill /f /im "%ProcessName1%"
taskkill /f /im "%ProcessName2%"
 exit