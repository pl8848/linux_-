@echo on
	 set ProcessName1=KuGou.exe
	 set processName=wmplayer.exe
	 set processName2=cmd.exe
	 start /min ""  "D:\Program Files\Windows Media Player\wmplayer.exe" 
	 ping -n 20 127.1 >nul  

 :panduang

	::遍历mp3文件
	::	初始化co

	set co = 1
	for /r "f:\music" %%a in (*.mp3) do (
		set fn=%%~na
		if "!fn:~0,1!" neq "0" (
		set /a co += 1
		echo %%~a>>1.txt
		)
	)

	::这里如果%co%<3运行KuGou.exe"酷狗,否则运行WMP

	echo  %co% 
	
	::if %co% lss 3 goto :kugou 
	if defined co (echo 变量str已经被赋值，其值为%co%) else (goto  :kugou ) 
	if %co% LSS 3  goto :kugou 
	if !co! GEQ 3  goto :wmplayer
	
	::echo "%te%"
	echo wmp
	goto :wmplayer

	echo 统计不成功
	ping -n 20 127.1 >nul 
	goto :exit
 
:kugou

taskkill /f /im "%ProcessName%"
taskkill /f /im "%ProcessName1%"
	start ""  "D:\Program Files\KuGou\KGMusic\KuGou.exe"

	::下面这句这里设置播放时间
	ping -n 800 127.1 >nul
	echo no有
	goto :exit

:wmplayer
   
taskkill /f /im "%ProcessName%"
taskkill /f /im "%ProcessName1%"
	  ping -n 2 127.1 >nul  
	  echo 正在启动点歌系统
	 start /max ""  "f:\musicbak\Playlists\自动播放列表.wpl"

   ::下面这句这里设置播放时间
	ping -n 800 127.1 >nul
	
	::这里先关播放器，否则不能删除
	
	taskkill /f /im "%ProcessName%"
	taskkill /f /im "%ProcessName1%"
	::这里删除前三个已播放过的文件
	for /f "delims=" %%i in ('dir "f:\Music\*.mp3"  /s /b') do copy "%%i" "F:\mp3样本\"
 
	for /f "tokens=1* delims=:" %%i in ('findstr /n ".*" 1.txt') do (
      if %%i==1   del /s /q "%%j"
      if %%i==2   del /s /q "%%j"
      if %%i==3   del /s /q "%%j"
		)
		
		ping -n 3 127.1 >nul

		::taskkill /f /im "%ProcessName2%" 这里不能先关CMD


goto :exit

:exit
del /q /s 1.txt
taskkill /f /im "%ProcessName%"
taskkill /f /im "%ProcessName1%"
taskkill /f /im "%ProcessName2%"
 exit