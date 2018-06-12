
@echo off
	 set ProcessName1=KuGou.exe
	 set processName=wmplayer.exe

	 set ProcessName3=DoubanRadio.exe
	 set processName2=cmd.exe

       
del /q /s 1.txt
del /q /s "f:\play\*.mp3"
del /q /s "f:\play\*.jpg"
del /q /s "Z:\it\music\*.jpg"
del /q /s "Z:\it\music\*.txt"

del /q /s 1.txt
del /q /s "f:\play\*.mp3"
del /q /s "f:\music\*.jpg"
del /q /s "f:\music\*.txt"

taskkill /f /im "%ProcessName%"
taskkill /f /im "%ProcessName1%"
taskkill /f /im "%ProcessName2%"
taskkill /f /im  "ProcessName3"



exit