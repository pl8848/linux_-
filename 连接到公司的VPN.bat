 @echo off
 COLOR a
 :::::::::::本文件使用批处理快速创建VPN连接 ::::::::::::::::
echo --------------------------------------------------------
echo  │此工具可创建一个连接到公司的VPN，方便同事          │ 
echo  │在外面出差或在家时使用LX或服务器共享资料。         │
echo --------------------------------------------------------
pause 
(echo [VPN]
echo Encoding=1
echo Type=2
echo AutoLogon=0
echo UseRasCredentials=1
echo DialParamsUID=37523232
echo Guid=FFE24A0FFDE7414DABC592B4CF13E35F
echo BaseProtocol=1
echo VpnStrategy=2
echo ExcludedProtocols=0
echo LcpExtensions=1
echo DataEncryption=256
echo SwCompression=1
echo NegotiateMultilinkAlways=0
echo SkipNwcWarning=0
echo SkipDownLevelDialog=0
echo SkipDoubleDialDialog=0
echo DialMode=1
echo DialPercent=75
echo DialSeconds=120
echo HangUpPercent=10
echo HangUpSeconds=120
echo OverridePref=15
echo RedialAttempts=3
echo RedialSeconds=60
echo IdleDisconnectSeconds=0
echo RedialOnLinkFailure=0
echo CallbackMode=0
echo CustomDialDll=
echo CustomDialFunc=
echo CustomRasDialDll=
echo AuthenticateServer=0
echo ShareMsFilePrint=1
echo BindMsNetClient=1
echo SharedPhoneNumbers=0
echo GlobalDeviceSettings=0
echo PrerequisiteEntry=
echo PrerequisitePbk=
echo PreferredPort=VPN4-0
echo PreferredDevice=WAN 微型端口 (L2TP^)
echo PreferredBps=0
echo PreferredHwFlow=1
echo PreferredProtocol=1
echo PreferredCompression=1
echo PreferredSpeaker=1
echo PreferredMdmProtocol=0
echo PreviewUserPw=1
echo PreviewDomain=0
echo PreviewPhoneNumber=0
echo ShowDialingProgress=1
echo ShowMonitorIconInTaskBar=1
echo CustomAuthKey=-1
echo AuthRestrictions=608
echo TypicalAuth=2
echo IpPrioritizeRemote=0
rem 上面这条如果是1就开启“IPv4使用远程网关” ，0就关掉
echo IpHeaderCompression=0
echo IpAddress=0.0.0.0
echo IpDnsAddress=0.0.0.0
echo IpDns2Address=0.0.0.0
echo IpWinsAddress=0.0.0.0
echo IpWins2Address=0.0.0.0
echo IpAssign=1
echo IpNameAssign=1
echo IpFrameSize=1006
echo IpDnsFlags=0
echo IpNBTFlags=1
echo TcpWindowSize=0
echo UseFlags=0
echo IpSecFlags=0
echo IpDnsSuffix=
echo NETCOMPONENTS=
echo ms_server=1
echo ms_msclient=1
echo ms_psched=1
echo MEDIA=rastapi
echo Port=VPN4-0
echo Device=WAN 微型端口 (L2TP^)
echo DEVICE=VPN
echo PhoneNumber=URL
rem 上面创建连接地址(这里为：URL)
echo AreaCode=
echo CountryCode=1
echo CountryID=1
echo UseDialingRules=0
echo Comment=
echo LastSelectedPhone=0
echo PromoteAlternates=0
echo TryNextAlternateOnFail=1)>%temp%\VPN.pbk
for /f "tokens=2,*" %%i in ('reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /v "Desktop"') do (set desk=%%j)
 
copy /y %temp%\vpn.pbk "%desk%\连接到公司的vpn.pbk

echo ---------------------------------创建完成--------------------------------
    
echo 1，双击桌面“连接到公司的VPN”图标进行连接，连接需要账号。              │
echo -------------------------------------------------------------------------
echo 2，账号可向IT申请。此连接用于在公司外面使用LX或服务器共享资料。         │
echo -------------------------------------------------------------------------

echo 3，使用完成可再次双击桌面“连接到公司的VPN”图标，点“挂断”来中止连接。│
echo -------------------------------------------------------------------------
echo 4，严禁外泄此工具及用于连接的账号，否则追究责任。                       │
echo -------------------------------------------------------------------------

pause 

exit


