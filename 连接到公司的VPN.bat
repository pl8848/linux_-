 @echo off
 COLOR a
 :::::::::::���ļ�ʹ����������ٴ���VPN���� ::::::::::::::::
echo --------------------------------------------------------
echo  ���˹��߿ɴ���һ�����ӵ���˾��VPN������ͬ��          �� 
echo  �������������ڼ�ʱʹ��LX��������������ϡ�         ��
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
echo PreferredDevice=WAN ΢�Ͷ˿� (L2TP^)
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
rem �������������1�Ϳ�����IPv4ʹ��Զ�����ء� ��0�͹ص�
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
echo Device=WAN ΢�Ͷ˿� (L2TP^)
echo DEVICE=VPN
echo PhoneNumber=URL
rem ���洴�����ӵ�ַ(����Ϊ��URL)
echo AreaCode=
echo CountryCode=1
echo CountryID=1
echo UseDialingRules=0
echo Comment=
echo LastSelectedPhone=0
echo PromoteAlternates=0
echo TryNextAlternateOnFail=1)>%temp%\VPN.pbk
for /f "tokens=2,*" %%i in ('reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /v "Desktop"') do (set desk=%%j)
 
copy /y %temp%\vpn.pbk "%desk%\���ӵ���˾��vpn.pbk

echo ---------------------------------�������--------------------------------
    
echo 1��˫�����桰���ӵ���˾��VPN��ͼ��������ӣ�������Ҫ�˺š�              ��
echo -------------------------------------------------------------------------
echo 2���˺ſ���IT���롣�����������ڹ�˾����ʹ��LX��������������ϡ�         ��
echo -------------------------------------------------------------------------

echo 3��ʹ����ɿ��ٴ�˫�����桰���ӵ���˾��VPN��ͼ�꣬�㡰�Ҷϡ�����ֹ���ӡ���
echo -------------------------------------------------------------------------
echo 4���Ͻ���й�˹��߼��������ӵ��˺ţ�����׷�����Ρ�                       ��
echo -------------------------------------------------------------------------

pause 

exit


