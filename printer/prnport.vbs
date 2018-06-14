'----------------------------------------------------------------------
'
' Copyright (c) Microsoft Corporation. All rights reserved.
'
' Abstract:
' prnport.vbs - Port script for WMI on Windows 
'     used to add, delete and list ports
'     also for getting and setting the port configuration
'
' Usage:
' prnport [-adlgt?] [-r port] [-s server] [-u user name] [-w password]
'                   [-o raw|lpr] [-h host address] [-q queue] [-n number]
'                   [-me | -md ] [-i SNMP index] [-y community] [-2e | -2d]"
'
' Examples
' prnport -a -s server -r IP_1.2.3.4 -e 1.2.3.4 -o raw -n 9100
' prnport -d -s server -r c:\temp\foo.prn
' prnport -l -s server
' prnport -g -s server -r IP_1.2.3.4
' prnport -t -s server -r IP_1.2.3.4 -me -y public -i 1 -n 9100
'
'----------------------------------------------------------------------

option explicit

'
' Debugging trace flags, to enable debug output trace message
' change gDebugFlag to true.
'
dim   gDebugFlag
const kDebugTrace = 1
const kDebugError = 2

gDebugFlag = false

'
' Operation action values.
'
const kActionAdd          = 0
const kActionDelete       = 1
const kActionList         = 2
const kActionUnknown      = 3
const kActionGet          = 4
const kActionSet          = 5

const kErrorSuccess       = 0
const KErrorFailure       = 1

const kFlagCreateOrUpdate = 0

const kNameSpace          = "root\cimv2"


'
' Constants for the parameter dictionary
'
const kServerName      = 1
const kPortName        = 2
const kDoubleSpool     = 3
const kPortNumber      = 4
const kPortType        = 5
const kHostAddress     = 6
const kSNMPDeviceIndex = 7
const kCommunityName   = 8
const kSNMP            = 9
const kQueueName       = 10
const kUserName        = 11
const kPassword        = 12

'
' Generic strings
'
const L_Empty_Text                 = ""
const L_Space_Text                 = " "
const L_Colon_Text                 = ":"
const L_LPR_Queue                  = "LPR"
const L_Error_Text                 = "错误"
const L_Success_Text               = "成功"
const L_Failed_Text                = "失败"
const L_Hex_Text                   = "0x"
const L_Printer_Text               = "打印机"
const L_Operation_Text             = "操作"
const L_Provider_Text              = "提供程序"
const L_Description_Text           = "描述"
const L_Debug_Text                 = "调试:"

'
' General usage messages
'
const L_Help_Help_General01_Text   = "用法: prnport [-adlgt?] [-r 端口][-s 服务器][-u 用户名][-w 密码]"
const L_Help_Help_General02_Text   = "               [-o raw|lpr][-h 主机地址][-q 队列][-n 编号]"
const L_Help_Help_General03_Text   = "               [-me | -md ][-i SNMP 索引][-y 团体][-2e | -2d]"
const L_Help_Help_General04_Text   = "参数:"
const L_Help_Help_General05_Text   = "-a     - 添加端口"
const L_Help_Help_General06_Text   = "-d     - 删除指定端口"
const L_Help_Help_General07_Text   = "-g     - 获得 TCP 端口的配置"
const L_Help_Help_General08_Text   = "-h     - 设备的 IP 地址"
const L_Help_Help_General09_Text   = "-i     - SNMP 索引，如果已启用 SNMP"
const L_Help_Help_General10_Text   = "-l     - 列出所有 TCP 端口"
const L_Help_Help_General11_Text   = "-m     - SNMP 类型。[e] 启用，[d] 禁用"
const L_Help_Help_General12_Text   = "-n     - 端口编号，应用于 TCP RAW 端口"
const L_Help_Help_General13_Text   = "-o     - 端口类型，raw 或 lpr"
const L_Help_Help_General14_Text   = "-q     - 队列名称，只应用于 TCP LPR 端口"
const L_Help_Help_General15_Text   = "-r     - 端口名"
const L_Help_Help_General16_Text   = "-s     - 服务器名"
const L_Help_Help_General17_Text   = "-t     - 设定 TCP 端口的配置"
const L_Help_Help_General18_Text   = "-u     - 用户名"
const L_Help_Help_General19_Text   = "-w     - 密码"
const L_Help_Help_General20_Text   = "-y     - 社区名，如果已启用 SNMP"
const L_Help_Help_General21_Text   = "-2     - 双后台处理程序，应用于 TCP LPR 端口。[e] 启用，[d] 禁用"
const L_Help_Help_General22_Text   = "-?     - 显示命令用法"
const L_Help_Help_General23_Text   = "例如:"
const L_Help_Help_General24_Text   = "prnport -l -s 服务器"
const L_Help_Help_General25_Text   = "prnport -d -s 服务器 -r IP_1.2.3.4"
const L_Help_Help_General26_Text   = "prnport -a -s 服务器 -r IP_1.2.3.4 -h 1.2.3.4 -o raw -n 9100"
const L_Help_Help_General27_Text   = "prnport -t -s 服务器 -r IP_1.2.3.4 -me -y public -i 1 -n 9100"
const L_Help_Help_General28_Text   = "prnport -g -s 服务器 -r IP_1.2.3.4"
const L_Help_Help_General29_Text   = "prnport -a -r IP_1.2.3.4 -h 1.2.3.4"
const L_Help_Help_General30_Text   = "注解:"
const L_Help_Help_General31_Text   = "上一个例子将试图获取指定 IP 地址的设备设置。"
const L_Help_Help_General32_Text   = "如果已检测到设备，那个设备的首选设置会被添加到 TCP 端口。"

'
' Messages to be displayed if the scripting host is not cscript
'
const L_Help_Help_Host01_Text      = "此脚本应该用 CSCRIPT.EXE 在命令提示符下执行。"
const L_Help_Help_Host02_Text      = "例如: CScript script.vbs 参数"
const L_Help_Help_Host03_Text      = ""
const L_Help_Help_Host04_Text      = "要设置 CScript 作为运行 .vbs 文件的默认应用程序，运行如下:"
const L_Help_Help_Host05_Text      = "     CScript //H:CScript //S"
const L_Help_Help_Host06_Text      = "然后可以运行 ""script.vbs arguments""，而无需在运行脚本前先使用 CScript。"

'
' General error messages
'
const L_Text_Error_General01_Text  = "无法决定脚本操作主机。"
const L_Text_Error_General02_Text  = "无法分析命令行。"
const L_Text_Error_General03_Text  = "Win32 错误代码"

'
' Miscellaneous messages
'
const L_Text_Msg_General01_Text    = "已添加端口"
const L_Text_Msg_General02_Text    = "无法删除端口"
const L_Text_Msg_General03_Text    = "无法获取端口"
const L_Text_Msg_General04_Text    = "已创建/更新端口"
const L_Text_Msg_General05_Text    = "无法创建/更新端口"
const L_Text_Msg_General06_Text    = "无法枚举端口"
const L_Text_Msg_General07_Text    = "枚举的端口数量"
const L_Text_Msg_General08_Text    = "已删除的端口"
const L_Text_Msg_General09_Text    = "无法获取 SWbemLocator 对象"
const L_Text_Msg_General10_Text    = "无法连接到 WMI 服务"


'
' Port properties
'
const L_Text_Msg_Port01_Text       = "服务器名"
const L_Text_Msg_Port02_Text       = "端口名"
const L_Text_Msg_Port03_Text       = "主机地址"
const L_Text_Msg_Port04_Text       = "协议 RAW"
const L_Text_Msg_Port05_Text       = "协议 LPR"
const L_Text_Msg_Port06_Text       = "端口编号"
const L_Text_Msg_Port07_Text       = "队列"
const L_Text_Msg_Port08_Text       = "已启用字节计数"
const L_Text_Msg_Port09_Text       = "已禁用字节计数"
const L_Text_Msg_Port10_Text       = "已启用 SNMP "
const L_Text_Msg_Port11_Text       = "已禁用 SNMP"
const L_Text_Msg_Port12_Text       = "团体"
const L_Text_Msg_Port13_Text       = "设备索引"

'
' Debug messages
'
const L_Text_Dbg_Msg01_Text        = "于函数 DelPort"
const L_Text_Dbg_Msg02_Text        = "于函数 CreateOrSetPort"
const L_Text_Dbg_Msg03_Text        = "于函数 ListPorts"
const L_Text_Dbg_Msg04_Text        = "于函数 GetPort"
const L_Text_Dbg_Msg05_Text        = "于函数 ParseCommandLine"

main

'
' Main execution starts here
'
sub main

    on error resume next

    dim iAction
    dim iRetval
    dim oParamDict

    '
    ' Abort if the host is not cscript
    '
    if not IsHostCscript() then

        call wscript.echo(L_Help_Help_Host01_Text & vbCRLF & L_Help_Help_Host02_Text & vbCRLF & _
                          L_Help_Help_Host03_Text & vbCRLF & L_Help_Help_Host04_Text & vbCRLF & _
                          L_Help_Help_Host05_Text & vbCRLF & L_Help_Help_Host06_Text & vbCRLF)

        wscript.quit

    end if

    set oParamDict = CreateObject("Scripting.Dictionary")

    iRetval = ParseCommandLine(iAction, oParamDict)

    if iRetval = 0 then

        select case iAction

            case kActionAdd
                iRetval = CreateOrSetPort(oParamDict)

            case kActionDelete
                iRetval = DelPort(oParamDict)

            case kActionList
                iRetval = ListPorts(oParamDict)

            case kActionGet
                iRetVal = GetPort(oParamDict)

            case kActionSet
                iRetVal = CreateOrSetPort(oParamDict)

            case else
                Usage(true)
                exit sub

        end select

    end if

end sub

'
' Delete a port
'
function DelPort(oParamDict)

    on error resume next

    DebugPrint kDebugTrace, L_Text_Dbg_Msg01_Text
    DebugPrint kDebugTrace, L_Text_Msg_Port01_Text & L_Space_Text & oParamDict(kServerName)
    DebugPrint kDebugTrace, L_Text_Msg_Port02_Text & L_Space_Text & oParamDict(kPortName)

    dim oService
    dim oPort
    dim iResult
    dim strServer
    dim strPort
    dim strUser
    dim strPassword

    iResult = kErrorFailure

    strServer   = oParamDict(kServerName)
    strPort     = oParamDict(kPortName)
    strUser     = oParamDict(kUserName)
    strPassword = oParamDict(kPassword)

    if WmiConnect(strServer, kNameSpace, strUser, strPassword, oService) then

        set oPort = oService.Get("Win32_TCPIPPrinterPort='" & strPort & "'")

    else

        DelPort = kErrorFailure

        exit function

    end if

    '
    ' Check if Get succeeded
    '
    if Err.Number = kErrorSuccess then

        '
        ' Try deleting the instance
        '
        oPort.Delete_

        if Err.Number = kErrorSuccess then

            wscript.echo L_Text_Msg_General08_Text & L_Space_Text & strPort

        else

            wscript.echo L_Text_Msg_General02_Text & L_Space_Text & L_Error_Text & L_Space_Text _
                         & L_Hex_Text & hex(Err.Number) & L_Space_Text & Err.Description

            '
            ' Try getting extended error information
            '
            call LastError()

        end if

    else

        wscript.echo L_Text_Msg_General02_Text & L_Space_Text & L_Error_Text & L_Space_Text _
                     & L_Hex_Text & hex(Err.Number) & L_Space_Text & Err.Description

        '
        ' Try getting extended error information
        '
        call LastError()

    end if

    DelPort = iResult

end function

'
' Add or update a port
'
function CreateOrSetPort(oParamDict)

    on error resume next

    dim oPort
    dim oService
    dim iResult
    dim PortType
    dim strServer
    dim strPort
    dim strUser
    dim strPassword

    DebugPrint kDebugTrace, L_Text_Dbg_Msg02_Text
    DebugPrint kDebugTrace, L_Text_Msg_Port01_Text & L_Space_Text & oParamDict.Item(kServerName)
    DebugPrint kDebugTrace, L_Text_Msg_Port02_Text & L_Space_Text & oParamDict.Item(kPortName)
    DebugPrint kDebugTrace, L_Text_Msg_Port06_Text & L_Space_Text & oParamDict.Item(kPortNumber)
    DebugPrint kDebugTrace, L_Text_Msg_Port07_Text & L_Space_Text & oParamDict.Item(kQueueName)
    DebugPrint kDebugTrace, L_Text_Msg_Port13_Text & L_Space_Text & oParamDict.Item(kSNMPDeviceIndex)
    DebugPrint kDebugTrace, L_Text_Msg_Port12_Text & L_Space_Text & oParamDict.Item(kCommunityName)
    DebugPrint kDebugTrace, L_Text_Msg_Port03_Text & L_Space_Text & oParamDict.Item(kHostAddress)

    strServer   = oParamDict(kServerName)
    strPort     = oParamDict(kPortName)
    strUser     = oParamDict(kUserName)
    strPassword = oParamDict(kPassword)

    '
    ' If the port exists, then get the settings. Later PutInstance will do an update
    '
    if WmiConnect(strServer, kNameSpace, strUser, strPassword, oService) then

        set oPort = oService.Get("Win32_TCPIPPrinterPort.Name='" & strPort & "'")

        '
        ' If get was unsuccessful then spawn a new port instance. Later PutInstance will do a create
        '
        if Err.Number <> kErrorSuccess then

            '
            ' Clear the previous error
            '
            Err.Clear

            set oPort = oService.Get("Win32_TCPIPPrinterPort").SpawnInstance_

        end if

    else

        CreateOrSetPort = kErrorFailure

        exit function

    end if

    if Err.Number <> kErrorSuccess then

        wscript.echo L_Text_Msg_General03_Text & L_Space_Text & L_Error_Text & L_Space_Text _
                     & L_Hex_Text & hex(Err.Number) & L_Space_Text & Err.Description

        CreateOrSetPort = kErrorFailure

        exit function

    end if

    oPort.Name          = oParamDict.Item(kPortName)
    oPort.HostAddress   = oParamDict.Item(kHostAddress)
    oPort.PortNumber    = oParamDict.Item(kPortNumber)
    oPort.SNMPEnabled   = oParamDict.Item(kSNMP)
    oPort.SNMPDevIndex  = oParamDict.Item(kSNMPDeviceIndex)
    oPort.SNMPCommunity = oParamDict.Item(kCommunityName)
    oPort.Queue         = oParamDict.Item(kQueueName)
    oPort.ByteCount     = oParamDict.Item(kDoubleSpool)

    PortType     = oParamDict.Item(kPortType)

    '
    ' Update the port object with the settings corresponding
    ' to the port type of the port to be added
    '
    select case lcase(PortType)

            case "raw"

                 oPort.Protocol      = 1

                 if Not IsNull(oPort.Queue) then

                     wscript.echo L_Error_Text & L_Colon_Text & L_Space_Text _
                     & L_Help_Help_General14_Text

                     CreateOrSetPort = kErrorFailure

                     exit function

                 end if

            case "lpr"

                 oPort.Protocol      = 2

                 if IsNull(oPort.Queue) then

                     oPort.Queue = L_LPR_Queue

                 end if

            case else

                 '
                 ' PutInstance will attempt to get the configuration of
                 ' the device based on its IP address. Those settings
                 ' will be used to add a new port
                 '
    end select

    '
    ' Try creating or updating the port
    '
    oPort.Put_(kFlagCreateOrUpdate)

    if Err.Number = kErrorSuccess then

        wscript.echo L_Text_Msg_General04_Text & L_Space_Text & oPort.Name

        iResult = kErrorSuccess

    else

        wscript.echo L_Text_Msg_General05_Text & L_Space_Text & oPort.Name & L_Space_Text _
                     & L_Error_Text & L_Space_Text & L_Hex_Text & hex(Err.Number) _
                     & L_Space_Text & Err.Description

        '
        ' Try getting extended error information
        '
        call LastError()

        iResult = kErrorFailure

    end if

    CreateOrSetPort = iResult

end function

'
' List ports on a machine.
'
function ListPorts(oParamDict)

    on error resume next

    DebugPrint kDebugTrace, L_Text_Dbg_Msg03_Text

    dim Ports
    dim oPort
    dim oService
    dim iRetval
    dim iTotal
    dim strServer
    dim strUser
    dim strPassword

    iResult = kErrorFailure

    strServer   = oParamDict(kServerName)
    strUser     = oParamDict(kUserName)
    strPassword = oParamDict(kPassword)

    if WmiConnect(strServer, kNameSpace, strUser, strPassword, oService) then

        set Ports = oService.InstancesOf("Win32_TCPIPPrinterPort")

    else

        ListPorts = kErrorFailure

        exit function

    end if

    if Err.Number <> kErrorSuccess then

        wscript.echo L_Text_Msg_General06_Text & L_Space_Text & L_Error_Text & L_Space_Text _
                     & L_Hex_Text & hex(Err.Number) & L_Space_Text & Err.Description

        ListPrinters = kErrorFailure

        exit function

    end if

    iTotal = 0

    for each oPort in Ports

        iTotal = iTotal + 1

        wscript.echo L_Empty_Text
        wscript.echo L_Text_Msg_Port01_Text & L_Space_Text & strServer
        wscript.echo L_Text_Msg_Port02_Text & L_Space_Text & oPort.Name
        wscript.echo L_Text_Msg_Port03_Text & L_Space_Text & oPort.HostAddress

        if oPort.Protocol = 1 then

            wscript.echo L_Text_Msg_Port04_Text
            wscript.echo L_Text_Msg_Port06_Text & L_Space_Text & oPort.PortNumber

        else

            wscript.echo L_Text_Msg_Port05_Text
            wscript.echo L_Text_Msg_Port07_Text & L_Space_Text & oPort.Queue

            if oPort.ByteCount then

                wscript.echo L_Text_Msg_Port08_Text

            else

                wscript.echo L_Text_Msg_Port09_Text

            end if

        end if

        if oPort.SNMPEnabled then

            wscript.echo L_Text_Msg_Port10_Text
            wscript.echo L_Text_Msg_Port12_Text & L_Space_Text & oPort.SNMPCommunity
            wscript.echo L_Text_Msg_Port13_Text & L_Space_Text & oPort.SNMPDevIndex

        else

            wscript.echo L_Text_Msg_Port11_Text

        end if

        Err.Clear

    next

    wscript.echo L_Empty_Text
    wscript.echo L_Text_Msg_General07_Text & L_Space_Text & iTotal

    ListPorts = kErrorSuccess

end function

'
' Gets the configuration of a port
'
function GetPort(oParamDict)

    on error resume next

    DebugPrint kDebugTrace, L_Text_Dbg_Msg04_Text
    DebugPrint kDebugTrace, L_Text_Msg_Port01_Text & L_Space_Text & oParamDict(kServerName)
    DebugPrint kDebugTrace, L_Text_Msg_Port02_Text & L_Space_Text & oParamDict(kPortName)

    dim oService
    dim oPort
    dim iResult
    dim strServer
    dim strPort
    dim strUser
    dim strPassword

    iResult = kErrorFailure

    strServer   = oParamDict(kServerName)
    strPort     = oParamDict(kPortName)
    strUser     = oParamDict(kUserName)
    strPassword = oParamDict(kPassword)

    if WmiConnect(strServer, kNameSpace, strUser, strPassword, oService) then

        set oPort = oService.Get("Win32_TCPIPPrinterPort.Name='" & strPort & "'")

    else

        GetPort = kErrorFailure

        exit function

    end if

    if Err.Number = kErrorSuccess then

        wscript.echo L_Empty_Text
        wscript.echo L_Text_Msg_Port01_Text & L_Space_Text & strServer
        wscript.echo L_Text_Msg_Port02_Text & L_Space_Text & oPort.Name
        wscript.echo L_Text_Msg_Port03_Text & L_Space_Text & oPort.HostAddress

        if oPort.Protocol = 1 then

            wscript.echo L_Text_Msg_Port04_Text
            wscript.echo L_Text_Msg_Port06_Text & L_Space_Text & oPort.PortNumber

        else

            wscript.echo L_Text_Msg_Port05_Text
            wscript.echo L_Text_Msg_Port07_Text & L_Space_Text & oPort.Queue

            if oPort.ByteCount then

                wscript.echo L_Text_Msg_Port08_Text

            else

                wscript.echo L_Text_Msg_Port09_Text

            end if

        end if

        if oPort.SNMPEnabled then

            wscript.echo L_Text_Msg_Port10_Text
            wscript.echo L_Text_Msg_Port12_Text & L_Space_Text & oPort.SNMPCommunity
            wscript.echo L_Text_Msg_Port13_Text & L_Space_Text & oPort.SNMPDevIndex

        else

            wscript.echo L_Text_Msg_Port11_Text

        end if

        iResult = kErrorSuccess

    else

        wscript.echo L_Text_Msg_General03_Text & L_Space_Text & L_Error_Text & L_Space_Text _
                     & L_Hex_Text & hex(Err.Number) & L_Space_Text & Err.Description

        '
        ' Try getting extended error information
        '
        call LastError()

    end if

    GetPort = iResult

end function

'
' Debug display helper function
'
sub DebugPrint(uFlags, strString)

    if gDebugFlag = true then

        if uFlags = kDebugTrace then

            wscript.echo L_Debug_Text & L_Space_Text & strString

        end if

        if uFlags = kDebugError then

            if Err <> 0 then

                wscript.echo L_Debug_Text & L_Space_Text & strString & L_Space_Text _
                             & L_Error_Text & L_Space_Text & L_Hex_Text & hex(Err.Number) _
                             & L_Space_Text & Err.Description

            end if

        end if

    end if

end sub

'
' Parse the command line into its components
'
function ParseCommandLine(iAction, oParamDict)

    on error resume next

    DebugPrint kDebugTrace, L_Text_Dbg_Msg05_Text

    dim oArgs
    dim iIndex

    iAction = kActionUnknown

    set oArgs = Wscript.Arguments

    while iIndex < oArgs.Count

        select case oArgs(iIndex)

            case "-g"
                iAction = kActionGet

            case "-t"
                iAction = kActionSet

            case "-a"
                iAction = kActionAdd

            case "-d"
                iAction = kActionDelete

            case "-l"
                iAction = kActionList

            case "-2e"
                oParamDict.Add kDoubleSpool, true

            case "-2d"
                oParamDict.Add kDoubleSpool, false

            case "-s"
                iIndex = iIndex + 1
                oParamDict.Add kServerName, RemoveBackslashes(oArgs(iIndex))

            case "-u"
                iIndex = iIndex + 1
                oParamDict.Add kUserName, oArgs(iIndex)

            case "-w"
                iIndex = iIndex + 1
                oParamDict.Add kPassword, oArgs(iIndex)

            case "-n"
                iIndex = iIndex + 1
                oParamDict.Add kPortNumber, oArgs(iIndex)

            case "-r"
                iIndex = iIndex + 1
                oParamDict.Add kPortName, oArgs(iIndex)

            case "-o"
                iIndex = iIndex + 1
                oParamDict.Add kPortType, oArgs(iIndex)

            case "-h"
                iIndex = iIndex + 1
                oParamDict.Add kHostAddress, oArgs(iIndex)

            case "-q"
                iIndex = iIndex + 1
                oParamDict.Add kQueueName, oArgs(iIndex)

            case "-i"
                iIndex = iIndex + 1
                oParamDict.Add kSNMPDeviceIndex, oArgs(iIndex)

            case "-y"
                iIndex = iIndex + 1
                oParamDict.Add kCommunityName, oArgs(iIndex)

            case "-me"
                oParamDict.Add kSNMP, true

            case "-md"
                oParamDict.Add kSNMP, false

            case "-?"
                Usage(True)
                exit function

            case else
                Usage(True)
                exit function

        end select

        iIndex = iIndex + 1

    wend

    if Err = kErrorSuccess then

        ParseCommandLine = kErrorSuccess

    else

        wscript.echo L_Text_Error_General02_Text & L_Space_Text & L_Error_Text & L_Space_Text _
                     & L_Hex_Text & hex(Err.Number) & L_Space_text & Err.Description


        ParseCommandLine = kErrorFailure

    end if

end  function

'
' Display command usage.
'
sub Usage(bExit)

    wscript.echo L_Help_Help_General01_Text
    wscript.echo L_Help_Help_General02_Text
    wscript.echo L_Help_Help_General03_Text
    wscript.echo L_Help_Help_General04_Text
    wscript.echo L_Help_Help_General05_Text
    wscript.echo L_Help_Help_General06_Text
    wscript.echo L_Help_Help_General07_Text
    wscript.echo L_Help_Help_General08_Text
    wscript.echo L_Help_Help_General09_Text
    wscript.echo L_Help_Help_General10_Text
    wscript.echo L_Help_Help_General11_Text
    wscript.echo L_Help_Help_General12_Text
    wscript.echo L_Help_Help_General13_Text
    wscript.echo L_Help_Help_General14_Text
    wscript.echo L_Help_Help_General15_Text
    wscript.echo L_Help_Help_General16_Text
    wscript.echo L_Help_Help_General17_Text
    wscript.echo L_Help_Help_General18_Text
    wscript.echo L_Help_Help_General19_Text
    wscript.echo L_Help_Help_General20_Text
    wscript.echo L_Help_Help_General21_Text
    wscript.echo L_Help_Help_General22_Text
    wscript.echo L_Empty_Text
    wscript.echo L_Help_Help_General23_Text
    wscript.echo L_Help_Help_General24_Text
    wscript.echo L_Help_Help_General25_Text
    wscript.echo L_Help_Help_General26_Text
    wscript.echo L_Help_Help_General27_Text
    wscript.echo L_Help_Help_General28_Text
    wscript.echo L_Help_Help_General29_Text
    wscript.echo L_Empty_Text
    wscript.echo L_Help_Help_General30_Text
    wscript.echo L_Help_Help_General31_Text
    wscript.echo L_Help_Help_General32_Text

    if bExit then

        wscript.quit(1)

    end if

end sub

'
' Determines which program is being used to run this script.
' Returns true if the script host is cscript.exe
'
function IsHostCscript()

    on error resume next

    dim strFullName
    dim strCommand
    dim i, j
    dim bReturn

    bReturn = false

    strFullName = WScript.FullName

    i = InStr(1, strFullName, ".exe", 1)

    if i <> 0 then

        j = InStrRev(strFullName, "\", i, 1)

        if j <> 0 then

            strCommand = Mid(strFullName, j+1, i-j-1)

            if LCase(strCommand) = "cscript" then

                bReturn = true

            end if

        end if

    end if

    if Err <> 0 then

        wscript.echo L_Text_Error_General01_Text & L_Space_Text & L_Error_Text & L_Space_Text _
                     & L_Hex_Text & hex(Err.Number) & L_Space_Text & Err.Description

    end if

    IsHostCscript = bReturn

end function

'
' Retrieves extended information about the last error that occurred
' during a WBEM operation. The methods that set an SWbemLastError
' object are GetObject, PutInstance, DeleteInstance
'
sub LastError()

    on error resume next

    dim oError

    set oError = CreateObject("WbemScripting.SWbemLastError")

    if Err = kErrorSuccess then

        wscript.echo L_Operation_Text            & L_Space_Text & oError.Operation
        wscript.echo L_Provider_Text             & L_Space_Text & oError.ProviderName
        wscript.echo L_Description_Text          & L_Space_Text & oError.Description
        wscript.echo L_Text_Error_General04_Text & L_Space_Text & oError.StatusCode

    end if

end sub

'
' Connects to the WMI service on a server. oService is returned as a service
' object (SWbemServices)
'
function WmiConnect(strServer, strNameSpace, strUser, strPassword, oService)

   on error resume next

   dim oLocator
   dim bResult

   oService = null

   bResult  = false

   set oLocator = CreateObject("WbemScripting.SWbemLocator")

   if Err = kErrorSuccess then

      set oService = oLocator.ConnectServer(strServer, strNameSpace, strUser, strPassword)

      if Err = kErrorSuccess then

          bResult = true

          oService.Security_.impersonationlevel = 3

          '
          ' Required to perform administrative tasks on the spooler service
          '
          oService.Security_.Privileges.AddAsString "SeLoadDriverPrivilege"

          Err.Clear

      else

          wscript.echo L_Text_Msg_General10_Text & L_Space_Text & L_Error_Text _
                       & L_Space_Text & L_Hex_Text & hex(Err.Number) & L_Space_Text _
                       & Err.Description

      end if

   else

       wscript.echo L_Text_Msg_General09_Text & L_Space_Text & L_Error_Text _
                    & L_Space_Text & L_Hex_Text & hex(Err.Number) & L_Space_Text _
                    & Err.Description

   end if

   WmiConnect = bResult

end function

'
' Remove leading "\\" from server name
'
function RemoveBackslashes(strServer)

    dim strRet

    strRet = strServer

    if Left(strServer, 2) = "\\" and Len(strServer) > 2 then

        strRet = Mid(strServer, 3)

    end if

    RemoveBackslashes = strRet

end function
