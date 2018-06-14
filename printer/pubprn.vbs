'----------------------------------------------------------------------
'    pubprn.vbs - publish printers from a non Windows 2000 server into Windows 2000 DS
'    
'
'     Arguments are:-
'        server - format server
'        DS container - format "LDAP:\\CN=...,DC=...."
'
'
'    Copyright (c) Microsoft Corporation 1997
'   All Rights Reserved
'----------------------------------------------------------------------

'--- Begin Error Strings ---

Dim L_PubprnUsage1_text
Dim L_PubprnUsage2_text
Dim L_PubprnUsage3_text      
Dim L_PubprnUsage4_text      
Dim L_PubprnUsage5_text      
Dim L_PubprnUsage6_text      

Dim L_GetObjectError1_text
Dim L_GetObjectError2_text

Dim L_PublishError1_text
Dim L_PublishError2_text     
Dim L_PublishError3_text
Dim L_PublishSuccess1_text


L_PubprnUsage1_text      =   "用法: [cscript] pubprn.vbs server ""LDAP://OU=..,DC=..."""
L_PubprnUsage2_text      =   "       server 是 Windows 服务器名 (例如: Server) 或者 UNC 打印机名 (\\Server\Printer)"
L_PubprnUsage3_text      =   "       ""LDAP://CN=...,DC=..."" 是目标容器的 DS 路径"
L_PubprnUsage4_text      =   ""
L_PubprnUsage5_text      =   "例 1: pubprn.vbs MyServer ""LDAP://CN=MyContainer,DC=MyDomain,DC=Company,DC=Com"""
L_PubprnUsage6_text      =   "例 2: pubprn.vbs \\MyServer\Printer ""LDAP://CN=MyContainer,DC=MyDomain,DC=Company,DC=Com"""

L_GetObjectError1_text   =   "错误: 路径"
L_GetObjectError2_text   =   " 没有找到。"
L_GetObjectError3_text   =   "错误: 无法访问"

L_PublishError1_text     =   "错误: Pubprn 无法公布打印机 -  "
L_PublishError2_text     =   " 因为运行 Windows 2000 或者以上版本。"
L_PublishError3_text     =   "公布打印机失败"
L_PublishError4_text     =   "错误: "
L_PublishSuccess1_text   =   "公布的打印机: "

'--- End Error Strings ---


set Args = Wscript.Arguments
if args.count < 2 then
    wscript.echo L_PubprnUsage1_text
    wscript.echo L_PubprnUsage2_text
    wscript.echo L_PubprnUsage3_text
    wscript.echo L_PubprnUsage4_text
    wscript.echo L_PubprnUsage5_text
    wscript.echo L_PubprnUsage6_text
    wscript.quit(1)
end if

ServerName= args(0)
Container = args(1)


on error resume next
Set PQContainer = GetObject(Container)

if err then
    wscript.echo L_GetObjectError1_text & Container & L_GetObjectError2_text
    wscript.quit(1)
end if
on error goto 0



if left(ServerName,1) = "\" then

    PublishPrinter ServerName, ServerName, Container

else

    on error resume next

    Set PrintServer = GetObject("WinNT://" & ServerName & ",computer")

    if err then
        wscript.echo L_GetObjectError3_text & ServerName & ": " & err.Description
        wscript.quit(1)
    end if

    on error goto 0


    For Each Printer In PrintServer
        if Printer.class = "PrintQueue" then PublishPrinter Printer.PrinterPath, ServerName, Container
    Next


end if




sub PublishPrinter(UNC, ServerName, Container)

    
    Set PQ = WScript.CreateObject("OlePrn.DSPrintQueue.1")

    PQ.UNCName = UNC
    PQ.Container = Container

    on error resume next

    PQ.Publish(2)

    if err then
        if err.number = -2147024772 then
            wscript.echo L_PublishError1_text & Chr(34) & ServerName & Chr(34) & L_PublishError2_text
            wscript.quit(1)
        else
            wscript.echo L_PublishError3_text & Chr(34) & UNC & Chr(34) & "."
            wscript.echo L_PublishError4_text & err.Description
        end if
    else
        wscript.echo L_PublishSuccess1_text & PQ.Path
    end if

    Set PQ = nothing

end sub
