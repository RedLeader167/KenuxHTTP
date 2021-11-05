' Logging
Sub LogEvent(Tx)
	WScript.StdOut.WriteLine("[������] " & Tx)
End Sub
Sub LogEventA(Tx)
	WScript.StdOut.WriteLine("    \--> " & Tx)
End Sub
Sub LogEventB(Tx)
	WScript.StdOut.WriteLine("    |" & vbCrlf & "    \--> " & Tx)
End Sub

' Configure socket
Function ConfigureWinSock(WS, Port)
	' Connect all Winsock_* methods to socket
	WScript.ConnectObject WS,"Winsock_"
	
	' Configure it
	WS.LocalPort = Port
	WS.Protocol = 0
	
	' And return
	ConfigureWinSock = WS
End Function

' Create valid HTTP response
Function HTTPCorrectData(Content, ContentType, Status)
	Dim outdata : outdata = ""
	
	' log status
	LogEventA(Status)
	
	' http version and status
	outdata = outdata & "HTTP/1.0 " & Status & vbCrlf
	' server name
	outdata = outdata & "Server: KenuxHTTP/" & KenuxHTTPVer & vbCrlf
	' content type and length
	outdata = outdata & "Content-Type: " & ContentType & vbCrlf
	outdata = outdata & "Content-Length: " & Len(Content) & vbCrlf
	' splitting headers and content
	outdata = outdata & vbCrlf
	' content itself
	outdata = outdata & Content
	
	' return it
	HTTPCorrectData = outdata
End Function

' Create Get request parameters
Function CreateGetParams(URL)
    dim URL2,a,t,a2,dic
	Set dic = CreateObject("Scripting.Dictionary")
    if len(URL)<>0 then
        URL2=right(URL,len(URL)-1)
        a=split(URL2,"&")
        for i=0 to UBound(a)
            a2=split(a(i),"=")
            dic.Add a2(0), a2(1)
        next
    end if
	'For each key in dic.Keys
	'	WScript.Echo key
	'	WScript.Echo dic(key)
	'Next
	Set CreateGetParams = dic
End Function

sub getxy
    dim s,a,t,a2
    s=window.location.search
    if len(s)<>0 then
        s=right(s,len(s)-1)
        a=split(s,"&")
        for i=0 to 1
            a2=split(a(i),"=")
            document.getElementById("coord" & i).innerHTML="name: " & a2(0) & "; value: " & a2(1)
        next
    end if
end sub