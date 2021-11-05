' Logging
Sub LogEvent(Tx)
	WScript.StdOut.WriteLine("[Сервер] " & Tx)
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

Sub AddHeader (Header)
	ReDim Preserve cHeaders(UBound(cHeaders) + 1)
	cHeaders(UBound(cHeaders)) = Header
End Sub

Sub ClearHeaders()
	cHeaders = Array()
End Sub

Dim cHeaders : cHeaders = Array()
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
	' custom headers
	For Each header in cHeaders
		outdata = outdata & header & vbCrlf
	Next
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

