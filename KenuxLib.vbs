' Ends With Helper Function
Function EndsWith(String, Value)
	EndsWith = False
	If InStrRev(String, Value) = (Len(String) + 1 - Len(Value)) Then
		EndsWith = True
	End If
End Function

' Starts With Helper Function
Function StartsWith(String, Value)
	StartsWith = False
	If InStr(1, String, Value) = 1 Then
		StartsWith = True
	End If
End Function

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