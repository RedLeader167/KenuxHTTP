Option Explicit

Dim FSO 			: Set FSO 			= WScript.CreateObject("Scripting.FileSystemObject")
Dim KenuxHTTPVer	: KenuxHTTPVer		= "1.0"
Dim bClose 			: bClose 			= True

Sub Winsock_ConnectionRequest(reqId)
	Winsock.Close()
	Winsock.Accept(reqId)
End Sub

Function EndsWith1(String, Value)
	Dim String1
	EndsWith = False
	If Len(Value) > Len(String) Then
		Exit Function
	End If
	String1 = Right(String, Len(String) + 1 - Len(Value))
	If String1 = Value Then
		EndsWith = True
	End If
End Function

Function EndsWith(String, Value)
	EndsWith = False
	'MsgBox InStrRev(String, Value) & ", " & (Len(String) + 1 - Len(Value)) & ", " & String & ", " & Value
	If InStrRev(String, Value) = (Len(String) + 1 - Len(Value)) Then
		EndsWith = True
	End If
End Function

Function StartsWith(String, Value)
	StartsWith = False
	If InStr(1, String, Value) = 1 Then
		StartsWith = True
	End If
End Function

Sub LogEvent(Tx)
	WScript.StdOut.WriteLine("[Сервер] " & Tx)
End Sub

Sub LogEventA(Tx)
	WScript.StdOut.WriteLine("    \--> " & Tx)
End Sub

Sub LogEventB(Tx)
	WScript.StdOut.WriteLine("    |")
	WScript.StdOut.WriteLine("    \--> " & Tx)
End Sub

Function ConfigureWinSock(WS, Port)
	WScript.ConnectObject WS,"Winsock_"
	WS.LocalPort = Port
	WS.Protocol = 0
	ConfigureWinSock = WS
End Function

Function HTTPCorrectData(Content, ContentType, Status)
	Dim outdata : outdata = ""
	LogEventA(Status)
	outdata = outdata & "HTTP/1.0 " & Status & vbCrlf
	'outdata = outdata & "Date: Fri, 31 Dec 1999 23:59:59 GMT" & vbCrlf
	outdata = outdata & "Server: KenuxHTTP/" & KenuxHTTPVer & vbCrlf
	outdata = outdata & "Content-Type: " & ContentType & vbCrlf
	outdata = outdata & "Content-Length: " & Len(Content) & vbCrlf
	'outdata = outdata & "Expires: Sat, 01 Jan 2000 00:59:59 GMT" & vbCrlf
	'outdata = outdata & "Last-modified: Fri, 09 Aug 1996 14:21:40 GMT" & vbCrlf
	outdata = outdata & vbCrlf
	outdata = outdata & Content
	HTTPCorrectData = outdata
End Function

Sub Winsock_DataArrival(bytTotal)
	Dim Data, Path, PData, PathSrv, ConType
	
	ConType = "text/plain"
	Data = ""
	Winsock.GetData Data, 8, bytTotal
	
	PData = Split(Data, vbCrlf)
	Path = Split(PData(0), " ")(1)
	PathSrv = Split(PData(0), " ")(1)
	If EndsWith(Path, "/") Then
		Path = Path & "index.html"
	End If
	If StartsWith(Path, "/") Then
		Path = FSO.GetParentFolderName(WScript.ScriptFullName) & "/www" & Path
	End If
	If EndsWith(Path, ".html") Then
		ConType = "text/html"
	End If
	
	LogEvent("Получено: "+PData(0))
	LogEventB("Путь: " + PathSrv)
	
	If FSO.FileExists(Path) Then
		Dim Content : Content = FSO.OpenTextFile(Path, 1).ReadAll()
		Winsock.SendData(HTTPCorrectData(Content, ConType, "200 OK"))
	Else
		'LogEvent("Файл не существует, возвращаю 404")
		Winsock.SendData(HTTPCorrectData("<title>404 Not Found</title><h1>404 Not Found</h1><p>The requested file was not found on the server.</p><hr><center>Powered by KenuxHTTP " & KenuxHTTPVer & "</center>", "text/html", "404 Not Found"))
	End If
End Sub

Sub Winsock_Error (number, desc, sCode, src, help, helpctx, cancelDisplay)
	LogEvent("Ошибка: "+desc)
	bClose = False
End Sub

Sub Winsock_Close
	bClose=False
End Sub
Dim Port : Port = 6789

Dim WinSock : Set WinSock = WScript.CreateObject("MSWinsock.Winsock")
ConfigureWinSock WinSock, Port
LogEvent("Сервер запущен.")
LogEventB("Порт: " & Port)

While True
	Winsock.Listen
	   
	Do While (bClose) 
		WScript.Sleep(500)
	Loop
	
	bClose = True
	Set WinSock = WScript.CreateObject("MSWinsock.Winsock")
	ConfigureWinSock WinSock, Port
Wend