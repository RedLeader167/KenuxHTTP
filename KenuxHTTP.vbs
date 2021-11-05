Option Explicit

' define variables
Dim FSO 			: Set FSO 			= WScript.CreateObject("Scripting.FileSystemObject")
Dim KenuxHTTPVer	: KenuxHTTPVer		= "1.0"
Dim bClose 			: bClose 			= True
Dim Port 			: Port 				= 6789
Dim WinSock 		: Set WinSock 		= WScript.CreateObject("MSWinsock.Winsock")

' include kenuxlib
include "KenuxLib.vbs"

' configure socket
ConfigureWinSock WinSock, Port

' including routine
Sub include(fSpec)
    With CreateObject("Scripting.FileSystemObject")
       executeGlobal .openTextFile(fSpec).readAll()
    End With
End Sub

' idk really copypasted from forums
Sub Winsock_ConnectionRequest(reqId)
	Winsock.Close()
	Winsock.Accept(reqId)
End Sub

' handle request
Sub Winsock_DataArrival(bytTotal)
	Dim Data, Path, PData, PathSrv, ConType
	
	' content type, plain text for default
	ConType = "text/plain"
	
	' recieve client data
	Data = ""
	Winsock.GetData Data, 8, bytTotal
	
	' get the client request lines
	PData = Split(Data, vbCrlf)
	
	' get the path
	Path = Split(PData(0), " ")(1)		' full path of internal static file
	PathSrv = Split(PData(0), " ")(1)	' path from client
	
	' if ends with /, add index.html
	' for example:
	' before: /somedir/
	' after:  /somedir/index.html
	If EndsWith(Path, "/") Then
		Path = Path & "index.html"
	End If
	
	' if starts with /, add path to WWW directory
	If StartsWith(Path, "/") Then
		Path = FSO.GetParentFolderName(WScript.ScriptFullName) & "/www" & Path
	End If
	
	' if ends with .html, set content type to html
	If EndsWith(Path, ".html") Then
		ConType = "text/html"
	End If
	
	' console info
	LogEvent("Получено: "+PData(0))
	LogEventB("Путь: " + PathSrv)
	
	' if file exists then
	If FSO.FileExists(Path) Then
		' read file content
		Dim Content : Content = FSO.OpenTextFile(Path, 1).ReadAll()
		
		' and send it to the socket (prepared by HTTPCorrectData)
		Winsock.SendData(HTTPCorrectData(Content, ConType, "200 OK"))

	' file not exists, send 404
	Else
		' send 404 Not Found error page
		Winsock.SendData(HTTPCorrectData("<title>404 Not Found</title><h1>404 Not Found</h1><p>The requested file was not found on the server.</p><hr><center>Powered by KenuxHTTP " & KenuxHTTPVer & "</center>", "text/html", "404 Not Found"))
	End If
End Sub

' on error
Sub Winsock_Error (number, desc, sCode, src, help, helpctx, cancelDisplay)
	' console info
	LogEvent("Ошибка: "+desc)
	
	' close connection
	bClose = False
End Sub

' on close
Sub Winsock_Close
	bClose=False
End Sub

' console info
LogEvent("Сервер запущен.")
LogEventB("Порт: " & Port)

' handle requests infinitely
While True
	' listen for connections
	Winsock.Listen
	
	' wait while request is handling
	Do While (bClose) 
		WScript.Sleep(500)
	Loop
	
	' close
	bClose = True
	
	' recreate socket
	' ---------------------------------------------------- MOST LIKELY PART TO LEAK MEMORY
	Set WinSock = WScript.CreateObject("MSWinsock.Winsock")
	ConfigureWinSock WinSock, Port
Wend