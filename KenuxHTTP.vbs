Option Explicit

' include kenuxlib
include "KenuxLib.vbs"
include "KenuxAPI.vbs"

' define variables
Dim FSO 			: Set FSO 			= WScript.CreateObject("Scripting.FileSystemObject")
Dim KenuxHTTPVer	: KenuxHTTPVer		= "1.0"
Dim bClose 			: bClose 			= True
Dim Port 			: Port 				= 6789
Dim WinSock 		: Set WinSock 		= WScript.CreateObject("MSWinsock.Winsock")

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
	Dim Data, Path, PData, PathSrv, ConType, GetParams
	
	' content type, plain text for default
	ConType = "text/plain"
	
	' recieve client data
	Data = ""
	Winsock.GetData Data, 8, bytTotal
	
	' get the client request lines
	PData = Split(Data, vbCrlf)
	
	' get the path
	Path 		= Split(Split(PData(0), " ")(1), "?")(0)	' full path of internal static file
	PathSrv 	= Split(Split(PData(0), " ")(1), "?")(0)	' path from client
	If (UBound(Split(Split(PData(0), " ")(1), "?")) + 1) = 2 Then
		Set GetParams 	= CreateGetParams("?" & Split(Split(PData(0), " ")(1), "?")(1))	' create get params from url
	Else
		Set GetParams 	= CreateObject("Scripting.Dictionary")
	End If
	' if ends with /, add index.html
	' for example:
	' before: /somedir/
	' after:  /somedir/index.html
	If KX_EndsWith(Path, "/") Then
		Path = Path & "index.html"
	End If
	
	' if starts with /, add path to WWW directory
	If KX_StartsWith(Path, "/") Then
		Path = FSO.GetParentFolderName(WScript.ScriptFullName) & "/www" & Path
	End If
	
	' console info
	LogEvent("Получено: "+PData(0))
	LogEventB("Путь: " + PathSrv)
	
	' if file exists then
	If FSO.FileExists(Path) Then
		' read file content
		Dim Content : Content = FSO.OpenTextFile(Path, 1).ReadAll()
		
		' if ends with .html, set content type to html
		If KX_EndsWith(Path, ".html") Then
			ConType = "text/html"
			' and send it to the socket (prepared by HTTPCorrectData)
			Winsock.SendData(HTTPCorrectData(Content, ConType, "200 OK"))
		' If it is VBS/KenuxHTTP file
		ElseIf KX_EndsWith(Path, ".vbk") Then
			ConType = "text/html"
			include Path
			Dim retcon : retcon = VBKPage_Main(Array(GetParams, PathSrv))
			Winsock.SendData(HTTPCorrectData(retcon, ConType, "200 OK"))
		Else
			Winsock.SendData(HTTPCorrectData(Content, ConType, "200 OK"))
		End If

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