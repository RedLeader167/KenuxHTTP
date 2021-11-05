On Error Resume Next

Set WinSock = CreateObject("MSWinSock.WinSock")

If Err.Number <> 0 Then
  MsgBox "Произошла ошибка при проверке WinSock: " & Err.Description & vbCrlf & "Похоже, что класс не лицензирован. Запустить лицензацию? (бесплатно)", vbYesNo + vbExclamation, "KenuxHTTP"
  Set WShell = CreateObject("WScript.Shell")
  WShell.Run "license.reg", 1, True
End If