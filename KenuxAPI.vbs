Dim KX_FSO 		: Set KX_FSO		= CreateObject("Scripting.FileSystemObject")
Dim KX_Shell	: Set KX_Shell		= CreateObject("WScript.Shell")

Function KX_EndsWith(String, Value)
	KX_EndsWith = False
	If InStrRev(String, Value) = (Len(String) + 1 - Len(Value)) Then
		KX_EndsWith = True
	End If
End Function

Function KX_StartsWith(String, Value)
	KX_StartsWith = False
	If InStr(1, String, Value) = 1 Then
		KX_StartsWith = True
	End If
End Function

Function KX_Log(String, Name)
	WScript.StdOut.WriteLine("["&FormatDateTime(Now, vbShortTime)&"] ["&Name&"]: "&String)
End Function

Class KX_IniFile
	Private IniPath
	
	Function SetPath(Path)
		If KX_FSO.FileExists(Path) Then
			IniPath = Path
		Else
			Err.Raise 600 + vbObjectError, "KenuxAPI->KX_IniFile->SetPath", "Path, provided for SetPath, points to an unexisting file."
		End If
	End Function
	
	Function ReadIni( Section, Key )
		Const ForReading   = 1
		Const ForWriting   = 2
		Const ForAppending = 8

		Dim intEqualPos, objIniFile
		Dim strFilePath, strKey, strLeftString, strLine, strSection

		ReadIni     = ""
		strFilePath = Trim( IniPath )
		strSection  = Trim( Section )
		strKey      = Trim( Key )

		If KX_FSO.FileExists( strFilePath ) Then
			Set objIniFile = KX_FSO.OpenTextFile( strFilePath, ForReading, False )
			
			Do While objIniFile.AtEndOfStream = False
				strLine = Trim( objIniFile.ReadLine )
				
				If LCase( strLine ) = "[" & LCase( strSection ) & "]" Then
					strLine = Trim( objIniFile.ReadLine )

					Do While Left( strLine, 1 ) <> "["
						intEqualPos = InStr( 1, strLine, "=", 1 )
						If intEqualPos > 0 Then
							strLeftString = Trim( Left( strLine, intEqualPos - 1 ) )
							
							If LCase( strLeftString ) = LCase( strKey ) Then
								ReadIni = Trim( Mid( strLine, intEqualPos + 1 ) )
								
								If ReadIni = "" Then
									ReadIni = " "
								End If
								
								Exit Do
							End If
						End If


						If objIniFile.AtEndOfStream Then Exit Do

						strLine = Trim( objIniFile.ReadLine )
					Loop
				Exit Do
				End If
			Loop
			objIniFile.Close
		Else
			Err.Raise 600 + vbObjectError, "KenuxAPI->KX_IniFile->ReadIni", "Path, used by ReadIni, points to an unexisting file."
		End If
	End Function
End Class