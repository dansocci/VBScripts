Dim MsgShutdown, Msg1, Msg2, StyleShutdown, TitleShutdown, Help, Ctxt, Shutdown, MyString, Tempo, ConfigParam
Set objShell = CreateObject("WScript.Shell")
MsgShutdown = "Deseja programar um desligamento?" & vbCrLf & _
      "Obs: Cancel - Cancela o programa definido anteriormente"    ' Define message.
StyleShutdown = vbYesNoCancel or vbInformation or vbDefaultButton1 ' Define buttons.
TitleShutdown = "Programar desligamento windows"    ' Define TitleShutdown.
Help = "DEMO.HLP"    ' Define Help file.
Ctxt = 1000    ' Define topic context. 
Msg1 = "Quantos minutos deseja?"
Msg2 = "Configurando tempo para desligar"



Shutdown = MsgBox(MsgShutdown, StyleShutdown, TitleShutdown)
If Shutdown = vbYes Then    ' User chose Yes.
	ConfigParam = InputBox(Msg1,Msg2,10)
	If ConfigParam > 0 Then
		Tempo = ConfigParam * 60
		strCommand = "shutdown /s /f /t " & Tempo
		' Teste = MsgBox(strCommand , vbOKOnly, "Resultado")
		objShell.Exec strCommand
		MsgBox "Computador programado para desligar em " & ConfigParam & " minutos."
	Else
		MsgBox "Obrigado e at" & Chr(233) & " a pr" & Chr(243) & "xima!"
	End If
ElseIF Shutdown = vbCancel Then ' User chose No.
	strCommand = "shutdown /a"
	' Teste = MsgBox(strCommand , vbOKOnly, "Resultado")
	objShell.Exec strCommand
	MsgBox "Programa cancelado!"
Else
  	MsgBox "Obrigado e at" & Chr(233) & " a pr" & Chr(243) & "xima!"
End If
