Dim Msg, Msg1, Msg2, Style, Title, Help, Ctxt, Response, MyString, Tempo, ConfigParam
Set objShell = CreateObject("WScript.Shell")
Msg = "Deseja programar um desligamento?" & vbCrLf & _
      "Obs: Cancel - Cancela o programa definido anteriormente"    ' Define message.
Style = vbYesNoCancel or vbInformation or vbDefaultButton1 ' Define buttons.
Title = "Programar desligamento windows"    ' Define title.
Help = "DEMO.HLP"    ' Define Help file.
Ctxt = 1000    ' Define topic context. 
Msg1 = "Quantos minutos deseja?"
Msg2 = "Configurando tempo para desligar"


Response = MsgBox(Msg, Style, Title)
If Response = vbYes Then    ' User chose Yes.
	ConfigParam = InputBox(Msg1,Msg2,10)
	Tempo = ConfigParam * 60
	strCommand = "shutdown /s /f /t " & Tempo
	' Teste = MsgBox(strCommand , vbOKOnly, "Resultado")
	objShell.Exec strCommand
ElseIF Response = vbCancel Then ' User chose No.
	strCommand = "shutdown /a"
	' Teste = MsgBox(strCommand , vbOKOnly, "Resultado")
	objShell.Exec strCommand
	MsgBox "Programa cancelado!"
Else
    MsgBox "Obrigado e at" & Chr(233) & " a pr" & Chr(243) & "xima!"
End If