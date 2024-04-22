Option Explicit
Dim ie, ipf, radius, WshShell

Set ie = CreateObject("InternetExplorer.Application")
Set WshShell = WScript.CreateObject("WScript.Shell")

Sub WaitForLoad
Do While IE.Busy
WScript.Sleep 500
Loop
End Sub

Sub Find(x)
Set ipf = ie.Document.All.Item(x)
End Sub

radius = InputBox("Insert the Radius", "Radius")

If IsEmpty(radius) Then
WScript.Quit

Else
ie.Navigate "https://access-vpn01.vodafone.pt/vpn/index.html"

Call WaitForLoad

ie.Visible = True

Call Find("Enter user name")
ipf.Value = "LimaD"
Call Find("passwd")
ipf.Value = "Dan2231992"
Call Find("passwd1")
ipf.Value = radius
WScript.Sleep 500
Call Find("Log_On")
ipf.Click

End if