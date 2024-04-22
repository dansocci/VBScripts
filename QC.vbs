Option Explicit
Dim ie, qc, email, WshShell

Set ie = CreateObject("InternetExplorer.Application")
Set WshShell = WScript.CreateObject("WScript.Shell")

Sub WaitForLoad
Do While IE.Busy
WScript.Sleep 500
Loop
End Sub

Sub Find(x)
Set qc = ie.Document.All.Item(x)
End Sub

ie.Navigate "https://alm-prod.vodafone.com/"

Call WaitForLoad

ie.Visible = True

Call Find("discoveryHint")
qc.email = "danilo.lima@vodafone.com"