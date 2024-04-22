Option Explicit
Dim ie, ipf, WshShell

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

ie.Navigate "https://onccpp-sit-pt-1.vf-pt.internal.vodafone.com/oncfo/a"

Call WaitForLoad

ie.Visible = True

Call Find("__ns2087359418_username")
ipf.Value = "351910088825"
Call Find("__ns2087359418_password")
ipf.Value = "CelFocus13!!"
Call Find("__ns2087359418_loginBtn")
' ipf.Click