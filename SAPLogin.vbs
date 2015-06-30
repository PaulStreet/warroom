'The below section will create an SAP session.
set WshShell = CreateObject("WScript.Shell")
 Set proc = WshShell.Exec("C:\Program Files\SAP\FrontEnd\SAPgui\saplogon.exe")
            Do While proc.Status = 0
            WScript.Sleep 100
      Loop
   Set SapGui = GetObject("SAPGUI")
Set Appl = SapGui.GetScriptingEngine

''Deprecated alternate code, wait for 6 seconds
'Dim dteWait
'dteWait = DateAdd("s", 6, Now())
'Do Until (Now() > dteWait)
'Loop

'Wait for 5 seconds then press enter.
WScript.Sleep 5000
WshShell.SendKeys "{ENTER}"

''This commented section of code doesn't seem to work for me.
'Set Connection = Appl.Openconnection("Test SAP", True)
'Set session = Connection.Children(0)
'session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = "USERNAME"
'session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = "PASSWORD"
'session.findById("wnd[0]/usr/txtRSYST-LANGU").Text = "E"
'session.findById("wnd[0]").sendVKey 0

'The below code is what I can record once I have gotten to the SAP login for the target server.
If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "USERNAME"
session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "PASSWORD"
session.findById("wnd[0]/usr/pwdRSYST-BCODE").setFocus
session.findById("wnd[0]/usr/pwdRSYST-BCODE").caretPosition = 10
session.findById("wnd[0]").sendVKey 0
