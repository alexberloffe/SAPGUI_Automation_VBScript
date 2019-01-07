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
session.findById("wnd[0]/tbar[0]/okcd").text = "SU01"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtUSR02-BNAME").text = "ACHAVARRO"
session.findById("wnd[0]/tbar[1]/btn[18]").press
session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpLOGO").select
session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpLOGO/ssubMAINAREA:SAPLSUU5:0101/pwdG_PASSWORD1").text = "********"
session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpLOGO/ssubMAINAREA:SAPLSUU5:0101/pwdG_PASSWORD2").text = "********"
session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpLOGO/ssubMAINAREA:SAPLSUU5:0101/pwdG_PASSWORD2").setFocus
session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpLOGO/ssubMAINAREA:SAPLSUU5:0101/pwdG_PASSWORD2").caretPosition = 8
session.findById("wnd[0]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[15]").press
