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
session.findById("wnd[0]/tbar[0]/okcd").text = "se"
session.findById("wnd[0]").sendVKey 82
session.findById("wnd[0]/tbar[0]/okcd").text = "se"
session.findById("wnd[0]").sendVKey 82
session.findById("wnd[0]/tbar[0]/okcd").text = "se"
session.findById("wnd[0]").sendVKey 82
session.findById("wnd[0]/tbar[0]/okcd").text = "se38"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").text = "Z_CARGA_FB09"
session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").caretPosition = 12
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
session.findById("wnd[0]").sendVKey 0
