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

Dim L

For L = 4 to 228

MsgBox L

session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellRow = L
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").clickCurrentCell
session.findById("wnd[0]/usr/subSUB_STATUS:SAPLSTATUS:0151/tabsTWB_STATUS_TABSTRIP/tabpTWB_STADM/ssubSUB_STATUS_ADMIN:SAPLSTATUS:0152/cmbD0100_ELEMENTS-STATUS").setFocus
session.findById("wnd[0]/usr/subSUB_STATUS:SAPLSTATUS:0151/tabsTWB_STATUS_TABSTRIP/tabpTWB_STADM/ssubSUB_STATUS_ADMIN:SAPLSTATUS:0152/cmbD0100_ELEMENTS-STATUS").key = "TEST_OK"
session.findById("wnd[0]/usr/subSUB_STATUS:SAPLSTATUS:0151/tabsTWB_STATUS_TABSTRIP/tabpTWB_STADM/ssubSUB_STATUS_ADMIN:SAPLSTATUS:0152/txtD0100_ELEMENTS-NOTE_TEXT").text = "Teste efetuado com sucesso em UAT1"
session.findById("wnd[0]/usr/subSUB_STATUS:SAPLSTATUS:0151/tabsTWB_STATUS_TABSTRIP/tabpTWB_STADM/ssubSUB_STATUS_ADMIN:SAPLSTATUS:0152/txtD0100_ELEMENTS-NOTE_TEXT").setFocus
session.findById("wnd[0]/usr/subSUB_STATUS:SAPLSTATUS:0151/tabsTWB_STATUS_TABSTRIP/tabpTWB_STADM/ssubSUB_STATUS_ADMIN:SAPLSTATUS:0152/txtD0100_ELEMENTS-NOTE_TEXT").caretPosition = 34
session.findById("wnd[0]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

Next
