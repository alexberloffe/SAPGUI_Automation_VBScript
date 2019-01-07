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
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").clickLink "000011","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[1]/shell").pressButton "TWB_APTU"
session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[0,24]").text = "JMANCHINELLI"
session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[0,24]").caretPosition = 12
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/chk[1,3]").selected = true
session.findById("wnd[1]/usr/chk[1,3]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").clickLink "000011","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[1]/shell").pressButton "TWB_APTU"
session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[0,24]").text = "MBAIGORRIA"
session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[0,24]").caretPosition = 10
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/chk[1,3]").selected = true
session.findById("wnd[1]/usr/chk[1,3]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
