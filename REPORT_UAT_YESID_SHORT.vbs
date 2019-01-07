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
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        81","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        81","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        81","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "F.03"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0
session.findById("wnd[0]").sendVKey 0
