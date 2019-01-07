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
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        818","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        818","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        818","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "AC06"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0

session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        818","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        818","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        818","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "AC06"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        819","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        819","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        819","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "CS03"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        820","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        820","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        820","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "CS03"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        821","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        821","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        821","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "F.01"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        822","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        822","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        822","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "F.03"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        823","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        823","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        823","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "F.08"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        824","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        824","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        824","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "F.13"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        825","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        825","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        825","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "F.98"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        826","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        826","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        826","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "F-03"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        827","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        827","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        827","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "F-05"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        828","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        828","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        828","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "F-43"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        829","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        829","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        829","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "F-44"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        830","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        830","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        830","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "F-47"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        831","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        831","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        831","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "F-51"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        832","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        832","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        832","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "F-63"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        833","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        833","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        833","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "F-65"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        834","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        834","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        834","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "FB00"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        835","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        835","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        835","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "FB02"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        836","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        836","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        836","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "FB03"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        837","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        837","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        837","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "FB08"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        838","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        838","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        838","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "FB60"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        839","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        839","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        839","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "FB65"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        840","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        840","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        840","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "FBCJ"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        841","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        841","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        841","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "FBL1"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        842","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        842","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        842","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "FBL1N"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        843","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        843","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        843","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "FBL2N"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        844","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        844","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        844","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "FBL3"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        845","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        845","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        845","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "FBL3N"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        846","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        846","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        846","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "FBL5"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        847","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        847","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        847","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "FBL5N"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        848","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        848","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        848","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "FBL6N"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        849","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        849","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        849","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "FBRA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        850","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        850","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        850","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "FBV0"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        851","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        851","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        851","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "FK03"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        852","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        852","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        852","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "FS03"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        853","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        853","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        853","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "FS10"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        854","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        854","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        854","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "FS10N"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        855","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        855","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        855","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "GB01"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        856","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        856","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        856","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "GB06"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        857","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        857","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        857","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "GCU1"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        858","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        858","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        858","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "GD13"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        859","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        859","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        859","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "GD20"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        860","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        860","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        860","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "GS02"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        861","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        861","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        861","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "KS03"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        862","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        862","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        862","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "KSB1"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        863","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        863","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        863","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "MB03"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        864","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        864","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        864","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "ME23N"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        865","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        865","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        865","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "ME2L"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        866","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        866","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        866","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "ME2M"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        867","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        867","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        867","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "ME2N"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        868","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        868","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        868","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "ME53N"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        869","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        869","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        869","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "ME80F"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        870","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        870","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        870","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "ME80FN"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        871","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        871","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        871","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "MIR6"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        872","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        872","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        872","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "MIRO"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        873","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        873","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        873","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "MR11"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        874","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        874","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        874","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "MR11SHOW"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        875","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        875","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        875","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "MR8M"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        876","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        876","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        876","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "OAWD"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        877","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        877","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        877","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "SM35"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        878","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        878","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        878","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "SQ00"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        879","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        879","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        879","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "SWIA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        880","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        880","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        880","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "XK03"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        881","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        881","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        881","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "ZB27"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        882","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        882","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        882","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "ZCA003"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        883","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        883","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        883","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "ZCA004"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        884","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        884","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        884","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "ZCA007"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        885","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        885","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        885","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "ZCOFI007"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        886","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        886","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        886","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "ZDIAC_15"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        887","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        887","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        887","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "ZDICO_15"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        888","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        888","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        888","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "ZDIDE_15"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        889","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        889","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        889","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "ZDIDEL_15"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        890","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        890","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        890","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "ZDIDT_15"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        891","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        891","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        891","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "ZDIEX_15"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        892","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        892","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        892","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "ZDIEXB_15"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        893","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        893","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        893","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "ZDIEXC_15"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        894","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        894","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        894","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "ZDIGE_15"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        895","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        895","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        895","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "ZDIIN_15"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        896","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        896","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        896","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "ZDIIV_15"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        897","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        897","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        897","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "ZDIPA_15"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        898","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        898","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        898","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "ZDIRE_15"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        899","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        899","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        899","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "ZDITE_15"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        900","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        900","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        900","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "ZDITO_15"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        901","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        901","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        901","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "ZDITT_15"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        902","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        902","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        902","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "ZDIVI_15"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        903","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        903","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        903","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "ZDIVP_15"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        904","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        904","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        904","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "ZDIVR_15"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        905","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        905","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        905","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "ZDIVZ_15"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        906","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        906","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        906","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "ZFI070"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        907","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        907","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        907","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "ZFI072"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        908","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        908","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        908","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "ZFI076"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        909","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        909","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        909","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "ZFI088"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        910","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        910","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        910","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "ZFI227"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        911","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        911","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        911","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "ZFI234"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        912","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        912","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        912","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "ZFI242"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        913","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        913","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        913","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "ZFIDE"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        914","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        914","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        914","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "ZMM86"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        915","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        915","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        915","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "ZMMR033"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        916","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        916","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        916","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "ZMMR040"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        917","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        917","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        917","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "ZMMR056"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        918","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        918","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        918","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "ZVIM_EXCEL"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "        919","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        919","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "        919","TEXT"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,0]").key = "BMTA"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").text = "ZVIM_PRIORIDAD_ICC"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,0]").caretPosition = 0

