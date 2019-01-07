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
session.findById("wnd[0]/tbar[0]/okcd").text = "stwb_2"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tblSAPLS_TWB_HD0100_TAB_CONTROL/btnBUTTON[0,10]").setFocus
session.findById("wnd[0]/usr/tblSAPLS_TWB_HD0100_TAB_CONTROL/btnBUTTON[0,10]").press
session.findById("wnd[0]/tbar[1]/btn[29]").press
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").selectItem "Root","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").ensureVisibleHorizontalItem "Root","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").clickLink "Root","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[1]/shell").pressButton "TWB_CALC_D"
session.findById("wnd[0]/usr/cntlTREE1/shellcont/shell/shellcont[1]/shell[1]").selectItem "          1","&Hierarchy"
session.findById("wnd[0]/usr/cntlTREE1/shellcont/shell/shellcont[1]/shell[1]").ensureVisibleHorizontalItem "          1","&Hierarchy"
session.findById("wnd[0]/usr/cntlTREE1/shellcont/shell/shellcont[1]/shell[1]").clickLink "          1","&Hierarchy"
session.findById("wnd[0]/usr/cntlTREE1/shellcont/shell/shellcont[1]/shell[0]").pressButton "&EXPAND"
session.findById("wnd[0]/usr/cntlTREE1/shellcont/shell/shellcont[1]/shell[0]").pressButton "FLAT"
session.findById("wnd[0]/tbar[1]/btn[46]").press
session.findById("wnd[0]/tbar[1]/btn[33]").press
session.findById("wnd[1]/usr/lbl[1,3]").setFocus
session.findById("wnd[1]/usr/lbl[1,3]").caretPosition = 3
session.findById("wnd[1]").sendVKey 2
session.findById("wnd[0]/tbar[1]/btn[43]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/okcd").text = "/NSTWB_2"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tblSAPLS_TWB_HD0100_TAB_CONTROL/btnBUTTON[0,3]").setFocus
session.findById("wnd[0]/usr/tblSAPLS_TWB_HD0100_TAB_CONTROL/btnBUTTON[0,3]").press
session.findById("wnd[0]/tbar[1]/btn[29]").press
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").selectItem "Root","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").ensureVisibleHorizontalItem "Root","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").clickLink "Root","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[1]/shell").pressButton "TWB_CALC_D"
session.findById("wnd[0]/usr/cntlTREE1/shellcont/shell/shellcont[1]/shell[1]").selectItem "          1","&Hierarchy"
session.findById("wnd[0]/usr/cntlTREE1/shellcont/shell/shellcont[1]/shell[1]").ensureVisibleHorizontalItem "          1","&Hierarchy"
session.findById("wnd[0]/usr/cntlTREE1/shellcont/shell/shellcont[1]/shell[1]").clickLink "          1","&Hierarchy"
session.findById("wnd[0]/usr/cntlTREE1/shellcont/shell/shellcont[1]/shell[0]").pressButton "&EXPAND"
session.findById("wnd[0]/usr/cntlTREE1/shellcont/shell/shellcont[1]/shell[0]").pressButton "FLAT"
session.findById("wnd[0]/tbar[1]/btn[46]").press
session.findById("wnd[0]/tbar[1]/btn[33]").press
session.findById("wnd[1]/usr/lbl[1,3]").setFocus
session.findById("wnd[1]/usr/lbl[1,3]").caretPosition = 3
session.findById("wnd[1]").sendVKey 2
session.findById("wnd[0]/tbar[1]/btn[43]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/okcd").text = "/NSTWB_2"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tblSAPLS_TWB_HD0100_TAB_CONTROL/btnBUTTON[0,7]").setFocus
session.findById("wnd[0]/usr/tblSAPLS_TWB_HD0100_TAB_CONTROL/btnBUTTON[0,7]").press
session.findById("wnd[0]/tbar[1]/btn[29]").press
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").selectItem "Root","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").ensureVisibleHorizontalItem "Root","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").clickLink "Root","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[1]/shell").pressButton "TWB_CALC_D"
session.findById("wnd[0]/usr/cntlTREE1/shellcont/shell/shellcont[1]/shell[1]").selectItem "          1","&Hierarchy"
session.findById("wnd[0]/usr/cntlTREE1/shellcont/shell/shellcont[1]/shell[1]").ensureVisibleHorizontalItem "          1","&Hierarchy"
session.findById("wnd[0]/usr/cntlTREE1/shellcont/shell/shellcont[1]/shell[1]").clickLink "          1","&Hierarchy"
session.findById("wnd[0]/usr/cntlTREE1/shellcont/shell/shellcont[1]/shell[0]").pressButton "&EXPAND"
session.findById("wnd[0]/usr/cntlTREE1/shellcont/shell/shellcont[1]/shell[0]").pressButton "FLAT"
session.findById("wnd[0]/tbar[1]/btn[46]").press
session.findById("wnd[0]/tbar[1]/btn[33]").press
session.findById("wnd[1]/usr/lbl[1,3]").setFocus
session.findById("wnd[1]/usr/lbl[1,3]").caretPosition = 6
session.findById("wnd[1]").sendVKey 2
session.findById("wnd[0]/tbar[1]/btn[43]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/okcd").text = "/NSTWB_2"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tblSAPLS_TWB_HD0100_TAB_CONTROL/btnBUTTON[0,6]").setFocus
session.findById("wnd[0]/usr/tblSAPLS_TWB_HD0100_TAB_CONTROL/btnBUTTON[0,6]").press
session.findById("wnd[0]/tbar[1]/btn[29]").press
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").selectItem "Root","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").ensureVisibleHorizontalItem "Root","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").clickLink "Root","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[1]/shell").pressButton "TWB_CALC_D"
session.findById("wnd[0]/usr/cntlTREE1/shellcont/shell/shellcont[1]/shell[1]").selectItem "          1","&Hierarchy"
session.findById("wnd[0]/usr/cntlTREE1/shellcont/shell/shellcont[1]/shell[1]").ensureVisibleHorizontalItem "          1","&Hierarchy"
session.findById("wnd[0]/usr/cntlTREE1/shellcont/shell/shellcont[1]/shell[1]").clickLink "          1","&Hierarchy"
session.findById("wnd[0]/usr/cntlTREE1/shellcont/shell/shellcont[1]/shell[0]").pressButton "&EXPAND"
session.findById("wnd[0]/usr/cntlTREE1/shellcont/shell/shellcont[1]/shell[0]").pressButton "FLAT"
session.findById("wnd[0]/tbar[1]/btn[46]").press
session.findById("wnd[0]/tbar[1]/btn[33]").press
session.findById("wnd[1]/usr/lbl[1,3]").setFocus
session.findById("wnd[1]/usr/lbl[1,3]").caretPosition = 3
session.findById("wnd[1]").sendVKey 2
session.findById("wnd[0]/tbar[1]/btn[43]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/okcd").text = "/NSTWB_2"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tblSAPLS_TWB_HD0100_TAB_CONTROL/btnBUTTON[0,1]").setFocus
session.findById("wnd[0]/usr/tblSAPLS_TWB_HD0100_TAB_CONTROL/btnBUTTON[0,1]").press
session.findById("wnd[0]/tbar[1]/btn[29]").press
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").selectItem "Root","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").ensureVisibleHorizontalItem "Root","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").clickLink "Root","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[1]/shell").pressButton "TWB_CALC_D"
session.findById("wnd[0]/usr/cntlTREE1/shellcont/shell/shellcont[1]/shell[1]").selectItem "          1","&Hierarchy"
session.findById("wnd[0]/usr/cntlTREE1/shellcont/shell/shellcont[1]/shell[1]").ensureVisibleHorizontalItem "          1","&Hierarchy"
session.findById("wnd[0]/usr/cntlTREE1/shellcont/shell/shellcont[1]/shell[1]").clickLink "          1","&Hierarchy"
session.findById("wnd[0]/usr/cntlTREE1/shellcont/shell/shellcont[1]/shell[0]").pressButton "&EXPAND"
session.findById("wnd[0]/usr/cntlTREE1/shellcont/shell/shellcont[1]/shell[0]").pressButton "FLAT"
session.findById("wnd[0]/tbar[1]/btn[46]").press
session.findById("wnd[0]/tbar[1]/btn[33]").press
session.findById("wnd[1]/usr/lbl[1,3]").setFocus
session.findById("wnd[1]/usr/lbl[1,3]").caretPosition = 3
session.findById("wnd[1]").sendVKey 2
session.findById("wnd[0]/tbar[1]/btn[43]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/okcd").text = "/NSTWB_2"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tblSAPLS_TWB_HD0100_TAB_CONTROL/btnBUTTON[0,5]").setFocus
session.findById("wnd[0]/usr/tblSAPLS_TWB_HD0100_TAB_CONTROL/btnBUTTON[0,5]").press
session.findById("wnd[0]/tbar[1]/btn[29]").press
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").selectItem "Root","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").ensureVisibleHorizontalItem "Root","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[1]/shell").pressButton "TWB_CALC_D"
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").clickLink "Root","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[1]/shell").pressButton "TWB_CALC_D"
session.findById("wnd[0]/usr/cntlTREE1/shellcont/shell/shellcont[1]/shell[1]").selectItem "          1","&Hierarchy"
session.findById("wnd[0]/usr/cntlTREE1/shellcont/shell/shellcont[1]/shell[1]").ensureVisibleHorizontalItem "          1","&Hierarchy"
session.findById("wnd[0]/usr/cntlTREE1/shellcont/shell/shellcont[1]/shell[1]").clickLink "          1","&Hierarchy"
session.findById("wnd[0]/usr/cntlTREE1/shellcont/shell/shellcont[1]/shell[0]").pressButton "&EXPAND"
session.findById("wnd[0]/usr/cntlTREE1/shellcont/shell/shellcont[1]/shell[0]").pressButton "FLAT"
session.findById("wnd[0]/tbar[1]/btn[46]").press
session.findById("wnd[0]/tbar[1]/btn[33]").press
session.findById("wnd[1]/usr/lbl[1,3]").setFocus
session.findById("wnd[1]/usr/lbl[1,3]").caretPosition = 5
session.findById("wnd[1]").sendVKey 2
session.findById("wnd[0]/tbar[1]/btn[43]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/okcd").text = "/NSTWB_2"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tblSAPLS_TWB_HD0100_TAB_CONTROL/btnBUTTON[0,2]").setFocus
session.findById("wnd[0]/usr/tblSAPLS_TWB_HD0100_TAB_CONTROL/btnBUTTON[0,2]").press
session.findById("wnd[0]/tbar[1]/btn[29]").press
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").selectItem "Root","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").ensureVisibleHorizontalItem "Root","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[1]/shell").pressButton "TWB_CALC_D"
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").clickLink "Root","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[1]/shell").pressButton "TWB_CALC_D"
session.findById("wnd[0]/usr/cntlTREE1/shellcont/shell/shellcont[1]/shell[1]").selectItem "          1","&Hierarchy"
session.findById("wnd[0]/usr/cntlTREE1/shellcont/shell/shellcont[1]/shell[1]").ensureVisibleHorizontalItem "          1","&Hierarchy"
session.findById("wnd[0]/usr/cntlTREE1/shellcont/shell/shellcont[1]/shell[1]").clickLink "          1","&Hierarchy"
session.findById("wnd[0]/usr/cntlTREE1/shellcont/shell/shellcont[1]/shell[0]").pressButton "&EXPAND"
session.findById("wnd[0]/usr/cntlTREE1/shellcont/shell/shellcont[1]/shell[0]").pressButton "FLAT"
session.findById("wnd[0]/tbar[1]/btn[46]").press
session.findById("wnd[0]/tbar[1]/btn[33]").press
session.findById("wnd[1]/usr/lbl[1,3]").setFocus
session.findById("wnd[1]/usr/lbl[1,3]").caretPosition = 4
session.findById("wnd[1]").sendVKey 2
session.findById("wnd[0]/tbar[1]/btn[43]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/okcd").text = "/N"
session.findById("wnd[0]").sendVKey 0
