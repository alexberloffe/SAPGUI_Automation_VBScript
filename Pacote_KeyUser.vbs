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
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").selectItem "000003","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").ensureVisibleHorizontalItem "000003","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").clickLink "000003","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").selectItem "000004","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").ensureVisibleHorizontalItem "000004","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").clickLink "000004","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").selectItem "000005","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").ensureVisibleHorizontalItem "000005","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").clickLink "000005","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").selectItem "000006","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").ensureVisibleHorizontalItem "000006","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").clickLink "000006","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").selectItem "000007","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").ensureVisibleHorizontalItem "000007","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").clickLink "000007","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").selectItem "000008","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").ensureVisibleHorizontalItem "000008","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").clickLink "000008","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").selectItem "000009","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").ensureVisibleHorizontalItem "000009","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").clickLink "000009","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").selectItem "000010","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").ensureVisibleHorizontalItem "000010","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").clickLink "000010","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").selectItem "000011","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").ensureVisibleHorizontalItem "000011","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").clickLink "000011","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").selectItem "000012","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").ensureVisibleHorizontalItem "000012","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").clickLink "000012","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").selectItem "000013","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").ensureVisibleHorizontalItem "000013","Hier_Column1"
session.findById("wnd[0]/usr/cntlCC_COLUMN_TREE_PACK/shellcont/shell/shellcont[2]/shell").clickLink "000013","Hier_Column1"
