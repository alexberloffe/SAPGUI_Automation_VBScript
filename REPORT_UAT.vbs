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
session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").selectedNode = "0000000010"
session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode "0000000010"
session.findById("wnd[0]/usr/cntlSOLAR_EVAL/shellcont/shell").selectItem "         45","1"
session.findById("wnd[0]/usr/cntlSOLAR_EVAL/shellcont/shell").ensureVisibleHorizontalItem "         45","1"
session.findById("wnd[0]/usr/cntlSOLAR_EVAL/shellcont/shell").topNode = "         30"
session.findById("wnd[0]/usr/cntlSOLAR_EVAL/shellcont/shell").clickLink "         45","1"
session.findById("wnd[0]/usr/ctxtID").text = "z_dtv_unic"
session.findById("wnd[0]/usr/ctxtID").caretPosition = 10
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/ctxtP_LAYOUT").setFocus
session.findById("wnd[0]/usr/ctxtP_LAYOUT").caretPosition = 0
session.findById("wnd[0]").sendVKey 4
session.findById("wnd[1]/usr/cntlGRID/shellcont/shell").clickCurrentCell
session.findById("wnd[0]/usr/chkS_REFRSH").setFocus
session.findById("wnd[0]/usr/chkS_REFRSH").selected = true
session.findById("wnd[0]/usr/radS_R_OPT2").select
session.findById("wnd[0]/usr/chkM_REFRSH").setFocus
session.findById("wnd[0]/usr/chkM_REFRSH").selected = true
session.findById("wnd[0]/usr/radM_R_OPT2").select
session.findById("wnd[0]/usr/chkP_NOTCTX").selected = true
session.findById("wnd[0]/usr/chkP_NOTCTX").setFocus
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[32]").press
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "0"
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell
session.findById("wnd[0]/mbar/menu[0]/menu[2]/menu[1]").select
session.findById("wnd[1]/tbar[0]/btn[0]").press
