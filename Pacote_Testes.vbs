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
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "          1","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "          1","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "          1","TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[0]/shell").pressButton "TREP"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").collapseNode "         27"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3643","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3643","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").topNode = "       3665"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3643","BOX",false
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3642","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3642","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3642","BOX",false
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3641","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3641","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3641","BOX",false
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3640","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3640","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3640","BOX",false
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3639","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3639","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3639","BOX",true
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3638","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3638","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3638","BOX",true
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3637","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3637","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3637","BOX",true
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3636","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3636","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3636","BOX",false
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3635","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3635","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3635","BOX",true
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3634","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3634","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3634","BOX",false
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3633","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3633","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3633","BOX",false
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3632","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3632","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").topNode = "       3662"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3632","BOX",false
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3631","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3631","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").topNode = "       3660"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3631","BOX",true
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3630","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3630","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3630","BOX",false
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3629","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3629","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3629","BOX",false
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3628","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3628","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").topNode = "       3657"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3628","BOX",false
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3627","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3627","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").topNode = "       3653"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3627","BOX",true
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3626","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3626","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3626","BOX",true
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3625","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3625","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3625","BOX",true
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3624","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3624","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").topNode = "       3650"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3624","BOX",false
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3623","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3623","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3623","BOX",false
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3622","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3622","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").topNode = "       3617"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3622","BOX",false
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3621","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3621","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3621","BOX",false
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3620","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3620","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3620","BOX",true
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3619","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3619","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").topNode = "       3614"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3619","BOX",false
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3618","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3618","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3618","BOX",true
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3576","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3576","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").topNode = "       3612"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3576","BOX",true
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3575","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3575","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3575","BOX",true
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3574","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3574","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3574","BOX",true
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3573","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3573","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3573","BOX",true
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3572","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3572","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3572","BOX",true
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3571","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3571","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3571","BOX",true
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3570","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3570","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3570","BOX",true
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3569","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3569","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3569","BOX",true
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3568","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3568","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").topNode = "       3605"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3568","BOX",false
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3567","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3567","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3567","BOX",false
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3566","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3566","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").topNode = "       3564"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3566","BOX",false
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3565","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3565","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").topNode = "       3597"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3565","BOX",false
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3564","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3564","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").topNode = "       3592"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3564","BOX",false
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3563","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3563","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").topNode = "       3587"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3563","BOX",false
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3562","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3562","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3562","BOX",true
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3561","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3561","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3561","BOX",true
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3560","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3560","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").topNode = "       3557"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3560","BOX",false
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3559","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3559","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3559","BOX",false
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3558","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3558","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").topNode = "       3578"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3558","BOX",false
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3557","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3557","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3557","BOX",false
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3556","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3556","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").topNode = "       3544"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3556","BOX",false
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3554","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3554","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3554","BOX",false
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3553","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3553","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").topNode = "       3548"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3553","BOX",false
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3552","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3552","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       3552","BOX",false
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").collapseNode "         23"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").collapseNode "       3272"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").collapseNode "       3273"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").collapseNode "       3274"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").collapseNode "       3275"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").collapseNode "       3276"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").collapseNode "       3360"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").collapseNode "       3361"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").collapseNode "       3362"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").collapseNode "       3374"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").collapseNode "       3442"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").collapseNode "       3443"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").collapseNode "       3444"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").collapseNode "       3445"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").collapseNode "       3446"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").collapseNode "       3447"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").collapseNode "       3448"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       3447","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       3447","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").topNode = "       2538"
