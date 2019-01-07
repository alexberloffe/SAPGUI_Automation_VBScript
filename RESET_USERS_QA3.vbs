'Carga Excel para SAP SOLMAN via Script
'Esse script reinicializa as senhas em QA3
'Alex Berloffe - 15/11/2018

'Conecta com o SAP

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

'Lê a transação no Excel

dim x,arrStr,status,p,c
dim fs,objTextFile
set fs=CreateObject("Scripting.FileSystemObject")
set objTextFile = fs.OpenTextFile("C:\guixt\scripts\USERS_ELVER_RESET_SOL.csv")

Do while NOT objTextFile.AtEndOfStream
	arrStr         = Split(objTextFile.ReadLine,";")
	SAPUSER        = arrStr(0)

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "SU01"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtSUID_ST_BNAME-BNAME").text = SAPUSER
session.findById("wnd[0]/usr/ctxtSUID_ST_BNAME-BNAME").caretPosition = 5
session.findById("wnd[0]/tbar[1]/btn[18]").press
session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpLOGO").select
session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpLOGO/ssubMAINAREA:SAPLSUID_MAINTENANCE:1101/pwdSUID_ST_NODE_PASSWORD_EXT-PASSWORD").text = "A1B2C3D4"
session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpLOGO/ssubMAINAREA:SAPLSUID_MAINTENANCE:1101/pwdSUID_ST_NODE_PASSWORD_EXT-PASSWORD2").text = "A1B2C3D4"
session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpLOGO/ssubMAINAREA:SAPLSUID_MAINTENANCE:1101/pwdSUID_ST_NODE_PASSWORD_EXT-PASSWORD2").setFocus
session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpLOGO/ssubMAINAREA:SAPLSUID_MAINTENANCE:1101/pwdSUID_ST_NODE_PASSWORD_EXT-PASSWORD2").caretPosition = 8
session.findById("wnd[0]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press


Loop

objTextFile.Close
set objTextFile = Nothing
set fs = Nothing

session.findById("wnd[0]/tbar[0]/btn[3]").press