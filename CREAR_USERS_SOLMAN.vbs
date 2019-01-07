'Carga Excel para SAP SOLMAN via Script
'Esse script cria as estruturas de teste do SOLMAN
'Alex Berloffe - 07/11/2018

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
set objTextFile = fs.OpenTextFile("C:\guixt\scripts\USERS_ELVER_SOL.CSV")

Do while NOT objTextFile.AtEndOfStream
	arrStr         = Split(objTextFile.ReadLine,";")
	SAPUSER        = arrStr(0)
	NOMBRE	       = arrStr(1)
	EMAIL	       = arrStr(2)


session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "SU01"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/txtUSR01-BNAME").text = "EQUIROGA"
session.findById("wnd[1]/usr/txtUSR02-BNAME").text = SAPUSER
session.findById("wnd[1]/tbar[0]/btn[5]").press
session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpLOGO/ssubMAINAREA:SAPLSUU5:0101/pwdG_PASSWORD1").text = "A1B2C3D4"
session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpLOGO/ssubMAINAREA:SAPLSUU5:0101/pwdG_PASSWORD2").text = "A1B2C3D4"
session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpLOGO/ssubMAINAREA:SAPLSUU5:0101/pwdG_PASSWORD2").setFocus
session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpLOGO/ssubMAINAREA:SAPLSUU5:0101/pwdG_PASSWORD2").caretPosition = 8
session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpADDR").select
session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpADDR/ssubMAINAREA:SAPLSZA5:0900/txtADDR3_DATA-NAME_FIRST").text = NOMBRE
session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpADDR/ssubMAINAREA:SAPLSZA5:0900/txtSZA5_D0700-SMTP_ADDR").text = EMAIL
session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpADDR/ssubMAINAREA:SAPLSZA5:0900/txtSZA5_D0700-SMTP_ADDR").setFocus
session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpADDR/ssubMAINAREA:SAPLSZA5:0900/txtSZA5_D0700-SMTP_ADDR").caretPosition = 16
session.findById("wnd[0]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[15]").press

Loop

objTextFile.Close
set objTextFile = Nothing
set fs = Nothing

session.findById("wnd[0]/tbar[0]/btn[3]").press