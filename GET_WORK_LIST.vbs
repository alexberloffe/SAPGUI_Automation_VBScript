'Busca Lista de Trabalhos de usuarios numa lista GET_WORK_LIST.txt
'Esse script busca as listas de trabalhos assinadas a cada colaborador
'Alex Berloffe - 09/12/2018

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
set objTextFile = fs.OpenTextFile("C:\guixt\scripts\GET_WORK_LIST.txt")

Do while NOT objTextFile.AtEndOfStream
	arrStr         = Split(objTextFile.ReadLine,";")
	SAPUSER        = arrStr(0)

session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode "0000000009"
session.findById("wnd[0]/usr/tblSAPLS_TWB_HD0100_TAB_CONTROL").columns.elementAt(1).width = 80
session.findById("wnd[0]/mbar/menu[4]/menu[11]").select
session.findById("wnd[0]/usr/ctxtTESTER").text = SAPUSER
session.findById("wnd[0]/usr/ctxtTESTER").caretPosition = 8
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/mbar/menu[0]/menu[2]/menu[2]").select
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "WORKLIST_" & SAPUSER& ".txt"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 17
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[15]").press

Loop

objTextFile.Close
set objTextFile = Nothing
set fs = Nothing

session.findById("wnd[0]/tbar[0]/btn[15]").press

'===================================================



