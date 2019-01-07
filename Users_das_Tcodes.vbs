'Carga Excel para SAP via Script
'Na transação xyz, pela transação mno buscar os users
'Alex Berloffe - 03/10/2018

'Em caso de erros, processa a próxima transação

On Error Resume Next

'set wshell = createObject("WScript.Shell")

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

'Abre a transação no SAP

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "S_BCE_68001400 "
session.findById("wnd[0]").sendVKey 0

'Lê a transação no Excel

dim x,arrStr,status,p
dim fs,objTextFile
set fs=CreateObject("Scripting.FileSystemObject")
set objTextFile = fs.OpenTextFile("C:\guixt\scripts\TCODES.csv")

Do while NOT objTextFile.AtEndOfStream
	arrStr         = Split(objTextFile.ReadLine,";")
	TCode	       = arrStr(0)

session.findById("wnd[0]/usr/tabsTABSTRIP_TAB/tabpTAB1/ssub%_SUBSCREEN_TAB:RSUSR002:1001/ctxtFDATE").text = "010118"
session.findById("wnd[0]/usr/tabsTABSTRIP_TAB/tabpTAB1/ssub%_SUBSCREEN_TAB:RSUSR002:1001/ctxtTDATE").text = "31129999"
session.findById("wnd[0]/usr/tabsTABSTRIP_TAB/tabpTAB1/ssub%_SUBSCREEN_TAB:RSUSR002:1001/ctxtTDATE").setFocus
session.findById("wnd[0]/usr/tabsTABSTRIP_TAB/tabpTAB1/ssub%_SUBSCREEN_TAB:RSUSR002:1001/ctxtTDATE").caretPosition = 8


session.findById("wnd[0]/usr/tabsTABSTRIP_TAB/tabpTAB3").select
session.findById("wnd[0]/usr/tabsTABSTRIP_TAB/tabpTAB3/ssub%_SUBSCREEN_TAB:RSUSR002:1003/ctxtTCODE").text = Tcode
session.findById("wnd[0]/usr/tabsTABSTRIP_TAB/tabpTAB3/ssub%_SUBSCREEN_TAB:RSUSR002:1003/ctxtTCODE").setFocus
session.findById("wnd[0]/usr/tabsTABSTRIP_TAB/tabpTAB3/ssub%_SUBSCREEN_TAB:RSUSR002:1003/ctxtTCODE").caretPosition = 4
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[33]").press
session.findById("wnd[1]/usr/cntlGRID/shellcont/shell").clickCurrentCell
session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").select
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = Tcode & ".txt"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

Const ForReading = 1

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objOutputFile = objFSO.CreateTextFile("C:\Users\GLB00156\AppData\Local\SAP\SAP GUI\tmp\All_Tcodes.txt")
Set objTextFile = objFSO.OpenTextFile("C:\Users\GLB00156\AppData\Local\SAP\SAP GUI\tmp\" & Tcode & ".txt", ForReading)
strText = objTextFile.ReadAll
objTextFile.Close
objOutputFile.WriteLine strText
objOutputFile.Close


Loop

objTextFile.Close
set objTextFile = Nothing
set fs = Nothing

session.findById("wnd[0]/tbar[0]/btn[3]").press
