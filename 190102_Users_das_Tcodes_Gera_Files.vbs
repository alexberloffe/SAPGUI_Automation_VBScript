'Programa: Users_das_Tcodes_Gera_Files.vbs
'Objetivo: Faz a leitura dos usuários de cada transação, uma por arquivo
'Data: 02/01/2019


'Carga Excel para SAP via Script
'Na transação xyz, pela transação mno buscar os users
'Alex Berloffe - 03/10/2018

'Em caso de erros, processa a próxima transação

'On Error Resume Next
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
session.findById("wnd[0]/tbar[0]/okcd").text = "S_BCE_68001400"
session.findById("wnd[0]").sendVKey 0

'Lê a transação no Excel

dim x,arrStr,ForReading
dim js, objTextFile
set js=CreateObject("Scripting.FileSystemObject")
set objTextFile = js.OpenTextFile("C:\Users\GLB00156\Documents\00000\Tcodes.txt")

Do while NOT objTextFile.AtEndOfStream
	arrStr         = Split(objTextFile.ReadLine,";")
	TCode	       = arrStr(0)

session.findById("wnd[0]/usr/tabsTABSTRIP_TAB/tabpTAB1/ssub%_SUBSCREEN_TAB:RSUSR002:1001/ctxtFDATE").text = "01012018"
session.findById("wnd[0]/usr/tabsTABSTRIP_TAB/tabpTAB1/ssub%_SUBSCREEN_TAB:RSUSR002:1001/ctxtTDATE").text = "31129999"
session.findById("wnd[0]/usr/tabsTABSTRIP_TAB/tabpTAB1/ssub%_SUBSCREEN_TAB:RSUSR002:1001/ctxtTDATE").setFocus
session.findById("wnd[0]/usr/tabsTABSTRIP_TAB/tabpTAB1/ssub%_SUBSCREEN_TAB:RSUSR002:1001/ctxtTDATE").caretPosition = 8
session.findById("wnd[0]/usr/tabsTABSTRIP_TAB/tabpTAB3").select
session.findById("wnd[0]/usr/tabsTABSTRIP_TAB/tabpTAB3/ssub%_SUBSCREEN_TAB:RSUSR002:1003/ctxtTCODE").text = Tcode
session.findById("wnd[0]/usr/tabsTABSTRIP_TAB/tabpTAB3/ssub%_SUBSCREEN_TAB:RSUSR002:1003/ctxtTCODE").setFocus
session.findById("wnd[0]/usr/tabsTABSTRIP_TAB/tabpTAB3/ssub%_SUBSCREEN_TAB:RSUSR002:1003/ctxtTCODE").caretPosition = 4
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").select
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\GLB00156\Documents\00000\"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = Tcode & ".txt"
session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus
session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 33
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

Loop

objTextFile.Close
set objTextFile = Nothing
set js = Nothing
set x = 0
set arrStr = nothing

session.findById("wnd[0]/tbar[0]/btn[3]").press
