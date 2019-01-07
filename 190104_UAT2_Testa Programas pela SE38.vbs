'Programa: Testa programas para DUMP na SE38.vbs
'Objetivo: Faz a leitura dos usuários de cada transação, uma por arquivo
'Data: 04/01/2019


'Carga Excel para SAP via Script
'Na transação xyz, pela transação mno buscar os users
'Alex Berloffe - 03/10/2018

'Em caso de erros, processa a próxima transação
' On Error Resume Next
' set wshell = createObject("WScript.Shell")

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

dim x,arrStr,ForReading
dim js, objTextFile
set js=CreateObject("Scripting.FileSystemObject")
set objTextFile = js.OpenTextFile("C:\guixt\scripts\UAT2_4JAN_PROGRAMS.txt")

Do while NOT objTextFile.AtEndOfStream
	arrStr         = Split(objTextFile.ReadLine,";")
	TCode	       = arrStr(0)

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "SE38"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").text = TCode
session.findById("wnd[0]/usr/ctxtRS38M-PROGRAMM").caretPosition = 16
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[0]/okcd").text = "/N"
session.findById("wnd[0]").sendVKey 0

Loop

objTextFile.Close
set objTextFile = Nothing
set js = Nothing
set x = 0
set arrStr = nothing

session.findById("wnd[0]/tbar[0]/btn[3]").press







