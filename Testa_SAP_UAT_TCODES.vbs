'Carga Excel para SAP via Script
'Esse script testará a abertura de todas as transações do UAT no ambiente LAB, para ver quais geram DUMP
'Alex Berloffe - 03/10/2018

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

dim x,arrStr,status,p
dim fs,objTextFile
set fs=CreateObject("Scripting.FileSystemObject")
set objTextFile = fs.OpenTextFile("C:\guixt\scripts\UAT_TCODES_PERU.csv")

Do while NOT objTextFile.AtEndOfStream
	arrStr         = Split(objTextFile.ReadLine,";")
	TC	       = arrStr(0)

'Abre a transação no SAP e fecha

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = TC
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
session.findById("wnd[0]").sendVKey 0

Loop

objTextFile.Close
set objTextFile = Nothing
set fs = Nothing

session.findById("wnd[0]/tbar[0]/btn[3]").press