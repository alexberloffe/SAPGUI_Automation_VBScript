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
set objTextFile = fs.OpenTextFile("C:\guixt\scripts\UAT_STRUCTURE.txt")

c = 0

Do while NOT objTextFile.AtEndOfStream
	arrStr         = Split(objTextFile.ReadLine,";")
	Structure      = arrStr(0)

'Adiciona o nó da estrutura e soma 1 ao contador - se o contador é maior que 14, pula página

session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC1/ssubTS_MAIN_SCA:SAPLSOLAR05:0101/subSUB:SAPLSTC2:0149/tblSAPLSTC2GENERIC_TABLE_50/ctxtEL4[0,14]").text = Structure
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC1/ssubTS_MAIN_SCA:SAPLSOLAR05:0101/subSUB:SAPLSTC2:0149/tblSAPLSTC2GENERIC_TABLE_50/cmbEL_LISTBOX_1[1,14]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC1/ssubTS_MAIN_SCA:SAPLSOLAR05:0101/subSUB:SAPLSTC2:0149/tblSAPLSTC2GENERIC_TABLE_50/cmbEL_LISTBOX_1[1,14]").key = "ZEHP7PRD"

c = c + 1

If c > 1 Then

session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC1/ssubTS_MAIN_SCA:SAPLSOLAR05:0101/subSUB:SAPLSTC2:0149/tblSAPLSTC2GENERIC_TABLE_50").verticalScrollbar.position = c

End If

Loop

objTextFile.Close
set objTextFile = Nothing
set fs = Nothing

session.findById("wnd[0]/tbar[0]/btn[3]").press