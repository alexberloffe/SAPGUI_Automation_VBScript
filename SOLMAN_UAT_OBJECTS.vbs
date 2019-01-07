'Carga Excel para SAP SOLMAN via Script
'Esse script cria os objetos de testes do SOLMAN
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
set objTextFile = fs.OpenTextFile("C:\guixt\scripts\UAT_OBJECTS_6000 - Cópia.txt")

Do while NOT objTextFile.AtEndOfStream
	arrStr         = Split(objTextFile.ReadLine,";")
	c		= arrStr(0)
	Tipo		= arrStr(1)
	CLog		= arrStr(2)
	Objo		= arrStr(3)

c = CInt (c)

'Adiciona o nó da estrutura e soma 1 ao contador - se o contador é maior que 14, pula página

session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/cmbEL_LISTBOX_1[0,14]").key = "BMRE"
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,14]").text = Objo
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,14]").setFocus
session.findById("wnd[0]/usr/subSA_FRAMEWORK:SAPLSOLAR04:0100/tabsTS_MAIN/tabpTS_MAIN_FC5/ssubTS_MAIN_SCA:SAPLSOLAR05:0104/subSUB:SAPLSTC2:0147/tblSAPLSTC2GENERIC_TABLE_48/ctxtEL1[2,14]").caretPosition = 0
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "         " & c,"TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "         " & c,"TEXT"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").clickLink "         " & c,"TEXT"
session.findById("wnd[0]/tbar[0]/btn[11]").press


Loop

objTextFile.Close
set objTextFile = Nothing
set fs = Nothing

session.findById("wnd[0]/tbar[0]/btn[3]").press